using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using RevitBallet.Agents;

namespace RevitBallet.Engines.Roslyn
{
    /// <summary>Globals visible to C# scripts.</summary>
    public class Globals
    {
        public UIApplication UIApp { get; set; }
        public UIDocument UIDoc { get; set; }
        public Document Doc { get; set; }
        public TextWriter Out { get; set; }
        public CancellationToken Cancel { get; set; }
    }

    internal sealed class SessionState
    {
        public ScriptState<object> State;
        public Globals Globals = new Globals();
    }

    internal sealed class Compiled
    {
        public Script<object> Script;
        public SessionState Session;   // null for isolated jobs
        public Globals Globals;
    }

    public sealed class RoslynEngine : IScriptEngine
    {
        public string Name => "roslyn";
        public string Language => "csharp";
        public string Version => typeof(CSharpScript).Assembly.GetName().Version.ToString(3);

        private ScriptOptions options;

        /// <summary>
        /// One assembly loader for every script. Without it Roslyn creates a new assembly load context per script; in Revit that cost
        /// about 400 MB of native memory per compile (invisible to the garbage collector, so the working set grew to tens of GB).
        /// </summary>
        private readonly Microsoft.CodeAnalysis.Scripting.Hosting.InteractiveAssemblyLoader loader =
            new Microsoft.CodeAnalysis.Scripting.Hosting.InteractiveAssemblyLoader();

        private ScriptOptions Options()
        {
            if (options != null) return options;
            // Explicit file references from known locations: avoids picking up incompatible assemblies loaded by other add-ins (e.g. pyRevit).
            string engineDir = Path.GetDirectoryName(typeof(RoslynEngine).Assembly.Location);
            var o = ScriptOptions.Default
                .AddReferences(MetadataReference.CreateFromFile(typeof(Globals).Assembly.Location))
                .AddReferences(MetadataReference.CreateFromFile(typeof(Document).Assembly.Location))
                .AddReferences(MetadataReference.CreateFromFile(typeof(UIDocument).Assembly.Location))
                .AddReferences(MetadataReference.CreateFromFile(typeof(System.Linq.Enumerable).Assembly.Location));
            foreach (var dll in new[] { "System.Collections.Immutable.dll", "System.Reflection.Metadata.dll" })
            {
                var p = Path.Combine(engineDir, dll);
                if (File.Exists(p)) o = o.AddReferences(MetadataReference.CreateFromFile(p));
            }
            // The host add-in: lets scripts call its helpers (SelectionStorage, PathHelper, ...) without qualification.
            var host = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.GetName().Name == "revit-ballet");
            if (host != null && !string.IsNullOrEmpty(host.Location)) o = o.AddReferences(MetadataReference.CreateFromFile(host.Location));
            o = o.AddImports("System", "System.IO", "System.Linq", "System.Collections.Generic", "Autodesk.Revit.DB", "Autodesk.Revit.UI");
            if (host != null) o = o.AddImports("RevitBallet.Commands");
            return options = o;
        }

        // A script compile leaves a few hundred MB of native memory behind (Roslyn's metadata for the Revit API) that only a garbage collection
        // releases, and the runtime cannot see it. Left alone it piles up for as long as RAM is plentiful, so collect every few compiles.
        // The collection runs on its own thread and never waits for finalizers: a blocking GC.WaitForPendingFinalizers() on the job worker thread
        // hung a Revit session for minutes (a finalizer of a Revit object can need the UI thread), and every job queued behind it.
        private static int compileCount;
        private const int CollectEvery = 4;

        public EngineCompileResult Compile(string code, EngineSession session, CancellationToken cancel)
        {
            try { return CompileCore(code, session, cancel); }
            finally
            {
                if (Interlocked.Increment(ref compileCount) % CollectEvery == 0)
                    ThreadPool.QueueUserWorkItem(_ => { try { GC.Collect(2, GCCollectionMode.Forced, true); } catch { } });
            }
        }

        private EngineCompileResult CompileCore(string code, EngineSession session, CancellationToken cancel)
        {
            var result = new EngineCompileResult();
            try
            {
                var state = session != null && session.Name != null ? (session.State as SessionState ?? (SessionState)(session.State = new SessionState())) : null;
                Script<object> script = state != null && state.State != null
                    ? state.State.Script.ContinueWith(code, Options())
                    : CSharpScript.Create(code, Options(), typeof(Globals), loader);
                var diags = script.GetCompilation().GetDiagnostics(cancel);
                foreach (var d in diags.Where(d => d.Severity >= DiagnosticSeverity.Warning))
                {
                    var pos = d.Location.GetLineSpan().StartLinePosition;
                    result.Diagnostics.Add(new EngineDiagnostic
                    {
                        Severity = d.Severity == DiagnosticSeverity.Error ? "error" : "warning",
                        Id = d.Id, Line = pos.Line + 1, Column = pos.Character + 1, Message = d.GetMessage(),
                    });
                }
                result.Success = !diags.Any(d => d.Severity == DiagnosticSeverity.Error);
                if (result.Success) result.Compiled = new Compiled { Script = script, Session = state, Globals = state != null ? state.Globals : new Globals() };
            }
            catch (OperationCanceledException) { result.Success = false; result.Diagnostics.Add(new EngineDiagnostic { Severity = "error", Id = "cancelled", Message = "cancelled" }); }
            catch (CompilationErrorException ex)
            {
                result.Success = false;
                foreach (var d in ex.Diagnostics) result.Diagnostics.Add(new EngineDiagnostic { Severity = "error", Id = d.Id, Message = d.GetMessage() });
            }
            return result;
        }

        public EngineRunResult Run(object compiled, EngineContext ctx)
        {
            var c = (Compiled)compiled;
            var rr = new EngineRunResult();
            try
            {
                var g = c.Globals;
                g.UIApp = (UIApplication)ctx.UIApp; g.UIDoc = (UIDocument)ctx.UIDoc; g.Doc = (Document)ctx.Doc;
                g.Out = ctx.Out; g.Cancel = ctx.Cancel;
                ScriptState<object> state = c.Session != null && c.Session.State != null
                    ? c.Script.RunFromAsync(c.Session.State, e => false, ctx.Cancel).GetAwaiter().GetResult()
                    : c.Script.RunAsync(g, e => false, ctx.Cancel).GetAwaiter().GetResult();
                if (c.Session != null) c.Session.State = state;
                if (state.Exception != null) throw state.Exception;
                rr.Ok = true; rr.Value = state.ReturnValue;
            }
            catch (OperationCanceledException ex) { rr.ErrorKind = "cancelled"; rr.ErrorMessage = ex.Message; }
            catch (CompilationErrorException ex)
            {
                rr.ErrorKind = "runtime"; rr.ErrorMessage = "compilation error at run time";
                foreach (var d in ex.Diagnostics) rr.Diagnostics.Add(new EngineDiagnostic { Severity = "error", Id = d.Id, Message = d.GetMessage() });
            }
            catch (Exception ex)
            {
                rr.ErrorKind = "runtime"; rr.ErrorMessage = ex.GetType().Name + ": " + ex.Message; rr.ErrorStack = ex.StackTrace;
            }
            return rr;
        }

        public void ResetSession(EngineSession session) { if (session != null) session.State = null; }

        public void Abort(EngineContext context) { /* Roslyn scripts can only stop cooperatively through the Cancel token */ }
    }
}
