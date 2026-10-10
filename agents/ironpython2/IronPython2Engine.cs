using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using IronPython.Hosting;
using IronPython.Runtime.Exceptions;
using Microsoft.Scripting;
using Microsoft.Scripting.Hosting;
using RevitBallet.Agents;

namespace RevitBallet.Engines.IronPython2
{
    /// <summary>Exposed to scripts as <c>cancel</c>: <c>cancel.is_set()</c> mirrors threading.Event.</summary>
    public sealed class CancelFlag
    {
        private readonly CancellationToken token;
        public CancelFlag(CancellationToken t) { token = t; }
        public bool is_set() { return token.IsCancellationRequested; }
        public bool IsCancellationRequested { get { return token.IsCancellationRequested; } }
    }

    /// <summary>Exposed to scripts as <c>out</c>: Python-style <c>write</c>/<c>flush</c> plus .NET-style <c>WriteLine</c>.</summary>
    public sealed class PyOut
    {
        private readonly TextWriter writer;
        public PyOut(TextWriter w) { writer = w; }
        public void write(object text) { writer.Write(text == null ? "" : text.ToString()); }
        public void writelines(System.Collections.IEnumerable lines) { foreach (var l in lines) write(l); }
        public void flush() { }
        public void Write(object text) { write(text); }
        public void WriteLine(object text) { writer.WriteLine(text == null ? "" : text.ToString()); }
    }

    /// <summary>A write-only stream that decodes UTF-8 and forwards to a TextWriter (IronPython's sys.stdout wants a Stream).</summary>
    internal sealed class WriterStream : Stream
    {
        private readonly TextWriter writer;
        private readonly Decoder decoder = new UTF8Encoding(false).GetDecoder();
        public WriterStream(TextWriter w) { writer = w; }
        public override void Write(byte[] buffer, int offset, int count)
        {
            var chars = new char[count + 4];
            int n = decoder.GetChars(buffer, offset, count, chars, 0);
            if (n > 0) writer.Write(chars, 0, n);
        }
        public override bool CanRead { get { return false; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return true; } }
        public override long Length { get { return 0; } }
        public override long Position { get { return 0; } set { } }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) { return 0; }
        public override long Seek(long offset, SeekOrigin origin) { return 0; }
        public override void SetLength(long value) { }
    }

    internal sealed class Collector : ErrorListener
    {
        public readonly List<EngineDiagnostic> Items = new List<EngineDiagnostic>();
        public override void ErrorReported(ScriptSource source, string message, SourceSpan span, int errorCode, Severity severity)
        {
            Items.Add(new EngineDiagnostic
            {
                Severity = severity == Severity.Warning ? "warning" : "error",
                Id = errorCode.ToString(), Line = span.Start.Line, Column = span.Start.Column, Message = message,
            });
        }
    }

    public sealed class IronPython2Engine : IScriptEngine
    {
        public string Name { get { return "ironpython2"; } }
        public string Language { get { return "python"; } }
        public string Version { get { return typeof(Python).Assembly.GetName().Version.ToString(3); } }

        private ScriptEngine engine;
        private CancellationToken activeCancel;

        private ScriptEngine Engine()
        {
            if (engine != null) return engine;
            // Debug=true is what makes the trace hook (used for cancellation) fire; the cost is irrelevant for Revit-API-bound scripts.
            var e = Python.CreateEngine(new Dictionary<string, object> { { "Debug", true } });
            var dir = Path.GetDirectoryName(typeof(IronPython2Engine).Assembly.Location);
            var paths = new List<string>();
            var zip = Path.Combine(dir, "Lib.zip");
            if (File.Exists(zip)) paths.Add(zip);
            e.SetSearchPaths(paths);
            // `from Autodesk.Revit.DB import *` works without clr.AddReference
            e.Runtime.LoadAssembly(typeof(Autodesk.Revit.DB.Document).Assembly);
            e.Runtime.LoadAssembly(typeof(Autodesk.Revit.UI.UIDocument).Assembly);
            e.Runtime.LoadAssembly(typeof(System.Linq.Enumerable).Assembly);
            return engine = e;
        }

        public EngineCompileResult Compile(string code, EngineSession session, CancellationToken cancel)
        {
            var result = new EngineCompileResult();
            var listener = new Collector();
            try
            {
                var source = Engine().CreateScriptSourceFromString(code, SourceCodeKind.File);
                var compiled = source.Compile(listener);
                result.Success = compiled != null;
                result.Compiled = compiled;
            }
            catch (Exception ex)
            {
                result.Success = false;
                if (listener.Items.Count == 0)
                    listener.Items.Add(new EngineDiagnostic { Severity = "error", Id = ex.GetType().Name, Message = ex.Message });
            }
            // listener reports each syntax error twice (parser + compiler); keep distinct ones
            var seen = new HashSet<string>();
            foreach (var d in listener.Items)
                if (seen.Add(d.Severity + d.Line + ":" + d.Column + d.Message)) result.Diagnostics.Add(d);
            if (result.Success && result.Compiled == null) result.Success = false;
            return result;
        }

        public EngineRunResult Run(object compiled, EngineContext ctx)
        {
            var rr = new EngineRunResult();
            var e = Engine();
            var code = (CompiledCode)compiled;
            try
            {
                ScriptScope scope;
                if (ctx.Session != null)
                    scope = ctx.Session.State as ScriptScope ?? (ScriptScope)(ctx.Session.State = e.CreateScope());
                else
                    scope = e.CreateScope();

                scope.SetVariable("uiapp", ctx.UIApp); scope.SetVariable("UIApp", ctx.UIApp); scope.SetVariable("__revit__", ctx.UIApp);
                scope.SetVariable("uidoc", ctx.UIDoc); scope.SetVariable("UIDoc", ctx.UIDoc);
                scope.SetVariable("doc", ctx.Doc); scope.SetVariable("Doc", ctx.Doc);
                scope.SetVariable("cancel", new CancelFlag(ctx.Cancel));
                scope.SetVariable("out", new PyOut(ctx.Out));
                object ignored; if (scope.TryGetVariable("result", out ignored)) scope.RemoveVariable("result");

                var utf8 = new UTF8Encoding(false);
                e.Runtime.IO.SetOutput(new WriterStream(ctx.Out), utf8);
                e.Runtime.IO.SetErrorOutput(new WriterStream(ctx.Out), utf8);

                activeCancel = ctx.Cancel;
                TracebackDelegate tracer = null;
                tracer = (frame, kind, payload) =>
                {
                    if (activeCancel.IsCancellationRequested) throw new OperationCanceledException("cancelled by host");
                    return tracer;
                };
                Python.SetTrace(e, tracer);
                try { code.Execute(scope); }
                finally { Python.SetTrace(e, null); }

                object value; scope.TryGetVariable("result", out value);
                rr.Ok = true; rr.Value = value;
            }
            catch (OperationCanceledException ex) { rr.ErrorKind = "cancelled"; rr.ErrorMessage = ex.Message; }
            catch (Exception ex)
            {
                if (ctx.Cancel.IsCancellationRequested && ex is System.Threading.ThreadAbortException) { rr.ErrorKind = "cancelled"; rr.ErrorMessage = "cancelled"; return rr; }
                rr.ErrorKind = "runtime";
                try
                {
                    var eo = e.GetService<ExceptionOperations>();
                    var full = eo.FormatException(ex);
                    rr.ErrorMessage = FirstLineOf(full);
                    rr.ErrorStack = full;
                }
                catch { rr.ErrorMessage = ex.GetType().Name + ": " + ex.Message; rr.ErrorStack = ex.StackTrace; }
            }
            return rr;
        }

        // FormatException gives "Traceback (...)\n  File ..., line n, in f\nValueError: bad": the last line is the readable message.
        private static string FirstLineOf(string formatted)
        {
            var lines = formatted.Replace("\r", "").Split('\n');
            for (int i = lines.Length - 1; i >= 0; i--) if (lines[i].Trim().Length > 0) return lines[i].Trim();
            return formatted;
        }

        public void ResetSession(EngineSession session) { if (session != null) session.State = null; }
        public void Abort(EngineContext context) { /* cooperative: the trace hook raises as soon as the cancellation token is set */ }
    }
}
