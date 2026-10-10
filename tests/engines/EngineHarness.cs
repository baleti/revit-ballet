using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;
using System.Text;
using System.Threading;
using Newtonsoft.Json;
using RevitBallet.Agents;
using Xunit;

[assembly: CollectionBehavior(DisableTestParallelization = true)]

namespace RevitBallet.Engines.Tests
{
    /// <summary>What a job does, minus the HTTP and queue: compile on one thread, run on the calling ("UI") thread, like the host.</summary>
    internal sealed class Outcome
    {
        public EngineCompileResult Compile;
        public EngineRunResult Run;
        public string Output;
        public bool Ok => Compile.Success && Run != null && Run.Ok;
        public object Value => Run?.Value;
        public EngineDiagnostic FirstError => Compile.Diagnostics.FirstOrDefault(d => d.Severity == "error");
    }

    internal static class EngineHarness
    {
        private static readonly Dictionary<string, IScriptEngine> engines = new Dictionary<string, IScriptEngine>();
        private static readonly List<string> dirs = new List<string>();
        private static bool resolver;

        public static string RepoRoot => Meta("RepoRoot");
        public static string Year => Meta("EngineYear");
        private static string Meta(string key) =>
            typeof(EngineHarness).Assembly.GetCustomAttributes<AssemblyMetadataAttribute>().First(a => a.Key == key).Value;

        public static string EngineDir(string name) => Path.Combine(RepoRoot, "agents", name, "bin", Year);

        public static bool CPythonRuntimeAvailable =>
            Environment.OSVersion.Platform == PlatformID.Win32NT && Directory.GetFiles(EngineDir("cpython3"), "python3*.dll").Any(f => System.Text.RegularExpressions.Regex.IsMatch(Path.GetFileName(f), @"^python3\d+\.dll$"));

        public static EngineManifest Manifest(string name) =>
            JsonConvert.DeserializeObject<EngineManifest>(File.ReadAllText(Path.Combine(EngineDir(name), "engine.json")));

        /// <summary>Loads the engine assembly named in the manifest into the default context, with its own folder on the probing path (as the add-in does).</summary>
        public static IScriptEngine Load(string name)
        {
            lock (engines)
            {
                IScriptEngine e;
                if (engines.TryGetValue(name, out e)) return e;
                var dir = EngineDir(name);
                var m = Manifest(name);
                if (!resolver)
                {
                    resolver = true;
                    AssemblyLoadContext.Default.Resolving += (ctx, an) =>
                    {
                        foreach (var d in dirs.ToArray()) { var p = Path.Combine(d, an.Name + ".dll"); if (File.Exists(p)) return ctx.LoadFromAssemblyPath(p); }
                        return null;
                    };
                }
                dirs.Add(dir);
                var asm = AssemblyLoadContext.Default.LoadFromAssemblyPath(Path.GetFullPath(Path.Combine(dir, m.Assembly)));
                e = (IScriptEngine)Activator.CreateInstance(asm.GetType(m.Type, true));
                return engines[name] = e;
            }
        }

        /// <summary>One job. cancelAfterMs &gt; 0 cancels the token (and calls Abort, like the host's timeout path) from another thread.</summary>
        public static Outcome Run(IScriptEngine engine, string code, EngineSession session = null, int cancelAfterMs = 0, int hardLimitMs = 30000)
        {
            var o = new Outcome();
            var cts = new CancellationTokenSource();
            var outp = new StringWriter();
            o.Compile = engine.Compile(code, session ?? new EngineSession(), cts.Token);
            if (!o.Compile.Success) { o.Output = ""; return o; }
            var ctx = new EngineContext { UIApp = new Autodesk.Revit.UI.UIApplication(), UIDoc = new Autodesk.Revit.UI.UIDocument(), Doc = new Autodesk.Revit.DB.Document(), Out = outp, Cancel = cts.Token, Session = session };
            if (cancelAfterMs > 0) new Thread(() => { Thread.Sleep(cancelAfterMs); cts.Cancel(); try { engine.Abort(ctx); } catch { } }) { IsBackground = true }.Start();
            var runner = new Thread(() => o.Run = engine.Run(o.Compile.Compiled, ctx)) { IsBackground = true };
            runner.Start();
            Assert.True(runner.Join(hardLimitMs), "the script did not stop within " + hardLimitMs + " ms");
            o.Output = outp.ToString();
            return o;
        }
    }
}
