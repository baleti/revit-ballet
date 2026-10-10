using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using RevitBallet.Agents;
using RevitBallet.Commands.Agent;
using Xunit;

[assembly: CollectionBehavior(DisableTestParallelization = true)]

namespace RevitBallet.Agent.Tests
{
    /// <summary>
    /// A scripted engine. The first word of the "code" is the instruction:
    ///   ok [value]        return value (a long when numeric)       sleep N     block N ms, ignoring cancellation
    ///   spin              loop until cancelled                      throw MSG   runtime error
    ///   compile-error     fail compilation (line 1)                 compile-slow N   compile takes N ms
    ///   print N           N lines of output                         bigout KB   KB kilobytes of output
    ///   console TEXT      Console.WriteLine(TEXT) on the run thread value-list N   return N ints
    ///   value-big N       return a string of N chars                set K V / get K   session state
    ///   ctx               return what the engine was given          hold        compile result keeps a Big object alive (leak test)
    /// </summary>
    internal sealed class FakeEngine : IScriptEngine
    {
        public string Name { get; }
        public string Language => "fake";
        public string Version => "1.0";
        public int Runs;
        public readonly List<string> RunLog = new List<string>();
        public readonly List<WeakReference> Held = new List<WeakReference>();

        public FakeEngine(string name) { Name = name; }

        public EngineCompileResult Compile(string code, EngineSession session, CancellationToken cancel)
        {
            var r = new EngineCompileResult();
            var parts = code.Split(new[] { ' ' }, 2);
            if (parts[0] == "compile-slow") { var end = DateTime.UtcNow.AddMilliseconds(int.Parse(parts[1])); while (DateTime.UtcNow < end && !cancel.IsCancellationRequested) Thread.Sleep(10); }
            if (parts[0] == "compile-error")
            {
                r.Diagnostics.Add(new EngineDiagnostic { Severity = "error", Id = "E1", Line = 1, Column = 2, Message = "bad code" });
                return r;
            }
            object compiled = code;
            if (parts[0] == "hold") { var big = new Big(); lock (Held) Held.Add(new WeakReference(big)); compiled = big; }
            r.Success = true; r.Compiled = compiled;
            return r;
        }

        public EngineRunResult Run(object compiled, EngineContext ctx)
        {
            var rr = new EngineRunResult();
            var code = compiled as string ?? "hold";
            var parts = code.Split(new[] { ' ' }, 2);
            var arg = parts.Length > 1 ? parts[1] : "";
            lock (RunLog) RunLog.Add(code);
            Interlocked.Increment(ref Runs);
            switch (parts[0])
            {
                case "ok": rr.Ok = true; long n; rr.Value = long.TryParse(arg, out n) ? (object)n : arg; break;
                case "sleep": Thread.Sleep(int.Parse(arg)); rr.Ok = true; break;
                case "spin":
                    while (!ctx.Cancel.IsCancellationRequested) Thread.Sleep(5);
                    rr.ErrorKind = "cancelled"; rr.ErrorMessage = "stopped by the cancellation token"; break;
                case "throw": rr.ErrorKind = "runtime"; rr.ErrorMessage = arg; rr.ErrorStack = "at fake"; break;
                case "print": for (int i = 0; i < int.Parse(arg); i++) ctx.Out.WriteLine("line " + i); rr.Ok = true; break;
                case "bigout": { var chunk = new string('x', 1023) + "\n"; for (int i = 0; i < int.Parse(arg); i++) ctx.Out.Write(chunk); rr.Ok = true; break; }
                case "console": Console.WriteLine(arg); rr.Ok = true; break;
                case "value-list": rr.Ok = true; rr.Value = Enumerable.Range(0, int.Parse(arg)).ToList(); break;
                case "value-big": rr.Ok = true; rr.Value = new string('v', int.Parse(arg)); break;
                case "set": { var kv = arg.Split(' '); var d = Dict(ctx); d[kv[0]] = kv[1]; rr.Ok = true; break; }
                case "get": { var d = Dict(ctx); rr.Ok = true; rr.Value = d.TryGetValue(arg, out var v) ? v : "(unset)"; break; }
                case "ctx": rr.Ok = true; rr.Value = $"{ctx.UIApp}/{ctx.UIDoc}/{ctx.Doc}/session={(ctx.Session == null ? "none" : ctx.Session.Name)}"; break;
                case "hold": rr.Ok = true; rr.Value = 1L; break;
                default: rr.Ok = true; rr.Value = code; break;
            }
            return rr;
        }

        private static Dictionary<string, string> Dict(EngineContext ctx)
        {
            if (ctx.Session == null) return new Dictionary<string, string>();
            return ctx.Session.State as Dictionary<string, string> ?? (Dictionary<string, string>)(ctx.Session.State = new Dictionary<string, string>());
        }

        public void ResetSession(EngineSession session) { if (session != null) session.State = null; }
        public void Abort(EngineContext context) { }
        public sealed class Big { public byte[] Payload = new byte[1024 * 1024]; }
    }

    /// <summary>A pretend Revit UI thread: runs RunReadyJob when raised, can be blocked like a modal dialog.</summary>
    internal sealed class FakeUi : IUiBridge, IAgentEnvironment, IDisposable
    {
        private readonly AutoResetEvent signal = new AutoResetEvent(false);
        private Thread thread; private volatile bool stop;
        public AgentCore Core;
        public volatile bool Blocked;
        public volatile bool Hung;
        public int Raises;
        public string AuditDirectory { get; set; }

        public void Attach(AgentCore core)
        {
            Core = core;
            thread = new Thread(Loop) { IsBackground = true, Name = "fake-ui" };
            thread.Start();
        }
        private void Loop()
        {
            while (!stop)
            {
                signal.WaitOne(20);
                if (stop) return;
                if (Blocked) continue;   // a modal dialog: the external event is not serviced
                if (Core != null && Core.HasReadyJob) Core.RunReadyJob("uiapp", "uidoc", "doc");
            }
        }
        public void RaiseDrain() { Interlocked.Increment(ref Raises); signal.Set(); }

        public bool UiResponsive => !Hung;
        public string ModalDialogTitle => Blocked ? "Fake modal dialog" : null;
        public int RevitYear => 2026;
        public string RevitVersion => "26.0.0-fake";
        public void Dispose() { stop = true; signal.Set(); }
    }

    internal sealed class FakeSource : IEngineSource
    {
        public readonly Dictionary<string, EngineEntry> Map = new Dictionary<string, EngineEntry>(StringComparer.OrdinalIgnoreCase);
        public IDictionary<string, EngineEntry> Engines => Map;
        public FakeEngine Add(string name)
        {
            var e = new FakeEngine(name);
            Map[name] = new EngineEntry { Manifest = new EngineManifest { Name = name, Language = "fake", Version = "1", Protocol = 1, Sessions = true }, Directory = "x", Instance = e };
            return e;
        }
        public IScriptEngine Load(EngineEntry entry) { return entry.LoadError != null ? null : entry.Instance; }
    }

    /// <summary>A core wired to a fake UI thread and fake engines "roslyn" and "ironpython2" (cpython3 is "not installed", "broken" fails to load).</summary>
    internal sealed class Rig : IDisposable
    {
        public readonly FakeUi Ui = new FakeUi();
        public readonly FakeSource Source = new FakeSource();
        public readonly FakeEngine Roslyn, Python;
        public readonly AgentCore Core;
        public readonly string AuditDir = Path.Combine(Path.GetTempPath(), "rb-agent-test-" + Guid.NewGuid().ToString("N"));

        public Rig(bool audit = false)
        {
            Roslyn = Source.Add("roslyn"); Python = Source.Add("ironpython2");
            Source.Map["broken"] = new EngineEntry { Manifest = new EngineManifest { Name = "broken", Language = "x", Version = "1", Protocol = 1 }, Directory = "x", LoadError = "boom while loading" };
            Ui.AuditDirectory = audit ? AuditDir : null;
            Core = new AgentCore(Ui, Ui, Source);
            Ui.Attach(Core);
            Core.Start();
        }

        public Job Run(string code, string engine = "roslyn", int timeoutMs = 10000, string session = null, string client = null, string requestId = null, int waitMs = 15000)
        {
            var job = Core.Submit(engine, code, timeoutMs, session, client, requestId);
            Assert.True(Core.WaitFinal(job, waitMs), $"job {job.Id} ({code}) did not finish within {waitMs} ms (state {job.State})");
            return job;
        }

        public void Dispose()
        {
            Core.Stop(); Ui.Dispose();
            try { if (Directory.Exists(AuditDir)) Directory.Delete(AuditDir, true); } catch { }
        }

        /// <summary>Call AgentHttp the way the server does, over in-memory streams.</summary>
        public (int Status, string Body) Http(string method, string target, string body = null, string client = null)
        {
            var ms = new MemoryStream();
            AgentHttp.Handle(Core, ms, method, target, body ?? "", "c1", "4242", 23717, client).GetAwaiter().GetResult();
            var text = Encoding.UTF8.GetString(ms.ToArray());
            var split = text.IndexOf("\r\n\r\n", StringComparison.Ordinal);
            var status = int.Parse(text.Substring(9, 3));
            return (status, split < 0 ? "" : text.Substring(split + 4));
        }
    }

    internal static class Wait
    {
        public static bool Until(Func<bool> cond, int ms = 5000) { var end = DateTime.UtcNow.AddMilliseconds(ms); while (DateTime.UtcNow < end) { if (cond()) return true; Thread.Sleep(10); } return cond(); }
    }
}
