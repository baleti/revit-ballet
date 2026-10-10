using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using RevitBallet.Agents;

namespace RevitBallet.Commands
{
    /// <summary>
    /// The stable part of the agent server that stays in the originally loaded add-in assembly: it owns the Idling hook, loads the reloadable runtime
    /// and swaps it. Keep it small and free of agent logic. The runtime (RevitBallet.Commands.Agent.AgentBootstrap) comes from the SAME folder
    /// InvokeAddinCommand uses for hot reload, %APPDATA%\revit-ballet\hot-reload\&lt;year&gt;\revit-ballet.dll, and falls back to this assembly when that
    /// file does not exist. POST /v1/admin/reload (or the first Idling event after start) loads it; the previous runtime is detached afterwards.
    /// </summary>
    internal static class AgentShell
    {
        private static readonly object gate = new object();
        private static volatile bool reloadRequested;
        private static TaskCompletionSource<string> reloadResult;
        private static bool resolverInstalled;
        private static ExternalEvent controlEvent;   // reload requests: Revit's Idling event is not raised continuously, an ExternalEvent is serviced whenever Revit is idle

        private sealed class ControlHandler : IExternalEventHandler
        {
            public void Execute(UIApplication app) { try { StartOrReload(app); } catch (Exception ex) { RevitBalletServer.LogToRevit("[AGENT] reload failed: " + ex); } }
            public string GetName() => "RevitBalletAgentShellControl";
        }

        public static void Attach(UIControlledApplication app) { app.Idling += OnIdling; }

        public static void Detach(UIControlledApplication app)
        {
            try { app.Idling -= OnIdling; } catch { }
            var rt = AgentRuntimeRegistry.Current;
            AgentRuntimeRegistry.Current = null;
            // the application object for Detach is not at hand during shutdown: the runtime only needs it to unsubscribe, which Revit does at exit anyway
            try { if (rt != null) rt.Detach(null); } catch { }
        }

        // ---------------------------------------------------------------- UI thread (Idling): start, reload, and the second drain path for queued jobs

        private static void OnIdling(object sender, IdlingEventArgs e)
        {
            var uiapp = sender as UIApplication;
            if (uiapp == null) return;
            try
            {
                if (controlEvent == null) controlEvent = ExternalEvent.Create(new ControlHandler());   // created here: Idling is an API context
                if (AgentRuntimeRegistry.Current == null || reloadRequested) { StartOrReload(uiapp); return; }
                AgentRuntimeRegistry.Current.OnIdle(uiapp);
            }
            catch (Exception ex) { RevitBalletServer.LogToRevit("[AGENT] shell error: " + ex); }
        }

        private static void StartOrReload(UIApplication app)
        {
            TaskCompletionSource<string> tcs;
            lock (gate)
            {
                // Idling and the control event can both fire for one request: only the first one reloads (a second, unrequested reload would
                // detach the runtime that has just started, under a job that is being submitted to it)
                if (!reloadRequested && AgentRuntimeRegistry.Current != null) return;
                tcs = reloadResult; reloadRequested = false; reloadResult = null;
            }
            var sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                var old = AgentRuntimeRegistry.Current;
                string from;
                var asm = LoadRuntimeAssembly(app, out from);
                var boot = asm.GetType("RevitBallet.Commands.Agent.AgentBootstrap", true);
                IAgentRuntime fresh;
                try
                {
                    fresh = (IAgentRuntime)boot.GetMethod("Create").Invoke(null,
                        new object[] { app, Path.GetDirectoryName(typeof(AgentShell).Assembly.Location), from });
                }
                catch (TargetInvocationException tie) { throw tie.InnerException ?? tie; }
                AgentRuntimeRegistry.Current = fresh;       // new requests go to the new runtime from here on
                if (old != null) { try { old.Detach(app); } catch (Exception ex) { RevitBalletServer.LogToRevit("[AGENT] detaching the old runtime failed: " + ex.Message); } }
                var msg = $"{{\"reloaded\":true,\"build\":\"{fresh.Build}\",\"loadedFrom\":\"{fresh.LoadedFrom}\",\"previousBuild\":{(old == null ? "null" : "\"" + old.Build + "\"")},\"ms\":{sw.ElapsedMilliseconds}}}";
                if (tcs != null) tcs.TrySetResult(msg);
            }
            catch (Exception ex)
            {
                // the old runtime keeps serving: a bad build never takes the agent down
                RevitBalletServer.LogToRevit("[AGENT] reload failed: " + ex);
                if (tcs != null) tcs.TrySetException(ex);
            }
        }

        private static Assembly LoadRuntimeAssembly(UIApplication app, out string from)
        {
            var year = app.Application.VersionNumber;
            var dll = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "revit-ballet", "hot-reload", year, "revit-ballet.dll");
            if (!File.Exists(dll)) { from = "installed"; return typeof(AgentShell).Assembly; }
            InstallResolver();
            var bytes = File.ReadAllBytes(dll);   // from bytes: no file lock, like InvokeAddinCommand
            from = "hot-reload";
#if NET8_0_OR_GREATER
            // A second copy of revit-ballet.dll cannot live in Revit's add-in context ("Assembly with same name is already loaded"), so each
            // reload gets a context of its own that hands every dependency (contract, Newtonsoft, ...) to the shell's context.
            var ctx = new RuntimeContext(System.Runtime.Loader.AssemblyLoadContext.GetLoadContext(typeof(AgentShell).Assembly), Path.GetDirectoryName(dll));
            using (var ms = new MemoryStream(bytes)) return ctx.Adopt(ctx.LoadFromStream(ms));
#else
            return Assembly.Load(bytes);
#endif
        }

        /// <summary>
        /// The reloaded assembly must share the contract assembly (and Newtonsoft, Core, ...) with the shell, or the shell and the runtime would no longer agree
        /// on the type IAgentRuntime. Resolve every dependency to the copy that is already loaded.
        /// </summary>
        private static void InstallResolver()
        {
            lock (gate)
            {
                if (resolverInstalled) return;
                resolverInstalled = true;
#if NET8_0_OR_GREATER
                // nothing to install: RuntimeContext resolves the reloaded runtime's dependencies
#else
                AppDomain.CurrentDomain.AssemblyResolve += (s, args) =>
                {
                    var simple = new AssemblyName(args.Name).Name;
                    if (simple == "revit-ballet") return null;   // the reloaded copy itself is loaded by bytes
                    return AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => !a.IsDynamic && a.GetName().Name == simple && !string.IsNullOrEmpty(SafeLocation(a)));
                };
#endif
            }
        }
#if NET8_0_OR_GREATER
        private sealed class RuntimeContext : System.Runtime.Loader.AssemblyLoadContext
        {
            private readonly System.Runtime.Loader.AssemblyLoadContext shell; private readonly string dir; private Assembly self;
            public RuntimeContext(System.Runtime.Loader.AssemblyLoadContext shell, string dir) : base("revit-ballet-runtime", false) { this.shell = shell; this.dir = dir; }
            public Assembly Adopt(Assembly a) { self = a; return a; }
            protected override Assembly Load(AssemblyName name)
            {
                if (self != null && name.Name == self.GetName().Name) return self;
                // shared with the shell: whatever is already loaded under that name (contract assembly, Newtonsoft.Json, Revit API ...)
                var loaded = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => !a.IsDynamic && a.GetName().Name == name.Name && GetLoadContext(a) != this);
                if (loaded != null) return loaded;
                var file = Path.Combine(dir, name.Name + ".dll");
                if (File.Exists(file)) { using (var ms = new MemoryStream(File.ReadAllBytes(file))) return LoadFromStream(ms); }
                return null;
            }
        }
#endif

        private static string SafeLocation(Assembly a) { try { return a.Location; } catch { return null; } }

        // ---------------------------------------------------------------- HTTP side (any thread)

        public static async Task HandleV1Async(Stream stream, string method, string path, string body, string connId, string sessionId, int port, string client)
        {
            if (method == "POST" && path.TrimEnd('/').Equals("/v1/admin/reload", StringComparison.OrdinalIgnoreCase)) { await ReloadAsync(stream); return; }
            var rt = AgentRuntimeRegistry.Current;
            if (rt == null) { await WriteJson(stream, 503, "{\"error\":{\"kind\":\"starting\",\"message\":\"the agent runtime has not started yet (Revit has not been idle since it started)\"}}"); return; }
            await rt.HandleAsync(stream, method, path, body, connId, sessionId, port, client);
        }

        public static async Task<LegacyReply> HandleLegacyRoslynAsync(string code)
        {
            var rt = AgentRuntimeRegistry.Current;
            if (rt == null) return new LegacyReply { Status = 503, Json = "{\"Success\":false,\"Error\":\"the agent runtime has not started yet\",\"Diagnostics\":[],\"ProcessingLog\":[]}" };
            return await rt.HandleLegacyRoslynAsync(code);
        }

        private static async Task ReloadAsync(Stream stream)
        {
            TaskCompletionSource<string> tcs;
            lock (gate)
            {
                if (reloadResult == null) reloadResult = new TaskCompletionSource<string>();
                tcs = reloadResult; reloadRequested = true;
            }
            try { var ev = controlEvent; if (ev != null) ev.Raise(); } catch { }
            var done = await Task.WhenAny(tcs.Task, Task.Delay(30000));
            if (done != tcs.Task)
            {
                await WriteJson(stream, 504, "{\"error\":{\"kind\":\"reload_timeout\",\"message\":\"Revit did not become idle within 30 s (busy, or a modal dialog is open); the reload is still queued and will happen at the next idle moment\"}}");
                return;
            }
            if (tcs.Task.IsFaulted)
            {
                var ex = tcs.Task.Exception.GetBaseException();
                var msg = (ex.GetType().Name + ": " + ex.Message).Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", " ").Replace("\n", " ");
                await WriteJson(stream, 500, "{\"error\":{\"kind\":\"reload_failed\",\"message\":\"" + msg + " (the previous runtime keeps serving)\"}}");
                return;
            }
            await WriteJson(stream, 200, tcs.Task.Result);
        }

        private static async Task WriteJson(Stream stream, int status, string json)
        {
            var bytes = Encoding.UTF8.GetBytes(json);
            var reason = status == 200 ? "OK" : status == 503 ? "Service Unavailable" : status == 504 ? "Gateway Timeout" : "Error";
            var head = Encoding.ASCII.GetBytes($"HTTP/1.1 {status} {reason}\r\nContent-Type: application/json; charset=utf-8\r\nContent-Length: {bytes.Length}\r\nConnection: close\r\n\r\n");
            await stream.WriteAsync(head, 0, head.Length);
            await stream.WriteAsync(bytes, 0, bytes.Length);
            await stream.FlushAsync();
        }
    }
}
