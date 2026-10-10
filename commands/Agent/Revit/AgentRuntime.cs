using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;
using RevitBallet.Agents;

namespace RevitBallet.Commands.Agent
{
    /// <summary>
    /// Entry point the stable shell (commands/AgentShell.cs) calls by reflection, in this assembly or in the hot-reload copy of it. Everything
    /// reachable from here can be replaced in a running Revit: the shell loads the new revit-ballet.dll and swaps the runtime.
    /// </summary>
    public static class AgentBootstrap
    {
        public static IAgentRuntime Create(object uiApplication, string addinDirectory, string loadedFrom)
        {
            return new AgentRuntime((UIApplication)uiApplication, addinDirectory, loadedFrom);
        }
    }

    /// <summary>Revit adapter around <see cref="AgentCore"/>: the ExternalEvent, the dialog events and the Revit-specific facts. Queue logic lives in Core/.</summary>
    internal sealed class AgentRuntime : IAgentRuntime
    {
        private readonly AgentCore core;
        private readonly ExternalEvent uiEvent;
        private readonly string build, loadedFrom;

        public string Build { get { return build; } }
        public string LoadedFrom { get { return loadedFrom; } }

        public AgentRuntime(UIApplication app, string addinDirectory, string loadedFrom)
        {
            this.loadedFrom = loadedFrom;
            var asm = typeof(AgentRuntime).Assembly;
            build = asm.GetName().Version + " " + asm.ManifestModule.ModuleVersionId.ToString("N").Substring(0, 12);
            EngineLoader.EnginesRootOverride = Path.Combine(addinDirectory, "engines");
            AgentLog.Sink = RevitBalletServer.LogToRevit;
            ValueConverter.Extensions.Add(RevitValues.TryToken);
            uiEvent = ExternalEvent.Create(new AgentExecHandler(this));
            core = new AgentCore(new ExternalEventBridge(uiEvent), new RevitEnvironment(), new EngineLoaderSource());
            core.AgentInfo = new JObject { ["build"] = build, ["loadedFrom"] = loadedFrom };
            RevitDialogs.Attach(app);
            core.Start();
            AgentLog.Write($"[AGENT] runtime {build} started (loaded from {loadedFrom})");
        }

        public Task HandleAsync(Stream stream, string method, string target, string body, string connId, string sessionId, int port, string client)
        {
            return AgentHttp.Handle(core, stream, method, target, body, connId, sessionId, port, client);
        }

        public Task<LegacyReply> HandleLegacyRoslynAsync(string code) { return LegacyRoslyn.HandleAsync(core, code); }

        public void OnIdle(object uiApplication)
        {
            if (core.HasReadyJob) RunReady((UIApplication)uiApplication);
        }

        internal void RunReady(UIApplication app)
        {
            var uidoc = app.ActiveUIDocument;
            core.RunReadyJob(app, uidoc, uidoc?.Document);
        }

        public void Detach(object uiApplication)
        {
            try { RevitDialogs.Detach((UIApplication)uiApplication); } catch { }
            core.Stop();
            try { uiEvent.Dispose(); } catch { }
            AgentLog.Write($"[AGENT] runtime {build} detached");
        }

        private sealed class ExternalEventBridge : IUiBridge
        {
            private readonly ExternalEvent e;
            public ExternalEventBridge(ExternalEvent e) { this.e = e; }
            public void RaiseDrain() { try { e.Raise(); } catch { } }
        }

        private sealed class RevitEnvironment : IAgentEnvironment
        {
            // From the Revit executable itself: independent of when the UIApplication becomes available. 24.x.y.z = Revit 2024.
            private readonly FileVersionInfo fvi = Process.GetCurrentProcess().MainModule.FileVersionInfo;
            public bool UiResponsive { get { return UiProbe.IsResponsive(); } }
            public DialogInfo CurrentDialog { get { return RevitDialogs.Current(); } }
            public bool TryRespondToDialog(string button, out string error) { return RevitDialogs.TryRespond(button, out error); }
            public int RevitYear { get { return 2000 + fvi.FileMajorPart; } }
            public string RevitVersion { get { return fvi.FileVersion; } }
            public System.Collections.Generic.IList<CommandInfo> Commands { get { return CommandCatalog.List(); } }
            public string AuditDirectory { get { return PathHelper.EnsureRuntimeSubdirectoryExists(Path.Combine("agent", "jobs")); } }
        }
    }

    internal sealed class AgentExecHandler : IExternalEventHandler
    {
        private readonly AgentRuntime runtime;
        public AgentExecHandler(AgentRuntime runtime) { this.runtime = runtime; }
        public void Execute(UIApplication app) { runtime.RunReady(app); }
        public string GetName() => "RevitBalletAgentExecHandler";
    }
}
