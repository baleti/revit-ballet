using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit.UI;
using RevitBallet.Agents;

namespace RevitBallet.Commands.Agent
{
    /// <summary>
    /// Revit adapter around <see cref="AgentCore"/>: owns the ExternalEvent, the Idling drain path and the Revit-specific facts. The queue and
    /// everything else is in Core/ (testable without Revit).
    /// </summary>
    internal static class AgentHost
    {
        public static AgentCore Core { get; private set; }
        private static ExternalEvent uiEvent;
        private static UIControlledApplication idlingApp;

        /// <summary>Must be called from a Revit API context (ExternalEvent.Create requirement).</summary>
        public static void Initialize()
        {
            if (Core != null) return;
            AgentLog.Sink = RevitBalletServer.LogToRevit;
            ValueConverter.Extensions.Add(RevitValues.TryToken);
            uiEvent = ExternalEvent.Create(new AgentExecHandler());
            var core = new AgentCore(new ExternalEventBridge(), new RevitEnvironment(), new EngineLoaderSource());
            core.Start();
            Core = core;
        }

        // ---- thin facade used by Server.cs
        public static Job Submit(string engine, string code, int timeoutMs, string session, string client = null, string requestId = null)
        { return Core.Submit(engine, code, timeoutMs, session, client, requestId); }

        public static bool WaitFinal(Job job, int waitMs) { return Core.WaitFinal(job, waitMs); }

        // ---- Revit glue
        private sealed class ExternalEventBridge : IUiBridge
        {
            public void RaiseDrain() { var e = uiEvent; if (e != null) e.Raise(); }
        }

        private sealed class RevitEnvironment : IAgentEnvironment
        {
            // From the Revit executable itself: independent of when the UIApplication becomes available. 24.x.y.z = Revit 2024.
            private readonly FileVersionInfo fvi = Process.GetCurrentProcess().MainModule.FileVersionInfo;
            public bool UiResponsive { get { return UiProbe.IsResponsive(); } }
            public string ModalDialogTitle { get { return UiProbe.ModalDialogTitle(); } }
            public int RevitYear { get { return 2000 + fvi.FileMajorPart; } }
            public string RevitVersion { get { return fvi.FileVersion; } }
            public string AuditDirectory { get { return PathHelper.EnsureRuntimeSubdirectoryExists(Path.Combine("agent", "jobs")); } }
        }

        internal static void RunReady(UIApplication app)
        {
            var uidoc = app.ActiveUIDocument;
            Core.RunReadyJob(app, uidoc, uidoc?.Document);
        }

        // Second drain path: Revit raises Idling when it is idle, which is more dependable than ExternalEvent while Revit is in the
        // background. Subscribed once at start-up (valid API context); the handler returns at once when nothing is waiting.
        public static void AttachIdling(UIControlledApplication app)
        {
            if (idlingApp != null) return;
            idlingApp = app; app.Idling += OnIdling;
        }

        public static void DetachIdling()
        {
            if (idlingApp == null) return;
            try { idlingApp.Idling -= OnIdling; } catch { }
            idlingApp = null;
        }

        private static void OnIdling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {
            var core = Core;
            if (core == null || !core.HasReadyJob) return;
            var uiapp = sender as UIApplication;
            if (uiapp != null) RunReady(uiapp);
        }
    }

    internal sealed class AgentExecHandler : IExternalEventHandler
    {
        public void Execute(UIApplication app) { AgentHost.RunReady(app); }
        public string GetName() => "RevitBalletAgentExecHandler";
    }
}
