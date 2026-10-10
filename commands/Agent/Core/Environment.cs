using System.Collections.Generic;
using RevitBallet.Agents;

namespace RevitBallet.Commands.Agent
{
    /// <summary>One engine found under engines/&lt;name&gt;/. Its assembly is loaded lazily, the first time a job asks for it.</summary>
    internal sealed class EngineEntry
    {
        public EngineManifest Manifest;
        public string Directory;
        public IScriptEngine Instance;
        public string LoadError;
        public readonly object LoadLock = new object();
    }

    /// <summary>How the core asks the UI thread to run queued jobs (in Revit: ExternalEvent.Raise). Must be callable from any thread.</summary>
    internal interface IUiBridge
    {
        /// <summary>Ask the UI thread to call <see cref="AgentCore.RunReadyJob"/> soon. Calls may coalesce.</summary>
        void RaiseDrain();
    }

    /// <summary>What the core needs to know about the process it runs in.</summary>
    internal interface IAgentEnvironment
    {
        bool UiResponsive { get; }
        /// <summary>Title of an open modal dialog, or null.</summary>
        string ModalDialogTitle { get; }
        int RevitYear { get; }
        string RevitVersion { get; }
        /// <summary>Folder for the per-job audit log; null disables it.</summary>
        string AuditDirectory { get; }
    }

    /// <summary>Where the engines come from (in Revit: manifests under engines/, loaded on first use).</summary>
    internal interface IEngineSource
    {
        IDictionary<string, EngineEntry> Engines { get; }
        /// <summary>Load the engine; returns null and sets <see cref="EngineEntry.LoadError"/> on failure.</summary>
        IScriptEngine Load(EngineEntry entry);
    }

    /// <summary>Log sink (in Revit: runtime\server.log). No-op by default so the core runs anywhere.</summary>
    internal static class AgentLog
    {
        public static System.Action<string> Sink = _ => { };
        public static void Write(string message) { try { Sink(message); } catch { } }
    }
}
