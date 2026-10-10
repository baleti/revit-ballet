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

    /// <summary>A modal dialog that is holding up the UI thread.</summary>
    internal sealed class DialogInfo
    {
        public string Title;
        /// <summary>"task" (Revit TaskDialog), "win32" (any other modal window) or "unknown".</summary>
        public string Kind = "unknown";
        public string Text;
        public List<string> Buttons = new List<string>();
        /// <summary>Revit's TaskDialog id when known (DialogBoxShowing), else null.</summary>
        public string DialogId;

        /// <summary>Same dialog as before? (title, id and the start of the text)</summary>
        public string Key { get { return Title + "|" + DialogId + "|" + (Text == null ? "" : Text.Substring(0, System.Math.Min(80, Text.Length))); } }

        public Newtonsoft.Json.Linq.JObject ToJson(double? sinceMs = null)
        {
            var o = new Newtonsoft.Json.Linq.JObject
            {
                ["title"] = Title, ["kind"] = Kind, ["text"] = Text, ["dialogId"] = DialogId,
                ["buttons"] = new Newtonsoft.Json.Linq.JArray(Buttons.ToArray()),
            };
            if (sinceMs.HasValue) o["sinceMs"] = System.Math.Round(sinceMs.Value);
            return o;
        }
    }

    /// <summary>One invokable add-in command (an IExternalCommand class of revit-ballet), as listed by GET /v1/commands.</summary>
    internal sealed class CommandInfo
    {
        public string Name;
        public string FullName;
        public string DisplayName;
        /// <summary>What the user must have ready before running it (the CommandMeta attribute): "", "Any", "Door", ...</summary>
        public string Input = "";
        public string Output = "";

        public Newtonsoft.Json.Linq.JObject ToJson()
        {
            return new Newtonsoft.Json.Linq.JObject { ["name"] = Name, ["displayName"] = DisplayName, ["input"] = Input ?? "", ["output"] = Output ?? "" };
        }
    }

    /// <summary>What the core needs to know about the process it runs in.</summary>
    internal interface IAgentEnvironment
    {
        bool UiResponsive { get; }
        /// <summary>The open modal dialog (what it says, its buttons), or null. Must be safe to call from any thread, also while the UI thread is blocked.</summary>
        DialogInfo CurrentDialog { get; }
        /// <summary>Press a button of the open dialog (matched by caption, case-insensitive). Returns false and an error text when that is not possible.</summary>
        bool TryRespondToDialog(string button, out string error);
        int RevitYear { get; }
        string RevitVersion { get; }
        /// <summary>Folder for the per-job audit log; null disables it.</summary>
        string AuditDirectory { get; }
        /// <summary>The commands that can be run through the "command" engine (empty when the host has none).</summary>
        System.Collections.Generic.IList<CommandInfo> Commands { get; }
    }

    /// <summary>Where the engines come from (in Revit: manifests under engines/, loaded on first use).</summary>
    internal interface IEngineSource
    {
        IDictionary<string, EngineEntry> Engines { get; }
        /// <summary>Load the engine; returns null and sets <see cref="EngineEntry.LoadError"/> on failure.</summary>
        IScriptEngine Load(EngineEntry entry);
        /// <summary>Forget the loaded instance so the next job gets a fresh one (the watchdog calls this after a hung compile).</summary>
        void Reset(EngineEntry entry);
    }

    /// <summary>Log sink (in Revit: runtime\server.log). No-op by default so the core runs anywhere.</summary>
    internal static class AgentLog
    {
        public static System.Action<string> Sink = _ => { };
        public static void Write(string message) { try { Sink(message); } catch { } }
    }
}
