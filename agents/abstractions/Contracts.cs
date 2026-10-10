using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace RevitBallet.Agents
{
    /// <summary>One script language runtime. Implemented by the assemblies under engines/&lt;name&gt;/. See docs/agent-protocol.md.</summary>
    public interface IScriptEngine
    {
        string Name { get; }
        string Language { get; }
        string Version { get; }

        /// <summary>Parse/compile on a worker thread. Must not touch the Revit API. Never throws for bad user code: report it as diagnostics.</summary>
        EngineCompileResult Compile(string code, EngineSession session, CancellationToken cancel);

        /// <summary>Run on the Revit UI thread. Never throws for errors in user code: report them in the result.</summary>
        EngineRunResult Run(object compiled, EngineContext context);

        /// <summary>Called when the user drops a session, and by the host when a session is replaced.</summary>
        void ResetSession(EngineSession session);

        /// <summary>Best effort: ask a running script to stop (the cancellation token is already set).</summary>
        void Abort(EngineContext context);
    }

    /// <summary>State kept between jobs of one named session. Engines store whatever they need in <see cref="State"/>.</summary>
    public sealed class EngineSession
    {
        public string Name { get; set; }
        public object State { get; set; }
        public int Jobs { get; set; }
        public DateTime Created { get; } = DateTime.UtcNow;
    }

    public sealed class EngineContext
    {
        /// <summary>Autodesk.Revit.UI.UIApplication</summary>
        public object UIApp { get; set; }
        /// <summary>Autodesk.Revit.UI.UIDocument or null</summary>
        public object UIDoc { get; set; }
        /// <summary>Autodesk.Revit.DB.Document or null</summary>
        public object Doc { get; set; }
        /// <summary>Everything the script prints goes here (per job).</summary>
        public TextWriter Out { get; set; }
        public CancellationToken Cancel { get; set; }
        /// <summary>The named session, or null for an isolated job.</summary>
        public EngineSession Session { get; set; }
    }

    public sealed class EngineDiagnostic
    {
        public string Severity { get; set; }   // "error" | "warning" | "info"
        public string Id { get; set; }
        public int Line { get; set; }          // 1-based, 0 = unknown
        public int Column { get; set; }
        public string Message { get; set; }
    }

    public sealed class EngineCompileResult
    {
        public bool Success { get; set; }
        /// <summary>Opaque; handed back to <see cref="IScriptEngine.Run"/>.</summary>
        public object Compiled { get; set; }
        public List<EngineDiagnostic> Diagnostics { get; } = new List<EngineDiagnostic>();
    }

    public sealed class EngineRunResult
    {
        public bool Ok { get; set; }
        /// <summary>When not Ok: "runtime", "cancelled" or "engine".</summary>
        public string ErrorKind { get; set; }
        public string ErrorMessage { get; set; }
        public string ErrorStack { get; set; }
        /// <summary>The script's result value (host converts it to JSON/text).</summary>
        public object Value { get; set; }
        public List<EngineDiagnostic> Diagnostics { get; } = new List<EngineDiagnostic>();
    }

    /// <summary>engines/&lt;name&gt;/engine.json</summary>
    public sealed class EngineManifest
    {
        public string Name { get; set; }
        public string Language { get; set; }
        public string Version { get; set; }
        public string Assembly { get; set; }
        public string Type { get; set; }
        public int Protocol { get; set; }
        public bool Sessions { get; set; }
    }
}
