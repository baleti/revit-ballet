using System.IO;
using System.Threading.Tasks;

namespace RevitBallet.Agents
{
    /// <summary>
    /// The reloadable part of the agent server (job queue, engines, dialogs, the /v1 API). The add-in keeps a small stable shell (HTTP listener, authentication,
    /// the Idling hook) and talks to the runtime only through this interface, so a new build of the runtime can be swapped in while Revit runs: the shell loads
    /// revit-ballet.dll from the hot-reload folder, the same one InvokeAddinCommand uses, and replaces <see cref="AgentRuntimeRegistry.Current"/>.
    /// This assembly is loaded once and shared by the shell and every runtime generation, which is what keeps the interface the same type on both sides.
    /// </summary>
    public interface IAgentRuntime
    {
        /// <summary>Identifies the loaded code (assembly version and module id), shown in GET /v1/info so a reload can be confirmed.</summary>
        string Build { get; }
        /// <summary>"installed" or "hot-reload".</summary>
        string LoadedFrom { get; }

        /// <summary>Serve one /v1 request: write the complete HTTP response to <paramref name="stream"/>.</summary>
        Task HandleAsync(Stream stream, string method, string target, string body, string connId, string sessionId, int port, string client);

        /// <summary>The legacy POST /roslyn (raw C# body), answered in the old response shape.</summary>
        Task<LegacyReply> HandleLegacyRoslynAsync(string code);

        /// <summary>Revit UI thread, API context (the Idling event): run a queued job if one is waiting.</summary>
        void OnIdle(object uiApplication);

        /// <summary>Revit UI thread, API context: this runtime is being replaced or Revit is closing. Unsubscribe events, stop the worker.</summary>
        void Detach(object uiApplication);
    }

    public sealed class LegacyReply
    {
        public int Status { get; set; }
        public string Json { get; set; }
    }

    /// <summary>The runtime currently serving requests. Null until the first Idling event after Revit starts.</summary>
    public static class AgentRuntimeRegistry
    {
        public static volatile IAgentRuntime Current;
    }
}
