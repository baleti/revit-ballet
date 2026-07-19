using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace RevitBallet.Commands
{
    /// <summary>
    /// Response from a /roslyn script execution on a peer session.
    /// Mirrors the server's ScriptResponse shape.
    /// </summary>
    public class RoslynResult
    {
        public bool Success { get; set; }
        public string Output { get; set; }
        public string Error { get; set; }
        public string[] Diagnostics { get; set; }
    }

    /// <summary>
    /// Shared HTTP client for talking to peer Revit sessions' Roslyn servers.
    /// Single place for certificate policy, auth token handling, and the
    /// parallel fan-out pattern used by InNetwork commands.
    /// </summary>
    public static class NetworkClient
    {
        // One client per process: avoids socket exhaustion from per-call
        // HttpClient creation on .NET Framework.
        private static readonly HttpClient client = CreateClient();

        private static HttpClient CreateClient()
        {
            var handler = new HttpClientHandler();
            // Peer sessions use self-signed certificates on 127.0.0.1. Trust
            // comes from the localhost binding plus the shared token, not the
            // certificate, so validation is intentionally disabled here - and
            // only here.
            handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;
            return new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(120) };
        }

        /// <summary>
        /// Reads the shared network auth token. Returns null when no session
        /// has generated one yet (caller should surface an error dialog).
        /// </summary>
        public static string GetSharedToken()
        {
            return RevitBalletServer.GetSharedAuthToken();
        }

        /// <summary>
        /// Escapes a string for embedding inside a double-quoted literal in a
        /// generated Roslyn script.
        /// </summary>
        public static string EscapeForScript(string value)
        {
            if (value == null) return "";
            return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }

        /// <summary>
        /// Executes a C# script on the session listening on the given port.
        /// Returns null on transport failure (connection refused, timeout).
        /// </summary>
        public static async Task<RoslynResult> ExecuteScriptAsync(int port, string script, string token)
        {
            try
            {
                var request = new HttpRequestMessage(HttpMethod.Post, $"https://127.0.0.1:{port}/roslyn")
                {
                    Content = new StringContent(script, Encoding.UTF8, "text/plain")
                };
                request.Headers.Add("X-Auth-Token", token);

                var response = await client.SendAsync(request).ConfigureAwait(false);
                var responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                return Newtonsoft.Json.JsonConvert.DeserializeObject<RoslynResult>(responseText);
            }
            catch (Exception ex)
            {
                Log.Warn("NetworkClient", $"Script execution on port {port} failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Synchronous convenience wrapper around ExecuteScriptAsync.
        /// </summary>
        public static RoslynResult ExecuteScript(int port, string script, string token)
        {
            return ExecuteScriptAsync(port, script, token).GetAwaiter().GetResult();
        }

        /// <summary>
        /// POSTs a JSON body to a pre-compiled query endpoint (e.g.
        /// /query/familytypes/counts) on the session listening on the given
        /// port. Returns the parsed result plus wall-clock time, with a null
        /// result on transport failure.
        /// </summary>
        public static async Task<(RoslynResult Result, double ElapsedMs)> PostQueryAsync(
            int port, string endpoint, string jsonBody, string token)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                var request = new HttpRequestMessage(HttpMethod.Post, $"https://127.0.0.1:{port}{endpoint}")
                {
                    Content = new StringContent(jsonBody, Encoding.UTF8, "application/json")
                };
                request.Headers.Add("X-Auth-Token", token);

                var response = await client.SendAsync(request).ConfigureAwait(false);
                var responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                stopwatch.Stop();
                return (Newtonsoft.Json.JsonConvert.DeserializeObject<RoslynResult>(responseText), stopwatch.Elapsed.TotalMilliseconds);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                Log.Warn("NetworkClient", $"Query {endpoint} on port {port} failed: {ex.Message}");
                return (null, stopwatch.Elapsed.TotalMilliseconds);
            }
        }

        /// <summary>
        /// Runs a per-document query-endpoint POST on many sessions in parallel.
        /// Returns each document paired with its result and elapsed time.
        /// </summary>
        public static List<(DocumentEntry Doc, RoslynResult Result, double ElapsedMs)> PostQueryOnDocuments(
            IEnumerable<DocumentEntry> documents,
            string endpoint,
            Func<DocumentEntry, string> jsonBodyFor,
            string token)
        {
            var docs = documents.ToList();
            var tasks = docs
                .Select(d => PostQueryAsync(d.Port, endpoint, jsonBodyFor(d), token))
                .ToArray();

            try
            {
                Task.WhenAll(tasks).GetAwaiter().GetResult();
            }
            catch
            {
                // Individual failures already yield null results below.
            }

            var results = new List<(DocumentEntry, RoslynResult, double)>();
            for (int i = 0; i < docs.Count; i++)
            {
                var (result, elapsed) = tasks[i].Status == TaskStatus.RanToCompletion
                    ? tasks[i].Result
                    : (null, 0.0);
                results.Add((docs[i], result, elapsed));
            }
            return results;
        }

        /// <summary>
        /// Runs a per-document script on many sessions in parallel and returns
        /// each document paired with its result (null result = transport failure).
        /// </summary>
        public static List<(DocumentEntry Doc, RoslynResult Result)> ExecuteOnDocuments(
            IEnumerable<DocumentEntry> documents,
            Func<DocumentEntry, string> scriptFor,
            string token)
        {
            var docs = documents.ToList();
            var tasks = docs
                .Select(d => ExecuteScriptAsync(d.Port, scriptFor(d), token))
                .ToArray();

            try
            {
                Task.WhenAll(tasks).GetAwaiter().GetResult();
            }
            catch
            {
                // Individual failures already yield null results below.
            }

            var results = new List<(DocumentEntry, RoslynResult)>();
            for (int i = 0; i < docs.Count; i++)
            {
                var r = tasks[i].Status == TaskStatus.RanToCompletion ? tasks[i].Result : null;
                results.Add((docs[i], r));
            }
            return results;
        }
    }
}
