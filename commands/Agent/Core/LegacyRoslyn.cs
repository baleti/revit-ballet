using System;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using RevitBallet.Agents;

namespace RevitBallet.Commands.Agent
{
    /// <summary>The old POST /roslyn (raw C# body) on top of the job queue, answered in the old {Success, Output, Error, Diagnostics, ProcessingLog} shape.</summary>
    internal static class LegacyRoslyn
    {
        public static async Task<LegacyReply> HandleAsync(AgentCore core, string code)
        {
            var log = new JArray { Stamp("Request received by server") };
            if (string.IsNullOrWhiteSpace(code))
                return Reply(400, false, null, "Empty request body", null, log);

            Job job;
            try { job = core.Submit("roslyn", code, 30000, null); }
            catch (AgentException ex)
            {
                var msg = ex.Kind == "engine" ? "The Roslyn engine is not available: " + ex.Message : ex.Message;
                return Reply(ex.Status == 409 ? 503 : ex.Status, false, null, msg, null, log);
            }
            log.Add(Stamp("Job " + job.Id + " queued"));
            await Task.Run(() => core.WaitFinal(job, 120000));
            var j = job.ToJson();
            log.Add(Stamp($"Job finished: {(string)j["state"]} (compile {j["timings"]["compileMs"]} ms, run {j["timings"]["runMs"]} ms)"));

            bool ok = (bool)j["ok"];
            var output = (string)j["output"];
            var diags = new JArray(((JArray)j["diagnostics"]).Select(d =>
                (JToken)$"({(int)d["line"]},{(int)d["column"]}): {(string)d["severity"]} {(string)d["id"]}: {(string)d["message"]}"));
            string error = null;
            if (ok)
            {
                var v = j["value"];
                if (v != null && v.Type != JTokenType.Null && (string)v["text"] != null)
                {
                    if (!string.IsNullOrEmpty(output)) output += "\n";
                    output += "Return value: " + (string)v["text"];
                }
            }
            else
            {
                var e = j["error"];
                var kind = e == null ? null : (string)e["kind"];
                var msg = e == null ? "unknown error" : (string)e["message"];
                if (kind == "compile") error = "Compilation failed";
                else if (kind == "timeout" || kind == "blocked_by_dialog") error = "Script execution timeout (30 seconds). " + msg;
                else error = msg + (string.IsNullOrEmpty((string)e?["stack"]) ? "" : "\n\nStack Trace:\n" + (string)e["stack"]);
            }
            return Reply(ok ? 200 : 400, ok, output, error, diags, log);
        }

        private static string Stamp(string m) { return "[" + DateTime.Now.ToString("HH:mm:ss.fff") + "] " + m; }

        private static LegacyReply Reply(int status, bool success, string output, string error, JArray diagnostics, JArray log)
        {
            var o = new JObject
            {
                ["Success"] = success, ["Output"] = output, ["Error"] = error,
                ["Diagnostics"] = diagnostics ?? new JArray(), ["ProcessingLog"] = log,
            };
            return new LegacyReply { Status = status, Json = o.ToString(Newtonsoft.Json.Formatting.None) };
        }
    }
}
