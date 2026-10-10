using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace RevitBallet.Commands.Agent
{
    /// <summary>The /v1 HTTP routes (docs/agent-protocol.md). Authentication has already happened in the server.</summary>
    internal static class AgentHttp
    {
        public static async Task Handle(AgentCore core, Stream stream, string method, string target, string body, string connId, string sessionId, int port, string client = null)
        {
            string path = target, query = "";
            var q = target.IndexOf('?');
            if (q >= 0) { path = target.Substring(0, q); query = target.Substring(q + 1); }
            path = path.TrimEnd('/');
            var seg = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);   // v1, jobs, j-0001, events

            try
            {
                if (seg.Length == 2 && seg[1] == "info" && method == "GET")
                { await Json(stream, 200, core.Info(sessionId, port)); return; }

                if (seg.Length == 2 && seg[1] == "jobs")
                {
                    if (method == "POST") { await SubmitJob(core, stream, body, client); return; }
                    if (method == "GET")
                    {
                        var arr = new JArray(core.List().Select(j => (JToken)j.ToJson()));
                        await Json(stream, 200, new JObject { ["jobs"] = arr }); return;
                    }
                }

                if (seg.Length >= 3 && seg[1] == "jobs")
                {
                    var job = core.Get(seg[2]);
                    if (job == null) throw new AgentException(404, "not_found", "unknown job " + seg[2]);
                    if (seg.Length == 3 && method == "GET")
                    {
                        int wait = QueryInt(query, "wait", 0);
                        if (wait > 0) await Task.Run(() => core.WaitFinal(job, wait));
                        await Json(stream, 200, job.ToJson()); return;
                    }
                    if (seg.Length == 3 && method == "DELETE")
                    {
                        core.Cancel(job);
                        await Json(stream, 200, job.ToJson()); return;
                    }
                    if (seg.Length == 4 && seg[3] == "events" && method == "GET")
                    { await Events(stream, job); return; }
                    if (seg.Length == 4 && seg[3] == "output" && method == "GET")
                    {
                        int offset = QueryInt(query, "offset", 0), limit = QueryInt(query, "limit", 65536), total;
                        var text = job.OutputSlice(offset, limit, out total);
                        await Json(stream, 200, new JObject { ["offset"] = offset, ["length"] = text.Length, ["total"] = total, ["dropped"] = job.OutputDropped, ["text"] = text });
                        return;
                    }
                }

                if (seg.Length == 2 && seg[1] == "dialog" && method == "GET")
                {
                    double since; var d = core.CurrentDialog(out since);
                    await Json(stream, 200, new JObject { ["dialog"] = d == null ? null : d.ToJson(since) }); return;
                }
                if (seg.Length == 3 && seg[1] == "dialog" && seg[2] == "respond" && method == "POST")
                {
                    JObject rq;
                    try { rq = JObject.Parse(body ?? ""); } catch (Exception) { throw new AgentException(400, "bad_request", "body must be a JSON object with a \"button\""); }
                    var button = (string)rq["button"];
                    if (string.IsNullOrWhiteSpace(button)) throw new AgentException(400, "bad_request", "\"button\" is required");
                    await Json(stream, 200, core.RespondToDialog(button)); return;
                }

                if (seg.Length == 2 && seg[1] == "commands" && method == "GET")
                {
                    var list = core.Commands;
                    await Json(stream, 200, new JObject { ["count"] = list.Count, ["commands"] = new JArray(list.Select(c => (JToken)c.ToJson())) }); return;
                }
                if (seg.Length == 4 && seg[1] == "commands" && seg[3] == "invoke" && method == "POST")
                { await InvokeCommand(core, stream, Uri.UnescapeDataString(seg[2]), body, client); return; }

                if (seg.Length == 2 && seg[1] == "sessions" && method == "GET")
                {
                    var arr = new JArray(core.ListSessions().Select(s => (JToken)new JObject
                    { ["name"] = s.Name, ["jobs"] = s.Jobs, ["created"] = s.Created.ToString("o") }));
                    await Json(stream, 200, new JObject { ["sessions"] = arr }); return;
                }
                if (seg.Length == 4 && seg[1] == "sessions" && method == "DELETE")
                {
                    if (!core.DropSession(seg[2], seg[3])) throw new AgentException(404, "not_found", "unknown session");
                    await Json(stream, 200, new JObject { ["dropped"] = seg[3] }); return;
                }

                throw new AgentException(404, "not_found", $"no such endpoint: {method} {path}");
            }
            catch (AgentException ex) { await Error(stream, ex.Status, ex.Kind, ex.Message); }
            catch (Exception ex)
            {
                AgentLog.Write($"[CONN-{connId}] /v1 error: {ex}");
                await Error(stream, 500, "internal", ex.Message);
            }
        }

        /// <summary>"report" (default), "fail", or {"respond": "<button>"}.</summary>
        private static DialogPolicy ParsePolicy(JToken t)
        {
            if (t == null || t.Type == JTokenType.Null) return null;
            if (t.Type == JTokenType.String)
            {
                var v = ((string)t).Trim().ToLowerInvariant();
                if (v == "report") return null;
                if (v == "fail") return new DialogPolicy { Mode = DialogMode.Fail };
                throw new AgentException(400, "bad_request", "dialogPolicy must be \"report\", \"fail\" or {\"respond\": \"<button>\"}");
            }
            var o = t as JObject;
            var button = o == null ? null : (string)o["respond"];
            if (string.IsNullOrWhiteSpace(button)) throw new AgentException(400, "bad_request", "dialogPolicy must be \"report\", \"fail\" or {\"respond\": \"<button>\"}");
            return new DialogPolicy { Mode = DialogMode.Respond, Button = button };
        }

        private static async Task SubmitJob(AgentCore core, Stream stream, string body, string client)
        {
            JObject req;
            try { req = JObject.Parse(body ?? ""); }
            catch (Exception) { throw new AgentException(400, "bad_request", "body must be a JSON object"); }
            var job = core.Submit((string)req["engine"], (string)req["code"], (int?)req["timeoutMs"] ?? 0, (string)req["session"], client, (string)req["requestId"], ParsePolicy(req["dialogPolicy"]));
            int wait = (int?)req["wait"] ?? 0;
            if (wait > 0) await Task.Run(() => core.WaitFinal(job, wait));
            await Json(stream, job.Done.IsSet ? 200 : 202, job.ToJson());
        }

        /// <summary>POST /v1/commands/{name}/invoke: a job of the built-in "command" engine, so it queues, times out and reports dialogs like any job.
        /// Commands often wait for the user (a picker, a dialog), so the default time limit is the maximum.</summary>
        private static async Task InvokeCommand(AgentCore core, Stream stream, string name, string body, string client)
        {
            JObject req = new JObject();
            if (!string.IsNullOrWhiteSpace(body))
            {
                try { req = JObject.Parse(body); }
                catch (Exception) { throw new AgentException(400, "bad_request", "body must be empty or a JSON object"); }
            }
            var cmd = core.Commands.FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
            if (cmd == null) throw new AgentException(404, "not_found", "unknown command '" + name + "'");
            var job = core.Submit("command", cmd.Name, (int?)req["timeoutMs"] ?? AgentCore.MaxTimeoutMs, null, client, (string)req["requestId"], ParsePolicy(req["dialogPolicy"]));
            int wait = (int?)req["wait"] ?? 0;
            if (wait > 0) await Task.Run(() => core.WaitFinal(job, wait));
            await Json(stream, job.Done.IsSet ? 200 : 202, job.ToJson());
        }

        private static async Task Events(Stream stream, Job job)
        {
            var head = "HTTP/1.1 200 OK\r\nContent-Type: text/event-stream\r\nCache-Control: no-cache\r\nConnection: close\r\n\r\n";
            await WriteAsync(stream, head);
            int sent = 0;
            while (true)
            {
                string[] pending; bool final;
                lock (job.Events)
                {
                    if (job.Events.Count == sent && !job.Done.IsSet) Monitor.Wait(job.Events, 15000);
                    pending = job.Events.Skip(sent).ToArray(); sent += pending.Length;
                    final = job.Done.IsSet && sent == job.Events.Count;
                }
                if (pending.Length == 0) await WriteAsync(stream, ": keepalive\n\n");
                foreach (var e in pending) await WriteAsync(stream, e);
                if (final && pending.Any(e => e.StartsWith("event: done"))) return;
                if (final && pending.Length == 0) return;
            }
        }

        private static int QueryInt(string query, string key, int dflt)
        {
            foreach (var kv in query.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var p = kv.Split(new[] { '=' }, 2);
                int v; if (p.Length == 2 && p[0] == key && int.TryParse(p[1], out v)) return Math.Max(0, Math.Min(v, 100000000));
            }
            return dflt;
        }

        private static string Reason(int status)
        {
            switch (status)
            {
                case 200: return "OK"; case 202: return "Accepted"; case 400: return "Bad Request"; case 401: return "Unauthorized";
                case 403: return "Forbidden"; case 204: return "No Content"; case 404: return "Not Found"; case 409: return "Conflict"; case 413: return "Payload Too Large";
                case 503: return "Service Unavailable"; default: return "Error";
            }
        }

        public static Task Error(Stream stream, int status, string kind, string message) =>
            Json(stream, status, new JObject { ["error"] = new JObject { ["kind"] = kind, ["message"] = message } });

        public static async Task Json(Stream stream, int status, JToken body)
        {
            var bytes = Encoding.UTF8.GetBytes(body.ToString(Newtonsoft.Json.Formatting.None));
            var head = $"HTTP/1.1 {status} {Reason(status)}\r\nContent-Type: application/json; charset=utf-8\r\nContent-Length: {bytes.Length}\r\nConnection: close\r\n\r\n";
            var h = Encoding.ASCII.GetBytes(head);
            await stream.WriteAsync(h, 0, h.Length);
            await stream.WriteAsync(bytes, 0, bytes.Length);
            await stream.FlushAsync();
        }

        private static async Task WriteAsync(Stream stream, string text)
        {
            var b = Encoding.UTF8.GetBytes(text);
            await stream.WriteAsync(b, 0, b.Length);
            await stream.FlushAsync();
        }
    }
}
