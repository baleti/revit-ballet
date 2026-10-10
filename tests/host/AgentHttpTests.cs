using System;
using System.Linq;
using System.Text;
using System.Threading;
using Newtonsoft.Json.Linq;
using RevitBallet.Commands.Agent;
using Xunit;

namespace RevitBallet.Agent.Tests
{
    /// <summary>The /v1 routes (authentication happens in the server, not here), driven through AgentHttp.Handle over in-memory streams.</summary>
    public class AgentHttpTests
    {
        private static JObject Json(string s) { return JObject.Parse(s); }
        private static string Body(string engine, string code, object extra = null)
        {
            var o = new JObject { ["engine"] = engine, ["code"] = code, ["timeoutMs"] = 10000 };
            if (extra != null) foreach (var p in JObject.FromObject(extra).Properties()) o[p.Name] = p.Value;
            return o.ToString();
        }

        [Fact]
        public void Info_is_json_with_protocol_and_engines()
        {
            using (var rig = new Rig())
            {
                var (status, body) = rig.Http("GET", "/v1/info");
                Assert.Equal(200, status);
                var j = Json(body);
                Assert.Equal(1, (int)j["protocol"]);
                Assert.Equal("idle", (string)j["ui"]["state"]);
                Assert.Equal(4242, (int)j["session"]["pid"]);
                Assert.Equal(23717, (int)j["session"]["port"]);
            }
        }

        [Fact]
        public void Post_job_waits_for_the_result_or_answers_202()
        {
            using (var rig = new Rig())
            {
                var (s1, b1) = rig.Http("POST", "/v1/jobs", Body("roslyn", "ok 9", new { wait = 5000 }));
                Assert.Equal(200, s1);
                Assert.Equal("done", (string)Json(b1)["state"]);
                Assert.Equal(9, (int)Json(b1)["value"]["json"]);

                rig.Ui.Blocked = true;
                var (s2, b2) = rig.Http("POST", "/v1/jobs", Body("roslyn", "ok 1"));       // no wait: accepted, not finished
                Assert.Equal(202, s2);
                Assert.Equal("queued", (string)Json(b2)["state"]);
                rig.Ui.Blocked = false;
            }
        }

        [Fact]
        public void Get_job_long_polls_and_unknown_jobs_are_404()
        {
            using (var rig = new Rig())
            {
                var job = rig.Core.Submit("roslyn", "sleep 300", 10000, null);
                var (s, b) = rig.Http("GET", "/v1/jobs/" + job.Id + "?wait=5000");
                Assert.Equal(200, s);
                Assert.Equal("done", (string)Json(b)["state"]);
                var (s404, b404) = rig.Http("GET", "/v1/jobs/j-9999");
                Assert.Equal(404, s404);
                Assert.Equal("not_found", (string)Json(b404)["error"]["kind"]);
                var (sl, bl) = rig.Http("GET", "/v1/jobs");
                Assert.Equal(200, sl);
                Assert.True(((JArray)Json(bl)["jobs"]).Count >= 1);
            }
        }

        [Fact]
        public void Delete_cancels_a_queued_job()
        {
            using (var rig = new Rig())
            {
                rig.Ui.Blocked = true;
                var job = rig.Core.Submit("roslyn", "ok 1", 10000, null);
                var (s, b) = rig.Http("DELETE", "/v1/jobs/" + job.Id);
                Assert.Equal(200, s);
                Assert.Equal("cancelled", (string)Json(b)["state"]);
                rig.Ui.Blocked = false;
            }
        }

        [Fact]
        public void Bad_requests_get_clear_statuses()
        {
            using (var rig = new Rig())
            {
                Assert.Equal(400, rig.Http("POST", "/v1/jobs", "not json").Status);
                Assert.Equal(400, rig.Http("POST", "/v1/jobs", "{\"engine\":\"roslyn\"}").Status);          // no code
                Assert.Equal(404, rig.Http("POST", "/v1/jobs", Body("nope", "ok 1")).Status);               // unknown engine
                var (s409, b409) = rig.Http("POST", "/v1/jobs", Body("cpython3", "ok 1"));                  // known but not installed
                Assert.Equal(409, s409);
                Assert.Equal("engine", (string)Json(b409)["error"]["kind"]);
                Assert.Equal(404, rig.Http("GET", "/v1/nothing").Status);
                Assert.Equal(404, rig.Http("PUT", "/v1/jobs").Status);
            }
        }

        [Fact]
        public void Output_endpoint_serves_slices()
        {
            using (var rig = new Rig())
            {
                var job = rig.Run("print 1000");
                var (s, b) = rig.Http("GET", "/v1/jobs/" + job.Id + "/output?offset=7&limit=12");
                Assert.Equal(200, s);
                var j = Json(b);
                Assert.Equal(12, (int)j["length"]);
                Assert.Equal(7, (int)j["offset"]);
                Assert.True((int)j["total"] > 1000);
                Assert.StartsWith("line 1", (string)j["text"]);
            }
        }

        [Fact]
        public void Sessions_can_be_listed_and_dropped()
        {
            using (var rig = new Rig())
            {
                rig.Run("set a 1", session: "mine");
                var list = (JArray)Json(rig.Http("GET", "/v1/sessions").Body)["sessions"];
                Assert.Contains(list, x => (string)x["name"] == "mine");
                Assert.Equal(200, rig.Http("DELETE", "/v1/sessions/roslyn/mine").Status);
                Assert.Equal(404, rig.Http("DELETE", "/v1/sessions/roslyn/mine").Status);
            }
        }

        [Fact]
        public void Events_stream_ends_with_the_done_event_and_carries_output()
        {
            using (var rig = new Rig())
            {
                var job = rig.Run("print 2");
                var (s, b) = rig.Http("GET", "/v1/jobs/" + job.Id + "/events");
                Assert.Equal(200, s);
                Assert.Contains("event: state", b);
                Assert.Contains("event: output", b);
                Assert.Contains("line 1", b);
                Assert.True(b.TrimEnd().EndsWith("}") && b.Contains("event: done"));
            }
        }

        [Fact]
        public void Client_name_and_request_id_reach_the_job()
        {
            using (var rig = new Rig())
            {
                var body = Body("roslyn", "ok 3", new { wait = 5000, requestId = "abc" });
                var first = Json(rig.Http("POST", "/v1/jobs", body, "agent-x").Body);
                var again = Json(rig.Http("POST", "/v1/jobs", body, "agent-x").Body);
                Assert.Equal("agent-x", (string)first["client"]);
                Assert.Equal((string)first["id"], (string)again["id"]);
                Assert.Equal(1, rig.Roslyn.Runs);
            }
        }

        [Fact]
        public void Non_ascii_code_survives_the_round_trip()
        {
            using (var rig = new Rig())
            {
                var code = "ok héllo € 中文";
                var j = Json(rig.Http("POST", "/v1/jobs", Body("roslyn", code, new { wait = 5000 })).Body);
                Assert.Equal("héllo € 中文", (string)j["value"]["json"]);
            }
        }
    }
}
