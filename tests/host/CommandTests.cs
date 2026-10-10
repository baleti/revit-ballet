using System;
using System.Linq;
using Newtonsoft.Json.Linq;
using RevitBallet.Agents;
using RevitBallet.Commands.Agent;
using Xunit;

namespace RevitBallet.Agent.Tests
{
    /// <summary>GET /v1/commands and POST /v1/commands/{name}/invoke: the catalogue and running a command as a job of the "command" engine.</summary>
    public class CommandTests
    {
        private static JObject Json(string s) { return JObject.Parse(s); }

        [Fact]
        public void List_returns_every_command_with_its_metadata()
        {
            using (var rig = new Rig())
            {
                var (status, body) = rig.Http("GET", "/v1/commands");
                Assert.Equal(200, status);
                var j = Json(body);
                Assert.Equal(3, (int)j["count"]);
                var cmds = (JArray)j["commands"];
                Assert.Equal(new[] { "DeselectRandomly", "ExportToPDF", "DrawCircleAtOrigin" }, cmds.Select(c => (string)c["name"]).ToArray());
                var pdf = cmds.Single(c => (string)c["name"] == "ExportToPDF");
                Assert.Equal("Export To PDF", (string)pdf["displayName"]);
                Assert.Equal("PDF", (string)pdf["output"]);
                Assert.Equal("Any", (string)cmds.Single(c => (string)c["name"] == "DeselectRandomly")["input"]);
            }
        }

        [Fact]
        public void List_is_empty_when_the_host_has_no_commands()
        {
            using (var rig = new Rig())
            {
                rig.Ui.Commands = new System.Collections.Generic.List<CommandInfo>();
                var j = Json(rig.Http("GET", "/v1/commands").Body);
                Assert.Equal(0, (int)j["count"]);
                Assert.Empty((JArray)j["commands"]);
            }
        }

        [Fact]
        public void Invoke_runs_the_command_as_a_job_and_waits_when_asked()
        {
            using (var rig = new Rig())
            {
                var (status, body) = rig.Http("POST", "/v1/commands/DeselectRandomly/invoke", "{\"wait\":5000}");
                Assert.Equal(200, status);
                var j = Json(body);
                Assert.Equal("done", (string)j["state"]);
                Assert.Equal("command", (string)j["engine"]);
                Assert.Equal("DeselectRandomly", (string)j["value"]["json"]);        // the fake engine echoes the command name it was given
                Assert.Equal("DeselectRandomly", rig.Command.RunLog.Single());
            }
        }

        [Fact]
        public void Invoke_without_wait_answers_202_and_the_job_finishes_later()
        {
            using (var rig = new Rig())
            {
                rig.Ui.Blocked = true;
                var (status, body) = rig.Http("POST", "/v1/commands/ExportToPDF/invoke");
                Assert.Equal(202, status);
                var id = (string)Json(body)["id"];
                Assert.Equal("queued", (string)Json(body)["state"]);
                rig.Ui.Blocked = false;
                var job = rig.Core.Get(id);
                Assert.True(rig.Core.WaitFinal(job, 5000));
                Assert.Equal("done", (string)Json(rig.Http("GET", "/v1/jobs/" + id).Body)["state"]);
            }
        }

        [Fact]
        public void Command_name_is_case_insensitive_and_url_decoded()
        {
            using (var rig = new Rig())
            {
                var j = Json(rig.Http("POST", "/v1/commands/drawcircleatorigin/invoke", "{\"wait\":5000}").Body);
                Assert.Equal("DrawCircleAtOrigin", (string)j["value"]["json"]);        // the catalogue's spelling is what runs
                Assert.Equal(200, rig.Http("POST", "/v1/commands/Draw%43ircleAtOrigin/invoke", "{\"wait\":5000}").Status);
            }
        }

        [Fact]
        public void Unknown_command_is_404_and_nothing_is_queued()
        {
            using (var rig = new Rig())
            {
                var (status, body) = rig.Http("POST", "/v1/commands/NoSuchThing/invoke");
                Assert.Equal(404, status);
                Assert.Equal("not_found", (string)Json(body)["error"]["kind"]);
                Assert.Empty(rig.Core.List());
            }
        }

        [Fact]
        public void Bad_body_is_400_empty_body_is_fine()
        {
            using (var rig = new Rig())
            {
                Assert.Equal(400, rig.Http("POST", "/v1/commands/ExportToPDF/invoke", "[1,2").Status);
                Assert.Equal(202, rig.Http("POST", "/v1/commands/ExportToPDF/invoke", "").Status);
                Assert.Equal(400, rig.Http("POST", "/v1/commands/ExportToPDF/invoke", "{\"dialogPolicy\":\"explode\"}").Status);
            }
        }

        [Fact]
        public void The_default_time_limit_is_the_maximum_because_commands_wait_for_people()
        {
            using (var rig = new Rig())
            {
                rig.Ui.Blocked = true;
                var id = (string)Json(rig.Http("POST", "/v1/commands/ExportToPDF/invoke").Body)["id"];
                Assert.Equal(AgentCore.MaxTimeoutMs, rig.Core.Get(id).TimeoutMs);
                var id2 = (string)Json(rig.Http("POST", "/v1/commands/ExportToPDF/invoke", "{\"timeoutMs\":5000}").Body)["id"];
                Assert.Equal(5000, rig.Core.Get(id2).TimeoutMs);
                rig.Ui.Blocked = false;
            }
        }

        [Fact]
        public void Invocations_queue_in_submission_order()
        {
            using (var rig = new Rig())
            {
                rig.Ui.Blocked = true;
                var ids = new[] { "DeselectRandomly", "ExportToPDF", "DrawCircleAtOrigin" }
                    .Select(n => (string)Json(rig.Http("POST", "/v1/commands/" + n + "/invoke").Body)["id"]).ToArray();
                rig.Ui.Blocked = false;
                foreach (var id in ids) Assert.True(rig.Core.WaitFinal(rig.Core.Get(id), 5000));
                Assert.Equal(new[] { "DeselectRandomly", "ExportToPDF", "DrawCircleAtOrigin" }, rig.Command.RunLog.ToArray());
            }
        }

        [Fact]
        public void A_modal_dialog_is_reported_on_the_command_and_the_policy_applies()
        {
            using (var rig = new Rig())
            {
                rig.Ui.Blocked = true;
                var (status, body) = rig.Http("POST", "/v1/commands/ExportToPDF/invoke", "{\"dialogPolicy\":\"fail\",\"wait\":5000}");
                Assert.Equal(200, status);
                var j = Json(body);
                Assert.Equal("failed", (string)j["state"]);
                Assert.Equal("blocked_by_dialog", (string)j["error"]["kind"]);
            }
        }

        [Fact]
        public void Request_id_makes_a_retry_return_the_same_job()
        {
            using (var rig = new Rig())
            {
                var a = Json(rig.Http("POST", "/v1/commands/DeselectRandomly/invoke", "{\"requestId\":\"r1\",\"wait\":5000}").Body);
                var b = Json(rig.Http("POST", "/v1/commands/DeselectRandomly/invoke", "{\"requestId\":\"r1\",\"wait\":5000}").Body);
                Assert.Equal((string)a["id"], (string)b["id"]);
                Assert.Single(rig.Command.RunLog);
            }
        }

        [Fact]
        public void The_command_engine_can_also_be_used_through_the_plain_job_api()
        {
            using (var rig = new Rig())
            {
                var j = Json(rig.Http("POST", "/v1/jobs", "{\"engine\":\"command\",\"code\":\"ExportToPDF\",\"wait\":5000}").Body);
                Assert.Equal("done", (string)j["state"]);
            }
        }

        [Fact]
        public void Without_the_command_engine_invoke_is_a_clear_error()
        {
            using (var rig = new Rig())
            {
                rig.Source.Map.Remove("command");
                var (status, body) = rig.Http("POST", "/v1/commands/ExportToPDF/invoke");
                Assert.True(status == 404 || status == 409, "status " + status);
                Assert.Equal("engine", (string)Json(body)["error"]["kind"]);
            }
        }
    }
}
