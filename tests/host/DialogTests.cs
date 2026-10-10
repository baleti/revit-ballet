using System;
using System.Linq;
using System.Threading;
using Newtonsoft.Json.Linq;
using RevitBallet.Commands.Agent;
using Xunit;

namespace RevitBallet.Agent.Tests
{
    /// <summary>A modal dialog that holds Revit up: what is reported, how fast, and what a job's policy does about it.</summary>
    public class DialogTests
    {
        [Fact]
        public void Info_describes_the_open_dialog_with_text_buttons_and_age()
        {
            using (var rig = new Rig())
            {
                Assert.Equal(JTokenType.Null, rig.Core.Info("4242", 1)["ui"]["modalDialog"].Type);
                rig.Ui.Blocked = true;
                rig.Core.Info("4242", 1);            // first sight of the dialog
                Thread.Sleep(150);
                var ui = rig.Core.Info("4242", 1)["ui"];
                Assert.Equal("modal", (string)ui["state"]);
                Assert.Equal("Fake modal dialog", (string)ui["modalDialog"]["title"]);
                Assert.Equal("Something unexpected happened", (string)ui["modalDialog"]["text"]);
                Assert.Equal(new[] { "OK", "Cancel" }, ((JArray)ui["modalDialog"]["buttons"]).Select(b => (string)b));
                Assert.True((double)ui["modalDialog"]["sinceMs"] >= 100);
                rig.Ui.Blocked = false;
            }
        }

        [Fact]
        public void A_job_waiting_behind_a_dialog_reports_it_quickly_while_it_is_still_queued()
        {
            using (var rig = new Rig())
            {
                rig.Ui.Blocked = true;
                var job = rig.Core.Submit("roslyn", "ok 1", 20000, null);
                Assert.True(Wait.Until(() => job.BlockedBy != null, 3000), "the job never reported the dialog");
                var json = job.ToJson();
                Assert.Equal("queued", (string)json["state"]);                              // not finished, not failed: it is waiting
                Assert.Equal("Fake modal dialog", (string)json["blockedBy"]["title"]);
                Assert.Equal(1, ((JArray)json["dialogs"]).Count);
                Assert.Contains("event: dialog", string.Join("", job.Events));
                rig.Ui.Blocked = false;                                                    // the dialog goes away: the job carries on
                Assert.True(rig.Core.WaitFinal(job, 5000));
                Assert.Equal(JobState.Done, job.State);
                Assert.Equal(JTokenType.Null, job.ToJson()["blockedBy"].Type);
            }
        }

        [Fact]
        public void A_dialog_that_only_flashes_by_is_not_reported()
        {
            using (var rig = new Rig())
            {
                var job = rig.Core.Submit("roslyn", "sleep 400", 20000, null);
                Assert.True(Wait.Until(() => job.State == JobState.Running));
                rig.Ui.Blocked = true; Thread.Sleep(100); rig.Ui.Blocked = false;           // shorter than the grace period
                Assert.True(rig.Core.WaitFinal(job, 5000));
                Assert.Empty((JArray)job.ToJson()["dialogs"]);
            }
        }

        [Fact]
        public void A_script_that_opens_a_dialog_is_reported_while_it_runs_not_after_the_timeout()
        {
            using (var rig = new Rig())
            {
                var job = rig.Core.Submit("roslyn", "open-dialog", 30000, null);            // 30 s timeout, the dialog must be seen long before
                Assert.True(Wait.Until(() => job.BlockedBy != null, 3000));
                Assert.Equal("running", (string)job.ToJson()["state"]);
                var answered = rig.Core.RespondToDialog("OK");
                Assert.True((bool)answered["responded"]);
                Assert.True(rig.Core.WaitFinal(job, 5000));
                Assert.Equal(JobState.Done, job.State);
                Assert.Equal("dialog answered", (string)job.ToJson()["value"]["json"]);
            }
        }

        [Fact]
        public void Policy_fail_ends_the_job_at_once_with_the_dialog_in_the_error()
        {
            using (var rig = new Rig())
            {
                var job = rig.Core.Submit("roslyn", "open-dialog", 30000, null, dialogPolicy: new DialogPolicy { Mode = DialogMode.Fail });
                Assert.True(rig.Core.WaitFinal(job, 5000), "should have failed long before the 30 s timeout");
                var e = job.ToJson()["error"];
                Assert.True(job.State == JobState.Failed, job.ToJson().ToString());
                Assert.Equal("blocked_by_dialog", (string)e["kind"]);
                Assert.Contains("Something unexpected happened", (string)e["message"]);
                Assert.Equal("Fake modal dialog", (string)e["dialog"]["title"]);
                Assert.Equal("OK", (string)e["dialog"]["buttons"][0]);
                Assert.Equal("fail", (string)job.ToJson()["dialogs"][0]["action"]);
                Assert.Equal(JobState.Done, rig.Run("ok 1").State);                       // the queue recovers once the UI thread is free again
            }
        }

        [Fact]
        public void Policy_respond_answers_the_dialog_and_the_job_completes()
        {
            using (var rig = new Rig())
            {
                var job = rig.Core.Submit("roslyn", "open-dialog", 30000, null, dialogPolicy: new DialogPolicy { Mode = DialogMode.Respond, Button = "ok" });
                Assert.True(rig.Core.WaitFinal(job, 5000));
                Assert.Equal(JobState.Done, job.State);
                Assert.Equal(new[] { "ok" }, rig.Ui.Responses.ToArray());
                Assert.Equal("responded:ok", (string)job.ToJson()["dialogs"][0]["action"]);
            }
        }

        [Fact]
        public void Policy_respond_with_a_button_the_dialog_does_not_have_fails_the_job()
        {
            using (var rig = new Rig())
            {
                var job = rig.Core.Submit("roslyn", "open-dialog", 30000, null, dialogPolicy: new DialogPolicy { Mode = DialogMode.Respond, Button = "Retry" });
                Assert.True(rig.Core.WaitFinal(job, 5000));
                Assert.Equal(JobState.Failed, job.State);
                Assert.Contains("Retry", (string)job.ToJson()["error"]["message"]);
                Assert.Contains("OK, Cancel", (string)job.ToJson()["error"]["message"]);
            }
        }

        [Fact]
        public void The_timeout_of_a_blocked_job_names_the_dialog_and_its_buttons()
        {
            using (var rig = new Rig())
            {
                rig.Ui.Blocked = true;
                var j = rig.Run("ok 1", timeoutMs: 700);
                var e = j.ToJson()["error"];
                Assert.Equal(JobState.Timeout, j.State);
                Assert.Equal("blocked_by_dialog", (string)e["kind"]);
                Assert.Contains("Fake modal dialog", (string)e["message"]);
                Assert.Contains("OK, Cancel", (string)e["message"]);
                rig.Ui.Blocked = false;
            }
        }

        [Fact]
        public void Dialog_endpoints_report_and_answer_independently_of_the_queue()
        {
            using (var rig = new Rig())
            {
                Assert.Equal(JTokenType.Null, JObject.Parse(rig.Http("GET", "/v1/dialog").Body)["dialog"].Type);
                Assert.Equal(404, rig.Http("POST", "/v1/dialog/respond", "{\"button\":\"OK\"}").Status);        // nothing to answer
                rig.Ui.Blocked = true;
                var stuck = rig.Core.Submit("roslyn", "ok 1", 20000, null);                                      // queue is held up by the dialog
                var (s, b) = rig.Http("GET", "/v1/dialog");
                Assert.Equal(200, s);
                Assert.Equal("Fake modal dialog", (string)JObject.Parse(b)["dialog"]["title"]);
                Assert.Equal(409, rig.Http("POST", "/v1/dialog/respond", "{\"button\":\"Retry\"}").Status);      // no such button
                Assert.Equal(400, rig.Http("POST", "/v1/dialog/respond", "{}").Status);
                var (sr, br) = rig.Http("POST", "/v1/dialog/respond", "{\"button\":\"Cancel\"}");
                Assert.Equal(200, sr);
                Assert.True((bool)JObject.Parse(br)["responded"]);
                Assert.True(rig.Core.WaitFinal(stuck, 5000));                                                    // answering released the queued job
                Assert.Equal(JobState.Done, stuck.State);
            }
        }

        [Fact]
        public void Dialog_policy_is_validated_in_the_post_body()
        {
            using (var rig = new Rig())
            {
                Assert.Equal(400, rig.Http("POST", "/v1/jobs", "{\"engine\":\"roslyn\",\"code\":\"ok 1\",\"dialogPolicy\":\"explode\"}").Status);
                Assert.Equal(400, rig.Http("POST", "/v1/jobs", "{\"engine\":\"roslyn\",\"code\":\"ok 1\",\"dialogPolicy\":{\"respond\":\"\"}}").Status);
                var ok = rig.Http("POST", "/v1/jobs", "{\"engine\":\"roslyn\",\"code\":\"ok 1\",\"dialogPolicy\":{\"respond\":\"OK\"},\"wait\":5000}");
                Assert.Equal(200, ok.Status);
            }
        }
    }
}
