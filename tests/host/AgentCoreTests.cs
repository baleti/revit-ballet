using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using RevitBallet.Commands.Agent;
using Xunit;

namespace RevitBallet.Agent.Tests
{
    public class AgentCoreTests
    {
        [Fact]
        public void Job_returns_value_and_output()
        {
            using (var rig = new Rig())
            {
                var j = rig.Run("print 3");
                Assert.Equal(JobState.Done, j.State);
                var json = j.ToJson();
                Assert.True((bool)json["ok"]);
                Assert.Equal("line 0\nline 1\nline 2\n", ((string)json["output"]).Replace("\r", ""));
                var v = rig.Run("ok 42");
                Assert.Equal(42, (int)v.ToJson()["value"]["json"]);
            }
        }

        [Fact]
        public void Compile_error_reports_kind_and_position()
        {
            using (var rig = new Rig())
            {
                var j = rig.Run("compile-error");
                var json = j.ToJson();
                Assert.Equal("failed", (string)json["state"]);
                Assert.Equal("compile", (string)json["error"]["kind"]);
                Assert.Equal(1, (int)json["diagnostics"][0]["line"]);
                Assert.Equal(0, rig.Roslyn.Runs);   // never ran
            }
        }

        [Fact]
        public void Runtime_error_reports_message_and_stack()
        {
            using (var rig = new Rig())
            {
                var e = rig.Run("throw kaboom").ToJson()["error"];
                Assert.Equal("runtime", (string)e["kind"]);
                Assert.Equal("kaboom", (string)e["message"]);
                Assert.Equal("at fake", (string)e["stack"]);
            }
        }

        [Fact]
        public void Engine_errors_are_http_like_status_codes()
        {
            using (var rig = new Rig())
            {
                Assert.Equal(404, Assert.Throws<AgentException>(() => rig.Core.Submit("nope", "ok 1", 1000, null)).Status);
                Assert.Equal(409, Assert.Throws<AgentException>(() => rig.Core.Submit("cpython3", "ok 1", 1000, null)).Status);   // known, not installed
                Assert.Equal(409, Assert.Throws<AgentException>(() => rig.Core.Submit("broken", "ok 1", 1000, null)).Status);     // load error
                Assert.Equal(400, Assert.Throws<AgentException>(() => rig.Core.Submit("roslyn", null, 1000, null)).Status);
                Assert.Equal(400, Assert.Throws<AgentException>(() => rig.Core.Submit("roslyn", "ok 1", 99999999, null)).Status);
                Assert.Equal(413, Assert.Throws<AgentException>(() => rig.Core.Submit("roslyn", new string('x', 5 * 1024 * 1024), 1000, null)).Status);
            }
        }

        [Fact]
        public void Jobs_run_in_submission_order_and_none_is_lost()
        {
            using (var rig = new Rig())
            {
                var jobs = Enumerable.Range(0, 100).Select(i => rig.Core.Submit("roslyn", "ok " + i, 10000, null)).ToList();
                foreach (var j in jobs) Assert.True(rig.Core.WaitFinal(j, 30000));
                Assert.All(jobs, j => Assert.Equal(JobState.Done, j.State));
                Assert.Equal(Enumerable.Range(0, 100).Select(i => "ok " + i), rig.Roslyn.RunLog);
                Assert.Equal(Enumerable.Range(0, 100).Select(i => (long)i), jobs.Select(j => (long)j.ToJson()["value"]["json"]));
            }
        }

        [Fact]
        public void Concurrent_submitters_lose_nothing()
        {
            using (var rig = new Rig())
            {
                var all = new System.Collections.Concurrent.ConcurrentBag<Job>();
                Parallel.For(0, 8, new ParallelOptions { MaxDegreeOfParallelism = 8 }, t =>
                {
                    for (int i = 0; i < 25; i++) all.Add(rig.Core.Submit(t % 2 == 0 ? "roslyn" : "ironpython2", "ok " + (t * 100 + i), 10000, null));
                });
                Assert.Equal(200, all.Count);
                foreach (var j in all) Assert.True(rig.Core.WaitFinal(j, 30000));
                Assert.All(all, j => Assert.Equal(JobState.Done, j.State));
                Assert.Equal(200, rig.Roslyn.Runs + rig.Python.Runs);
                Assert.Equal(200, all.Select(j => j.Id).Distinct().Count());
            }
        }

        [Fact]
        public void Cancel_a_queued_job_removes_it_and_the_rest_continue()
        {
            using (var rig = new Rig())
            {
                var first = rig.Core.Submit("roslyn", "sleep 600", 10000, null);
                var second = rig.Core.Submit("roslyn", "ok 2", 10000, null);
                var third = rig.Core.Submit("roslyn", "ok 3", 10000, null);
                Assert.True(Wait.Until(() => first.State == JobState.Running));
                rig.Core.Cancel(second);
                Assert.Equal(JobState.Cancelled, second.State);
                Assert.True(rig.Core.WaitFinal(third, 10000));
                Assert.Equal(JobState.Done, third.State);
                Assert.DoesNotContain("ok 2", rig.Roslyn.RunLog);
            }
        }

        [Fact]
        public void Cancel_a_running_job_stops_a_cooperative_script()
        {
            using (var rig = new Rig())
            {
                var spin = rig.Core.Submit("roslyn", "spin", 20000, null);
                Assert.True(Wait.Until(() => spin.State == JobState.Running));
                rig.Core.Cancel(spin);
                Assert.True(rig.Core.WaitFinal(spin, 5000));
                Assert.Equal(JobState.Cancelled, spin.State);
                Assert.Equal(JobState.Done, rig.Run("ok 1").State);
            }
        }

        [Fact]
        public void Cancel_while_compiling_discards_the_job()
        {
            using (var rig = new Rig())
            {
                var slow = rig.Core.Submit("roslyn", "compile-slow 1500", 10000, null);
                Assert.True(Wait.Until(() => slow.State == JobState.Compiling));
                rig.Core.Cancel(slow);
                Assert.True(rig.Core.WaitFinal(slow, 5000));
                Assert.Equal(JobState.Cancelled, slow.State);
                Assert.Equal(0, rig.Roslyn.Runs);
                Assert.Equal(JobState.Done, rig.Run("ok 1").State);
            }
        }

        [Fact]
        public void Timeout_cancels_the_token_and_the_next_job_runs()
        {
            using (var rig = new Rig())
            {
                var j = rig.Run("spin", timeoutMs: 300);
                Assert.Equal(JobState.Timeout, j.State);
                Assert.Equal("timeout", j.ToJson()["error"]["kind"]);
                Assert.Equal(JobState.Done, rig.Run("ok 1").State);
            }
        }

        [Fact]
        public void A_job_that_ignores_cancellation_keeps_later_jobs_queued_until_it_returns()
        {
            using (var rig = new Rig())
            {
                var stuck = rig.Core.Submit("roslyn", "sleep 1500", 300, null);
                var next = rig.Core.Submit("roslyn", "ok 7", 10000, null);
                Assert.True(rig.Core.WaitFinal(stuck, 5000));
                Assert.Equal(JobState.Timeout, stuck.State);
                Assert.Equal(JobState.Queued, next.State);          // still waiting: the UI thread is busy with the stuck job
                var info = rig.Core.Info("4242", 23717);
                Assert.Equal(stuck.Id, (string)info["queue"]["running"]);
                Assert.True((double)info["queue"]["oldestQueuedMs"] > 0);
                Assert.True(rig.Core.WaitFinal(next, 10000));
                Assert.Equal(JobState.Done, next.State);
            }
        }

        [Fact]
        public void A_blocked_UI_thread_is_reported_as_blocked_by_dialog_and_the_job_never_runs()
        {
            using (var rig = new Rig())
            {
                rig.Ui.Blocked = true;
                var j = rig.Run("ok 1", timeoutMs: 400);
                var e = j.ToJson()["error"];
                Assert.Equal(JobState.Timeout, j.State);
                Assert.Equal("blocked_by_dialog", (string)e["kind"]);
                Assert.Contains("Fake modal dialog", (string)e["message"]);
                var info = rig.Core.Info("4242", 23717);
                Assert.Equal("modal", (string)info["ui"]["state"]);
                rig.Ui.Blocked = false;
                Thread.Sleep(200);
                Assert.Equal(0, rig.Roslyn.Runs);                   // the abandoned job is skipped when the UI wakes up
                Assert.Equal(JobState.Done, rig.Run("ok 2").State);
            }
        }

        [Fact]
        public void Sessions_keep_state_per_engine_and_name_and_can_be_dropped()
        {
            using (var rig = new Rig())
            {
                rig.Run("set a 1", session: "s1");
                Assert.Equal("1", (string)rig.Run("get a", session: "s1").ToJson()["value"]["json"]);
                Assert.Equal("(unset)", (string)rig.Run("get a", session: "s2").ToJson()["value"]["json"]);               // other name
                Assert.Equal("(unset)", (string)rig.Run("get a", engine: "ironpython2", session: "s1").ToJson()["value"]["json"]); // other engine
                Assert.Equal("(unset)", (string)rig.Run("get a").ToJson()["value"]["json"]);                               // isolated job
                Assert.True(rig.Core.DropSession("roslyn", "s1"));
                Assert.Equal("(unset)", (string)rig.Run("get a", session: "s1").ToJson()["value"]["json"]);
                Assert.False(rig.Core.DropSession("roslyn", "never"));
            }
        }

        [Fact]
        public void The_engine_gets_the_UI_context_and_the_session()
        {
            using (var rig = new Rig())
                Assert.Equal("uiapp/uidoc/doc/session=s9", (string)rig.Run("ctx", session: "s9").ToJson()["value"]["json"]);
        }

        [Fact]
        public void Output_is_capped_inline_but_available_in_full_through_slices()
        {
            using (var rig = new Rig())
            {
                var j = rig.Run("bigout 600");          // 600 KiB
                var json = j.ToJson();
                Assert.True((bool)json["outputTruncated"]);
                Assert.Equal(Job.MaxInline, ((string)json["output"]).Length);
                Assert.Equal(600 * 1024, (int)json["outputLength"]);
                int total; var slice = j.OutputSlice(500000, 100, out total);
                Assert.Equal(100, slice.Length); Assert.Equal(600 * 1024, total);
                var big = rig.Run("bigout 5000");       // 5 MiB: more than is kept
                Assert.True(big.OutputDropped);
                int t2; big.OutputSlice(0, 10, out t2);
                Assert.Equal(Job.MaxStored, t2);
            }
        }

        [Fact]
        public void Values_are_converted_and_capped_without_silent_truncation()
        {
            using (var rig = new Rig())
            {
                var list = rig.Run("value-list 20000").ToJson()["value"];
                Assert.Equal(20000, ((JArray)list["json"]).Count);                 // used to be cut at 10,000 without a flag
                Assert.False((bool)list["truncated"]);
                var huge = rig.Run("value-list 300000").ToJson()["value"];          // too many items: no JSON, flagged
                Assert.Equal(JTokenType.Null, huge["json"].Type);
                Assert.True((bool)huge["truncated"]);
                var text = rig.Run("value-big 300000").ToJson()["value"];           // text longer than the inline cap
                Assert.True((bool)text["truncated"]);
                Assert.Equal(Job.MaxInline, ((string)text["text"]).Length);
            }
        }

        [Fact]
        public void Console_output_belongs_to_the_job_that_wrote_it()
        {
            using (var rig = new Rig())
            {
                var stop = false;
                var noise = Task.Run(() => { while (!Volatile.Read(ref stop)) { Console.WriteLine("noise from another thread"); Thread.Sleep(1); } });
                var j = rig.Run("console hello");
                Volatile.Write(ref stop, true); noise.Wait();
                Assert.Equal("hello", ((string)j.ToJson()["output"]).Trim());
            }
        }

        [Fact]
        public void Finished_jobs_do_not_keep_compiled_scripts_alive()
        {
            // regression: the last 100 jobs stayed in memory with their whole compilation (25 GB in a benchmark)
            using (var rig = new Rig())
            {
                for (int i = 0; i < 130; i++) rig.Run("hold");
                for (int i = 0; i < 3; i++) { GC.Collect(); GC.WaitForPendingFinalizers(); }
                int alive; lock (rig.Roslyn.Held) alive = rig.Roslyn.Held.Count(w => w.IsAlive);
                Assert.Equal(0, alive);
                Assert.Equal(100, rig.Core.List().Count);             // the history itself is still capped at 100
            }
        }

        [Fact]
        public void Request_id_makes_a_retry_return_the_same_job_without_running_twice()
        {
            using (var rig = new Rig())
            {
                var a = rig.Run("ok 5", requestId: "req-1", client: "tester");
                var b = rig.Core.Submit("roslyn", "ok 5", 10000, null, "tester", "req-1");
                Assert.Same(a, b);
                Assert.Equal(1, rig.Roslyn.Runs);
                Assert.Equal("tester", (string)a.ToJson()["client"]);
            }
        }

        [Fact]
        public void Audit_log_records_code_client_and_result_per_job()
        {
            using (var rig = new Rig(audit: true))
            {
                var j = rig.Run("ok 11", client: "ci-test");
                Assert.True(Wait.Until(() => Directory.Exists(rig.AuditDir) && Directory.GetFiles(rig.AuditDir, "*-" + j.Id + ".json").Length == 1));
                var f = Directory.GetFiles(rig.AuditDir, "*-" + j.Id + ".json").Single();
                var doc = JObject.Parse(File.ReadAllText(f));
                Assert.Equal("ok 11", (string)doc["code"]);
                Assert.Equal("ci-test", (string)doc["client"]);
                Assert.Equal("done", (string)doc["state"]);
                Assert.Contains("-" + Environment.ProcessId + "-", Path.GetFileName(f));
            }
        }

        [Fact]
        public void Info_describes_engines_ui_state_and_queue()
        {
            using (var rig = new Rig())
            {
                var info = rig.Core.Info("4242", 23717);
                Assert.Equal(1, (int)info["protocol"]);
                Assert.Equal(2026, (int)info["revit"]["year"]);
                Assert.Equal("idle", (string)info["ui"]["state"]);
                var engines = ((JArray)info["engines"]).ToDictionary(e => (string)e["name"]);
                Assert.True((bool)engines["roslyn"]["available"]);
                Assert.False((bool)engines["cpython3"]["available"]);
                Assert.Equal("engine not installed", (string)engines["cpython3"]["reason"]);
                Assert.False((bool)engines["broken"]["available"]);
                Assert.Contains("boom", (string)engines["broken"]["reason"]);
                rig.Ui.Hung = true;
                Assert.Equal("unresponsive", (string)rig.Core.Info("4242", 23717)["ui"]["state"]);
                rig.Ui.Hung = false;
                var running = rig.Core.Submit("roslyn", "sleep 400", 10000, null);
                Assert.True(Wait.Until(() => running.State == JobState.Running));
                Assert.Equal("busy", (string)rig.Core.Info("4242", 23717)["ui"]["state"]);
                Assert.True(rig.Core.WaitFinal(running, 5000));
            }
        }

        [Fact]
        public void Job_history_is_capped_and_drops_the_oldest_finished_jobs()
        {
            using (var rig = new Rig())
            {
                var first = rig.Run("ok 0");
                for (int i = 1; i < 120; i++) rig.Run("ok " + i);
                Assert.Null(rig.Core.Get(first.Id));
                Assert.Equal(100, rig.Core.List().Count);
            }
        }

        [Fact]
        public void The_core_asks_the_UI_thread_again_while_a_job_waits()
        {
            using (var rig = new Rig())
            {
                rig.Ui.Blocked = true;
                var j = rig.Core.Submit("roslyn", "ok 1", 3000, null);
                Thread.Sleep(2300);                       // the worker re-raises about once a second
                Assert.True(rig.Ui.Raises >= 3, "raises: " + rig.Ui.Raises);
                rig.Ui.Blocked = false;
                Assert.True(rig.Core.WaitFinal(j, 5000));
                Assert.Equal(JobState.Done, j.State);
            }
        }
    }
}
