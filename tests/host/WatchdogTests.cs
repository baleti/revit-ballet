using System;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using Xunit;

namespace RevitBallet.Agent.Tests
{
    /// <summary>A compiler that hangs must not take the queue with it, and the caller must be told what happened.</summary>
    public class WatchdogTests
    {
        [Fact]
        public void Deadlocked_compile_is_reported_as_engine_hung_and_the_queue_keeps_moving()
        {
            using (var rig = new Rig())
            {
                rig.Core.CompileStallMs = 600;
                var hung = rig.Run("compile-hang", waitMs: 20000);
                Assert.Equal("failed", Job(hung)["state"]);
                Assert.Equal("engine_hung", (string)Job(hung)["error"]["kind"]);
                Assert.Contains("no progress", (string)Job(hung)["error"]["message"]);
                Assert.Equal(1, rig.Source.Resets);
                var next = rig.Run("ok 7");
                Assert.True(next.State == RevitBallet.Commands.Agent.JobState.Done, "the next job must run");
                Assert.Equal(1, (int)rig.Core.Info("4242", 1)["hungCompiles"]);
            }
        }

        [Fact]
        public void Endless_loop_in_the_compiler_is_cut_off_by_the_cap()
        {
            using (var rig = new Rig())
            {
                rig.Core.CompileStallMs = 600000; rig.Core.CompileCapMs = 800;
                var sw = Stopwatch.StartNew();
                var hung = rig.Run("compile-spin", timeoutMs: 300, waitMs: 20000);
                Assert.Equal("engine_hung", (string)Job(hung)["error"]["kind"]);
                Assert.Contains("still running", (string)Job(hung)["error"]["message"]);
                Assert.True(sw.ElapsedMilliseconds < 10000);
                Assert.True(rig.Run("ok 1").State == RevitBallet.Commands.Agent.JobState.Done);
            }
        }

        [Fact]
        public void A_long_compile_that_keeps_working_is_not_called_hung()
        {
            using (var rig = new Rig())
            {
                rig.Core.CompileStallMs = 700; rig.Core.CompileCapMs = 60000;
                var job = rig.Run("compile-work 3000", timeoutMs: 10000, waitMs: 20000);   // 4x longer than the stall limit, but always using CPU
                Assert.Equal("done", (string)Job(job)["state"]);
                Assert.Equal(0, rig.Source.Resets);
            }
        }

        private static JObject Job(RevitBallet.Commands.Agent.Job j) { return j.ToJson(); }
    }
}
