using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using RevitBallet.Commands.Agent;
using Xunit;

namespace RevitBallet.Agent.Tests
{
    /// <summary>POST /roslyn (raw C# body) on the job queue, in the old response shape that scripts and docs still use.</summary>
    public class LegacyRoslynTests
    {
        [Fact]
        public async Task Success_has_output_and_the_return_value_appended()
        {
            using (var rig = new Rig())
            {
                var r = await LegacyRoslyn.HandleAsync(rig.Core, "ok 3");
                var j = JObject.Parse(r.Json);
                Assert.Equal(200, r.Status);
                Assert.True((bool)j["Success"]);
                Assert.Contains("Return value: 3", (string)j["Output"]);
                Assert.True(((JArray)j["ProcessingLog"]).Count >= 2);
            }
        }

        [Fact]
        public async Task Compile_errors_keep_the_old_shape()
        {
            using (var rig = new Rig())
            {
                var r = await LegacyRoslyn.HandleAsync(rig.Core, "compile-error");
                var j = JObject.Parse(r.Json);
                Assert.Equal(400, r.Status);
                Assert.False((bool)j["Success"]);
                Assert.Equal("Compilation failed", (string)j["Error"]);
                Assert.Contains("(1,2): error E1: bad code", (string)j["Diagnostics"][0]);
            }
        }

        [Fact]
        public async Task Runtime_errors_carry_message_and_stack_and_empty_bodies_are_rejected()
        {
            using (var rig = new Rig())
            {
                var e = JObject.Parse((await LegacyRoslyn.HandleAsync(rig.Core, "throw kaboom")).Json);
                Assert.Contains("kaboom", (string)e["Error"]);
                Assert.Contains("Stack Trace", (string)e["Error"]);
                var empty = await LegacyRoslyn.HandleAsync(rig.Core, "  ");
                Assert.Equal(400, empty.Status);
                Assert.Equal("Empty request body", (string)JObject.Parse(empty.Json)["Error"]);
            }
        }

        [Fact]
        public void Info_carries_the_agent_build_block_the_runtime_sets()
        {
            using (var rig = new Rig())
            {
                rig.Core.AgentInfo = new JObject { ["build"] = "1.2.3 abc", ["loadedFrom"] = "hot-reload" };
                var info = rig.Core.Info("4242", 1);
                Assert.Equal("1.2.3 abc", (string)info["agent"]["build"]);
                Assert.Equal("hot-reload", (string)info["agent"]["loadedFrom"]);
            }
        }
    }
}
