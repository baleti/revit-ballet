using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using RevitBallet.Core.Query;
using RevitBallet.Core.Remote;
using Xunit;

namespace RevitBallet.Core.Tests
{
    public class RemoteConfigTests
    {
        [Fact]
        public void DefaultsWhenEmptyOrBroken()
        {
            foreach (var json in new[] { null, "", "   ", "{not json", "[]" })
            {
                var c = RemoteConfig.Parse(json);
                Assert.Equal("Win+`", c.Hotkey);
                Assert.True(c.Local);
                Assert.Empty(c.Targets);
            }
        }

        [Fact]
        public void ParsesHotkeyLocalAndTargets()
        {
            var c = RemoteConfig.Parse("{\"hotkey\":\"Alt+`\",\"local\":false,\"targets\":[{\"name\":\"vm\",\"kind\":\"ssh\",\"ssh\":\"user@vm\"},{\"name\":\"two\",\"ssh\":\"two\"}]}");
            Assert.Equal("Alt+`", c.Hotkey);
            Assert.False(c.Local);
            Assert.Equal(new[] { "vm", "two" }, c.Targets.Select(t => t.Name).ToArray());
            Assert.All(c.Targets, t => { Assert.Equal("ssh", t.Kind); Assert.Null(t.Unreachable); });
        }

        [Fact]
        public void BadTargetsAreListedButMarkedUnreachable()
        {
            var c = RemoteConfig.Parse("{\"targets\":[{\"name\":\"a\"},{\"name\":\"b\",\"ssh\":\"-oProxyCommand=evil\"},{\"name\":\"c\",\"ssh\":\"x y\"},{\"name\":\"d\",\"kind\":\"qemu-vsock\",\"ssh\":\"h\"},{\"ssh\":\"nameless\"}]}");
            Assert.Equal(new[] { "a", "b", "c", "d" }, c.Targets.Select(t => t.Name).ToArray());       // the nameless one is dropped
            Assert.All(c.Targets, t => Assert.NotNull(t.Unreachable));
            Assert.Contains("start with '-'", c.Targets[1].Unreachable);
            Assert.Contains("not supported", c.Targets[3].Unreachable);
        }

        [Fact]
        public void AllTargetsPutsThisMachineFirstUnlessDisabled()
        {
            var c = RemoteConfig.Parse("{\"targets\":[{\"name\":\"vm\",\"ssh\":\"vm\"}]}");
            Assert.Equal(new[] { "local", "vm" }, c.AllTargets().Select(t => t.Name).ToArray());
            c.Local = false;
            Assert.Equal(new[] { "vm" }, c.AllTargets().Select(t => t.Name).ToArray());
        }

        [Fact]
        public void TransportCommands()
        {
            var local = RemoteTransport.Command(new RemoteTarget { Name = "local", Kind = "local" });
            Assert.Equal("powershell.exe", local.FileName);
            Assert.EndsWith("-Command -", local.Arguments);
            var ssh = RemoteTransport.Command(new RemoteTarget { Name = "vm", Kind = "ssh", Ssh = "user@vm" });
            Assert.Equal("ssh.exe", ssh.FileName);
            Assert.Contains("BatchMode=yes", ssh.Arguments);
            Assert.Contains(" user@vm powershell ", ssh.Arguments);
            Assert.EndsWith("-Command -", ssh.Arguments);
            Assert.Throws<System.InvalidOperationException>(() => RemoteTransport.Command(new RemoteTarget { Name = "x", Kind = "ssh", Unreachable = "nope" }));
        }
    }

    public class RemoteScriptTests
    {
        [Fact]
        public void EmbeddedScriptsAreTheSharedOnes()
        {
            var probe = RemoteScripts.Load("probe.ps1");
            Assert.Contains("23717", probe); Assert.Contains("/v1/commands", probe); Assert.Contains("ConvertTo-Json -Depth 8 -Compress", probe);
            var invoke = RemoteScripts.Load("invoke.ps1");
            Assert.Contains("EscapeDataString", invoke); Assert.Contains("Remove-Item $tmp", invoke);
            Assert.Throws<System.InvalidOperationException>(() => RemoteScripts.Load("nope.ps1"));
        }

        [Fact]
        public void ParametersAreOneLineAndApostropheSafe()
        {
            var s = RemoteScripts.WithParams("invoke.ps1", new JObject { ["command"] = "O'Brien", ["port"] = 23717, ["wait"] = 0 });
            string first = s.Split('\n')[0];
            Assert.StartsWith("$P = '", first);
            Assert.EndsWith("' | ConvertFrom-Json", first);
            Assert.Contains("O''Brien", first);
            Assert.DoesNotContain("@'", first);
            Assert.Contains("v1/commands", s);
        }
    }

    public class RemoteProbeTests
    {
        private static string Json(object instances, object commands, string error = null)
        {
            return new JObject { ["hostname"] = "H", ["error"] = error, ["instances"] = JToken.FromObject(instances), ["commands"] = JToken.FromObject(commands) }.ToString(Newtonsoft.Json.Formatting.None);
        }
        private static object Inst(int port = 23717, int pid = 100, int year = 2026, string state = "idle", string dialog = null)
        {
            return new { port, pid, year, version = "26.1", build = "b", state, responsive = true, dialog, queued = 0, running = (string)null };
        }
        private static readonly object[] Cmds = {
            new { name = "DeselectRandomly", displayName = "Deselect Randomly", input = "Any", output = "" },
            new { name = "ExportToPDF", displayName = "Export To PDF", input = "", output = "PDF" } };

        [Fact]
        public void ParsesInstancesAndCommands()
        {
            var r = ProbeResult.Parse("t", 0, Json(new[] { Inst(23717, 1), Inst(23718, 2, 2024) }, new Dictionary<string, object> { { "2026", Cmds }, { "2024", Cmds.Take(1).ToArray() } }), "");
            Assert.Null(r.Error);
            Assert.Equal(2, r.Instances.Count);
            Assert.Equal(2, r.Commands["2026"].Count);
            Assert.Single(r.Commands["2024"]);
        }

        [Fact]
        public void PowerShellCollapsedSingleElementArraysAreHandled()
        {
            var r = ProbeResult.Parse("t", 0, Json(Inst(), new Dictionary<string, object> { { "2026", Cmds[0] } }), "");
            Assert.Single(r.Instances);
            Assert.Single(r.Commands["2026"]);
        }

        [Fact]
        public void AnOlderAgentWithoutACatalogueGivesNoCommandsNotBrokenEntries()
        {
            var r = ProbeResult.Parse("t", 0, Json(new[] { Inst() }, new Dictionary<string, object> { { "2026", new[] { new { name = (string)null, displayName = (string)null } } } }), "");
            Assert.Empty(r.Commands["2026"]);
            Assert.Empty(RemoteRows.Build(new[] { r }));
        }

        [Theory]
        [InlineData(255, "", "ssh: connect to host t port 22: Connection refused\n", "Connection refused")]
        [InlineData(124, "", "", "timed out")]
        [InlineData(0, "welcome banner\n", "", "cannot reach")]
        [InlineData(0, "{not json", "", "unreadable")]
        public void FailuresAreResultsNotExceptions(int code, string stdout, string stderr, string expected)
        {
            var r = ProbeResult.Parse("t", code, stdout, stderr);
            Assert.Contains(expected, r.Error);
            Assert.Empty(r.Instances);
        }

        [Fact]
        public void NoiseBeforeTheJsonLineIsIgnored()
        {
            var r = ProbeResult.Parse("t", 0, "Warning: something\r\n" + Json(new[] { Inst() }, new Dictionary<string, object> { { "2026", Cmds } }) + "\r\n", "");
            Assert.Single(r.Instances);
        }

        [Fact]
        public void TheProbeReportsItsOwnError()
        {
            var r = ProbeResult.Parse("t", 0, Json(new object[0], new Dictionary<string, object>(), "no agent token on this machine"), "");
            Assert.Equal("no agent token on this machine", r.Error);
        }
    }

    public class RemoteRowsTests
    {
        private static List<RemoteRow> Rows()
        {
            var cmds2026 = new[] { Cmd("DeselectRandomly", "Deselect Randomly", "Any", ""), Cmd("ExportToPDF", "Export To PDF", "", "PDF") };
            var a = new ProbeResult { Target = "h/cid20" };
            a.Instances.Add(Inst(23717, 5728, 2026, "idle", null)); a.Instances.Add(Inst(23718, 12, 2024, "idle", null));
            a.Commands["2026"] = cmds2026.ToList(); a.Commands["2024"] = cmds2026.Take(1).ToList();
            var b = new ProbeResult { Target = "h/win" };
            b.Instances.Add(Inst(23721, 17864, 2023, "modal", "Autodesk Revit 2023"));
            b.Commands["2023"] = cmds2026.Skip(1).ToList();
            var c = ProbeResult.Failed("h/down", "cannot reach: x");
            return RemoteRows.Build(new[] { a, b, c });
        }
        private static JObject Cmd(string name, string display, string input, string output) { return new JObject { ["name"] = name, ["displayName"] = display, ["input"] = input, ["output"] = output }; }
        private static JObject Inst(int port, int pid, int year, string state, string dialog) { return new JObject { ["port"] = port, ["pid"] = pid, ["year"] = year, ["state"] = state, ["dialog"] = dialog }; }
        private static string[] Ids(QueryResult r) { return r.Rows.Select(x => ((string)x["target"]).Replace("h/", "") + ":" + x["class"]).ToArray(); }
        private static QueryResult Q(string q) { return QueryEngine.Run(Rows().Select(r => (IDictionary<string, object>)r.Fields).ToList(), RemoteRows.Schema(), q); }

        [Fact]
        public void OneRowPerCommandPerInstanceUsingThatYearsCatalogue()
        {
            var rows = Rows();
            Assert.Equal(4, rows.Count);                      // 2026: 2, 2024: 1, win 2023: 1; the unreachable target adds none
            Assert.Equal(rows.Count, rows.Select(r => r.Id).Distinct().Count());
            Assert.Contains(rows, r => r.Id == "h/cid20|23718|DeselectRandomly");
            Assert.DoesNotContain(rows, r => r.Id == "h/cid20|23718|ExportToPDF");
        }

        [Fact]
        public void RowShapeAndFields()
        {
            var r = Rows()[0];
            Assert.Equal("Deselect Randomly", r.Preview);
            Assert.Equal("DeselectRandomly", r.Command);
            Assert.Equal(23717, r.Port);
            Assert.Equal("Any", r.Fields["input"]);
            Assert.Equal("h/cid20 · 2026 · 5728", r.Fields["instance"]);
            Assert.Equal("idle", r.Fields["state"]);
            Assert.Equal("dialog: Autodesk Revit 2023", Rows().Single(x => x.Target == "h/win").Fields["state"]);
        }

        [Fact]
        public void DefaultColumnsAreInputAndInstance()
        {
            Assert.Equal(new[] { "input", "instance" }, Q("").Columns.ToArray());
            Assert.Equal(new[] { "input", "instance", "command" }, Q("/at command").Columns.ToArray());
        }

        [Fact]
        public void FreeTextCoversNamesTargetAndYearButNotInput()
        {
            Assert.Equal(2, Q("pdf").Rows.Count);                         // cid20's and win's ExportToPDF
            Assert.Equal(new[] { "win:ExportToPDF" }, Ids(Q("2023")));
            Assert.Empty(Q("pdf deselect").Rows);                         // words AND: no row is both
            Assert.Empty(Q("any").Rows);                                  // "Any" is an input value: only //input any finds it
            Assert.Single(Q("cid20 pdf").Rows);
        }

        [Fact]
        public void ScopedFiltersSortsAndNegation()
        {
            Assert.Equal(2, Q("//input any").Rows.Count);
            Assert.Equal(new[] { "cid20:DeselectRandomly", "cid20:ExportToPDF", "cid20:DeselectRandomly" }.OrderBy(x => x).ToArray(), Ids(Q("//target cid20")).OrderBy(x => x).ToArray());
            Assert.Equal(new[] { "win:ExportToPDF" }, Ids(Q("//state dialog")));
            Assert.Equal(3, Q("!//state dialog").Rows.Count);
            Assert.Equal("12", (string)Q("/s pid").Rows[0]["pid"]);       // numeric, not lexicographic: 12 < 5728 < 17864
            Assert.Equal("17864", (string)Q("/s pid desc").Rows[0]["pid"]);
            Assert.Empty(Q("//zzz foo").Rows);
            Assert.Equal(4, Q("//").Rows.Count);                          // half-typed stays inert
        }
    }

    public class RemoteOutcomeTests
    {
        private static string[] D(string json) { string w, d; RemoteOutcome.Describe(JObject.Parse(json), out w, out d); return new[] { w, d }; }

        [Fact]
        public void Outcomes()
        {
            Assert.Equal(new[] { "done", "ok" }, D("{\"state\":\"done\",\"output\":\"ok\\r\\n\"}"));
            Assert.Equal(new[] { "cancelled", "" }, D("{\"state\":\"cancelled\"}"));
            Assert.Equal(new[] { "cancelled", "" }, D("{\"state\":\"failed\",\"error\":{\"kind\":\"cancelled\"}}"));
            Assert.Equal("still running", D("{\"state\":\"running\"}")[0]);
            Assert.Equal(new[] { "failed", "No active document." }, D("{\"state\":\"failed\",\"error\":{\"kind\":\"runtime\",\"message\":\"No active document.\"}}"));
            Assert.Equal(new[] { "blocked by a dialog", "T: stuck" }, D("{\"state\":\"failed\",\"error\":{\"kind\":\"blocked_by_dialog\",\"dialog\":{\"title\":\"T\",\"text\":\"stuck \"}}}"));
            Assert.Equal(new[] { "failed", "boom" }, D("{\"error\":{\"kind\":\"transport\",\"message\":\"boom\"}}"));
            Assert.Equal("failed", D("{}")[0]);
        }
    }
}

namespace RevitBallet.Core.Tests
{
    public class HotkeySpecTests
    {
        [Theory]
        [InlineData("Win+`", 8u, "`")]
        [InlineData("Alt+`", 1u, "`")]
        [InlineData("Ctrl+Alt+K", 3u, "K")]
        [InlineData("ctrl+alt+k", 3u, "K")]
        [InlineData("Win+Shift+F9", 12u, "F9")]
        [InlineData("Super+Space", 8u, "Space")]
        [InlineData("Mod+`", 8u, "`")]
        [InlineData("Ctrl++", 2u, "+")]
        [InlineData(" Win + ` ", 8u, "`")]
        public void ParsesModifiersAndKey(string text, uint mods, string key)
        {
            HotkeySpec h;
            Assert.True(HotkeySpec.TryParse(text, out h));
            Assert.Equal(mods, h.Modifiers);
            Assert.Equal(key, h.Key);
        }

        [Theory]
        [InlineData(null)] [InlineData("")] [InlineData("   ")] [InlineData("`")] [InlineData("K")] [InlineData("Win+")] [InlineData("Hyper+K")]
        [InlineData("Win+Banana")] [InlineData("Win+F25")] [InlineData("Win+F0")]
        public void RejectsWhatItCannotRegister(string text)
        {
            HotkeySpec h;
            Assert.False(HotkeySpec.TryParse(text, out h));
            Assert.Null(h);
        }

        [Fact]
        public void BacktickAndRoundTrip()
        {
            HotkeySpec h;
            Assert.True(HotkeySpec.TryParse("Win+`", out h));
            Assert.True(h.IsBacktick);
            Assert.Equal("Win+`", h.ToString());
            Assert.True(HotkeySpec.TryParse("Alt+Ctrl+Win+Shift+F9", out h));
            Assert.Equal("Win+Ctrl+Alt+Shift+F9", h.ToString());
            Assert.False(HotkeySpec.TryParse("Win+K", out h) && h.IsBacktick);
        }
    }
}
