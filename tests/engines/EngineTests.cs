using System;
using System.Collections;
using System.IO;
using System.Linq;
using RevitBallet.Agents;
using Xunit;

namespace RevitBallet.Engines.Tests
{
    public class ManifestTests
    {
        [Theory, InlineData("roslyn"), InlineData("ironpython2"), InlineData("cpython3")]
        public void Each_engine_ships_a_valid_manifest_and_its_assembly(string name)
        {
            var m = EngineHarness.Manifest(name);
            Assert.Equal(name, m.Name);
            Assert.Equal(1, m.Protocol);
            Assert.True(File.Exists(Path.Combine(EngineHarness.EngineDir(name), m.Assembly)), m.Assembly + " missing");
        }

        [Theory, InlineData("roslyn"), InlineData("ironpython2"), InlineData("cpython3")]
        public void No_Autodesk_binaries_are_shipped_with_an_engine(string name)
        {
            var bad = Directory.GetFiles(EngineHarness.EngineDir(name), "*", SearchOption.AllDirectories)
                .Select(Path.GetFileName).Where(f => f.StartsWith("RevitAPI", StringComparison.OrdinalIgnoreCase) || f.StartsWith("AdWindows", StringComparison.OrdinalIgnoreCase) || f.StartsWith("UIFramework", StringComparison.OrdinalIgnoreCase)).ToList();
            Assert.Empty(bad);
        }
    }

    /// <summary>The same behaviours for every engine; code differs only by language.</summary>
    public class EngineBehaviourTests
    {
        public static System.Collections.Generic.IEnumerable<object[]> Engines()
        {
            yield return new object[] { "roslyn" };
            yield return new object[] { "ironpython2" };
            yield return new object[] { "cpython3" };
        }

        private static IScriptEngine Engine(string name)
        {
            Skip.If(name == "cpython3" && !EngineHarness.CPythonRuntimeAvailable, "the bundled Python runtime can only be loaded on Windows (run fetch-python first)");
            return EngineHarness.Load(name);
        }
        private static bool Py(string name) => name != "roslyn";

        [SkippableTheory, MemberData(nameof(Engines))]
        public void Value_and_output(string name)
        {
            var e = Engine(name);
            var o = EngineHarness.Run(e, Py(name) ? "print('hi')\nout.write('there\\n')\nresult = 6 * 7" : "Console.WriteLine(\"unused\"); Out.WriteLine(\"hi\"); Out.WriteLine(\"there\"); return 6 * 7;");
            Assert.True(o.Ok, o.Run?.ErrorMessage ?? o.FirstError?.Message);
            Assert.Equal(42L, Convert.ToInt64(o.Value));
            Assert.Contains("hi", o.Output); Assert.Contains("there", o.Output);
        }

        [SkippableTheory, MemberData(nameof(Engines))]
        public void Compile_errors_carry_a_line_number(string name)
        {
            var e = Engine(name);
            var o = EngineHarness.Run(e, Py(name) ? "x = 1\ny = (" : "int a = 1;\nint x = ;");
            Assert.False(o.Compile.Success);
            Assert.True(o.FirstError != null && o.FirstError.Line >= 1, "diagnostic line: " + o.FirstError?.Line);
        }

        [SkippableTheory, MemberData(nameof(Engines))]
        public void Runtime_errors_report_message_and_stack(string name)
        {
            var e = Engine(name);
            var o = EngineHarness.Run(e, Py(name) ? "def f():\n    raise ValueError('bad thing')\nf()" : "throw new System.InvalidOperationException(\"bad thing\");");
            Assert.True(o.Compile.Success);
            Assert.False(o.Run.Ok);
            Assert.Equal("runtime", o.Run.ErrorKind);
            Assert.Contains("bad thing", o.Run.ErrorMessage);
            Assert.False(string.IsNullOrEmpty(o.Run.ErrorStack));
        }

        [SkippableTheory, MemberData(nameof(Engines))]
        public void Sessions_carry_state_between_jobs_and_reset(string name)
        {
            var e = Engine(name);
            var s = new EngineSession { Name = "t" };
            Assert.True(EngineHarness.Run(e, Py(name) ? "counter = 5" : "var counter = 5;", s).Ok);
            var o = EngineHarness.Run(e, Py(name) ? "result = counter + 1" : "return counter + 1;", s);
            Assert.True(o.Ok, o.Run?.ErrorMessage ?? o.FirstError?.Message);
            Assert.Equal(6L, Convert.ToInt64(o.Value));
            e.ResetSession(s);
            var after = EngineHarness.Run(e, Py(name) ? "result = counter" : "return counter;", s);
            Assert.False(after.Ok);
        }

        [SkippableTheory, MemberData(nameof(Engines))]
        public void Scripts_see_the_Revit_globals(string name)
        {
            var e = Engine(name);
            var o = EngineHarness.Run(e, Py(name) ? "result = doc.Title + '/' + uiapp.VersionNumber" : "return Doc.Title + \"/\" + UIApp.VersionNumber;");
            Assert.True(o.Ok, o.Run?.ErrorMessage ?? o.FirstError?.Message);
            Assert.Equal("stub document/2026", Convert.ToString(o.Value));
        }

        [SkippableTheory, MemberData(nameof(Engines))]
        public void A_script_that_checks_the_cancellation_token_stops(string name)
        {
            var e = Engine(name);
            var code = Py(name) ? "import time\nwhile not cancel.is_set():\n    time.sleep(0.01)\nresult = 1" : "while (!Cancel.IsCancellationRequested) { System.Threading.Thread.Sleep(10); } return 1;";
            var o = EngineHarness.Run(e, code, cancelAfterMs: 400, hardLimitMs: 15000);
            Assert.True(o.Compile.Success);
            Assert.NotNull(o.Run);   // returned instead of hanging
        }

        [SkippableTheory, InlineData("ironpython2"), InlineData("cpython3")]
        public void Python_engines_stop_a_hot_loop_without_cooperation(string name)
        {
            var e = Engine(name);
            var o = EngineHarness.Run(e, "while True:\n    pass", cancelAfterMs: 500, hardLimitMs: 15000);
            Assert.False(o.Run.Ok);
            Assert.Equal("cancelled", o.Run.ErrorKind);
        }

        [SkippableTheory, InlineData("ironpython2"), InlineData("cpython3")]
        public void Python_engines_import_the_standard_library(string name)
        {
            var e = Engine(name);
            var o = EngineHarness.Run(e, "import os, re, json, collections\nresult = json.dumps({'a': 1})");
            Assert.True(o.Ok, o.Run?.ErrorMessage ?? o.FirstError?.Message);
            Assert.Equal("{\"a\": 1}", Convert.ToString(o.Value));
        }

        [SkippableTheory, MemberData(nameof(Engines))]
        public void Non_ascii_text_survives(string name)
        {
            var e = Engine(name);
            var code = Py(name) ? "result = u'h\\u00e9llo \\u20ac'" : "return \"h\\u00e9llo \\u20ac\";";
            var o = EngineHarness.Run(e, code);
            Assert.True(o.Ok, o.Run?.ErrorMessage ?? o.FirstError?.Message);
            Assert.Equal("h\u00e9llo \u20ac", Convert.ToString(o.Value));
        }

        [Fact]
        public void Roslyn_scripts_can_use_common_framework_types()
        {
            // on .NET Framework a script used to see only mscorlib and System.Core: System.Diagnostics.Process did not even compile
            var e = EngineHarness.Load("roslyn");
            var o = EngineHarness.Run(e, "return System.Diagnostics.Process.GetCurrentProcess().Id > 0 && new System.Xml.XmlDocument() != null;");
            Assert.True(o.Ok, o.Run?.ErrorMessage ?? o.FirstError?.Message);
            Assert.Equal(true, o.Value);
        }

        [SkippableTheory, MemberData(nameof(Engines))]
        public void A_script_can_be_compiled_and_run_many_times(string name)
        {
            var e = Engine(name);
            for (int i = 0; i < 5; i++)
            {
                var o = EngineHarness.Run(e, Py(name) ? "result = " + i : "return " + i + ";");
                Assert.True(o.Ok); Assert.Equal((long)i, Convert.ToInt64(o.Value));
            }
        }
    }
}
