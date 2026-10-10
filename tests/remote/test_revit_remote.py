"""Unit tests for tools/remote/revit-remote with a fake transport (no ssh, no VMs). Run: python3 -m unittest discover -s tests/remote -v"""
import importlib.machinery
import importlib.util
import json
import os
import shlex
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
_path = os.path.join(HERE, "..", "..", "tools", "remote", "revit-remote")
_loader = importlib.machinery.SourceFileLoader("revit_remote", _path)
_spec = importlib.util.spec_from_loader("revit_remote", _loader)
rr = importlib.util.module_from_spec(_spec)
_loader.exec_module(rr)

CMDS_2026 = [
    {"name": "DeselectRandomly", "displayName": "Deselect Randomly", "input": "Any", "output": ""},
    {"name": "ExportToPDF", "displayName": "Export To PDF", "input": "", "output": "PDF"},
]


def probe_json(instances, commands, error=None):
    return json.dumps({"hostname": "H", "error": error, "instances": instances, "commands": commands})


def inst(port=23717, pid=100, year=2026, state="idle", dialog=None):
    return {"port": port, "pid": pid, "year": year, "version": "26.1", "build": "b", "state": state, "responsive": True, "dialog": dialog, "queued": 0, "running": None}


class FakeRunner:
    """Records every call; answers from a table of (substring of argv or stdin) -> (rc, stdout, stderr)."""
    def __init__(self, answers=None, default=(0, "", "")):
        self.calls = []
        self.answers = answers or []
        self.default = default

    def __call__(self, argv, stdin_text, timeout):
        self.calls.append({"argv": argv, "stdin": stdin_text, "timeout": timeout})
        joined = " ".join(argv) + "\n" + (stdin_text or "")
        for needle, answer in self.answers:
            if needle in joined:
                return answer
        return self.default


class TargetExpansion(unittest.TestCase):
    def test_parse_vm_listing_ignores_junk(self):
        text = "revit-win 3\nci-20261010-222858 20\n\nnot a vm\nfoo bar\nlast 7\n"
        self.assertEqual(rr.parse_vm_listing(text), [
            {"name": "revit-win", "cid": 3}, {"name": "ci-20261010-222858", "cid": 20}, {"name": "last", "cid": 7}])

    def test_qemu_vsock_target_expands_to_one_target_per_running_vm(self):
        run = FakeRunner([("pgrep", (0, "revit-win 3\nci-1 20\nci-2 21\n", ""))])
        cfg = {"targets": [{"name": "host2", "kind": "qemu-vsock", "ssh": "host2", "guest_key": "~/.ssh/k"}]}
        t = rr.expand_targets(cfg, run)
        self.assertEqual([x["name"] for x in t], ["host2/revit-win", "host2/cid20", "host2/cid21"])
        self.assertEqual(t[1]["cid"], 20)
        self.assertEqual(t[1]["guest_key"], "~/.ssh/k")
        self.assertEqual(t[1]["guest_user"], "user")
        self.assertEqual(run.calls[0]["argv"][0], "ssh")
        self.assertIn("host2", run.calls[0]["argv"])

    def test_unreachable_vm_host_is_a_target_with_a_reason_not_a_crash(self):
        run = FakeRunner([("pgrep", (255, "", "no route"))])
        t = rr.expand_targets({"targets": [{"name": "h", "kind": "qemu-vsock", "ssh": "h"}]}, run)
        self.assertEqual(len(t), 1)
        self.assertIn("cannot list", t[0]["unreachable"])

    def test_plain_ssh_target_passes_through(self):
        t = rr.expand_targets({"targets": [{"name": "pc", "ssh": "pc.local"}]}, FakeRunner())
        self.assertEqual(t, [{"name": "pc", "ssh": "pc.local", "kind": "ssh"}])

    def test_missing_config_is_an_empty_target_list(self):
        self.assertEqual(rr.load_config("/nonexistent/remote.json"), {"targets": []})


class Transport(unittest.TestCase):
    def test_ssh_target_runs_powershell_reading_stdin(self):
        argv = rr.transport_command({"kind": "ssh", "ssh": "pc"})
        self.assertEqual(argv[0], "ssh")
        self.assertIn("pc", argv)
        self.assertEqual(argv[-6:], ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "-"])
        self.assertIn("BatchMode=yes", argv)

    def test_vsock_guest_is_a_nested_ssh_with_a_proxycommand(self):
        t = {"kind": "qemu-vsock-guest", "ssh": "host2", "cid": 20, "guest_key": "~/.ssh/k", "guest_user": "user"}
        argv = rr.transport_command(t)
        self.assertEqual(argv[0], "ssh")
        self.assertIn("host2", argv)
        remote = argv[-1]
        parts = shlex.split(remote)                       # what the remote shell will see
        self.assertEqual(parts[0], "ssh")
        self.assertIn("ProxyCommand=socat - VSOCK-CONNECT:20:22", parts)
        self.assertIn("user@vsock20", parts)
        self.assertEqual(parts[-1], "-")
        self.assertIn("~/.ssh/k", remote)                  # the tilde must reach the remote shell unquoted so it expands there
        self.assertNotIn("'~", remote)

    def test_unknown_kind_is_an_error(self):
        with self.assertRaises(ValueError):
            rr.transport_command({"kind": "carrier-pigeon"})

    def test_script_params_are_one_line_and_apostrophe_safe(self):
        s = rr.script_with_params("invoke.ps1", {"command": "O'Brien", "port": 1})
        first = s.split("\n")[0]
        self.assertTrue(first.startswith("$P = '") and first.endswith("' | ConvertFrom-Json"))
        self.assertIn("O''Brien", first)                   # PowerShell single-quote escaping
        self.assertNotIn("@'", first)                      # no here-strings: they broke over the two ssh hops
        self.assertIn("v1/commands", s)


class Probing(unittest.TestCase):
    T = {"name": "t", "kind": "ssh", "ssh": "t"}

    def test_probe_parses_instances_and_commands(self):
        out = probe_json([inst(23717, 1), inst(23718, 2, 2024)], {"2026": CMDS_2026, "2024": CMDS_2026[:1]})
        pr = rr.probe_target(self.T, FakeRunner(default=(0, out + "\n", "")))
        self.assertEqual(pr["target"], "t")
        self.assertEqual(len(pr["instances"]), 2)
        self.assertEqual(len(pr["commands"]["2026"]), 2)

    def test_powershell_collapses_single_element_arrays(self):
        out = json.dumps({"hostname": "H", "error": None, "instances": inst(), "commands": {"2026": CMDS_2026[0]}})
        pr = rr.probe_target(self.T, FakeRunner(default=(0, out, "")))
        self.assertEqual(len(pr["instances"]), 1)
        self.assertEqual(pr["commands"]["2026"], [CMDS_2026[0]])

    def test_older_agent_without_a_catalogue_yields_no_commands_not_broken_entries(self):
        out = json.dumps({"hostname": "H", "error": None, "instances": [inst()], "commands": {"2026": [{"name": None, "displayName": None}]}})
        pr = rr.probe_target(self.T, FakeRunner(default=(0, out, "")))
        self.assertEqual(pr["commands"]["2026"], [])
        self.assertEqual(rr.build_rows([pr]), [])

    def test_transport_failure_is_reported_per_target(self):
        pr = rr.probe_target(self.T, FakeRunner(default=(255, "", "ssh: connect to host t port 22: Connection refused\n")))
        self.assertIn("Connection refused", pr["error"])
        self.assertEqual(pr["instances"], [])

    def test_timeout_is_reported(self):
        pr = rr.probe_target(self.T, FakeRunner(default=(124, "", "timed out after 30s")))
        self.assertIn("timed out", pr["error"])

    def test_garbage_answer_is_reported(self):
        pr = rr.probe_target(self.T, FakeRunner(default=(0, "{not json", "")))
        self.assertIn("unreadable", pr["error"])
        pr = rr.probe_target(self.T, FakeRunner(default=(0, "welcome banner\n", "")))
        self.assertIn("cannot reach", pr["error"])

    def test_banner_noise_before_the_json_line_is_ignored(self):
        out = "Warning: something\n" + probe_json([inst()], {"2026": CMDS_2026})
        self.assertEqual(len(rr.probe_target(self.T, FakeRunner(default=(0, out, "")))["instances"]), 1)

    def test_unreachable_target_is_not_probed(self):
        run = FakeRunner()
        pr = rr.probe_target({"name": "x", "kind": "qemu-vsock", "unreachable": "cannot list"}, run)
        self.assertEqual(pr["error"], "cannot list")
        self.assertEqual(run.calls, [])

    def test_probe_all_runs_every_target_and_filters_by_substring(self):
        out = probe_json([inst()], {"2026": CMDS_2026})
        run = FakeRunner(default=(0, out, ""))
        targets = [{"name": "a/one", "kind": "ssh", "ssh": "a"}, {"name": "b/two", "kind": "ssh", "ssh": "b"}]
        self.assertEqual({p["target"] for p in rr.probe_all(targets, run)}, {"a/one", "b/two"})
        self.assertEqual([p["target"] for p in rr.probe_all(targets, run, only="two")], ["b/two"])


class Rows(unittest.TestCase):
    def probes(self):
        return [
            {"target": "h/cid20", "error": None, "instances": [inst(23717, 11, 2026), inst(23718, 12, 2024)],
             "commands": {"2026": CMDS_2026, "2024": CMDS_2026[:1]}},
            {"target": "h/cid21", "error": "cannot reach: x", "instances": [], "commands": {}},
            {"target": "h/win", "error": None, "instances": [inst(23717, 7, 2026, "modal", "Autodesk Revit 2026")], "commands": {"2026": CMDS_2026[1:]}},
        ]

    def test_one_row_per_command_per_instance_using_that_years_catalogue(self):
        rows = rr.build_rows(self.probes())
        self.assertEqual(len(rows), 2 + 1 + 1)             # 2026: 2 cmds, 2024: 1 cmd, win 2026: 1 cmd
        ids = [r["id"] for r in rows]
        self.assertEqual(len(set(ids)), len(ids))
        self.assertIn("h/cid20|23718|DeselectRandomly", ids)
        self.assertNotIn("h/cid20|23718|ExportToPDF", ids)

    def test_row_shape_matches_the_pickers(self):
        r = rr.build_rows(self.probes())[0]
        self.assertEqual(set(r), {"id", "preview", "haystack", "target", "port", "command", "fields"})
        self.assertEqual(r["preview"], "Deselect Randomly")
        self.assertEqual(r["command"], "DeselectRandomly")
        self.assertEqual(r["port"], 23717)
        f = r["fields"]
        self.assertEqual(f["command"], "Deselect Randomly")
        self.assertEqual(f["class"], "DeselectRandomly")
        self.assertEqual(f["input"], "Any")
        self.assertEqual(f["target"], "h/cid20")
        self.assertEqual(f["year"], "2026")
        self.assertEqual(f["pid"], "11")
        self.assertEqual(f["instance"], "h/cid20 · 2026 · 11")
        self.assertTrue(all(isinstance(v, str) for v in f.values()))

    def test_haystack_is_lowercase_and_short(self):
        r = rr.build_rows(self.probes())[0]
        self.assertEqual(r["haystack"], r["haystack"].lower())
        self.assertIn("deselect randomly", r["haystack"])
        self.assertNotIn("any", r["haystack"].split())    # `input` text must not leak into free-text matches

    def test_modal_dialog_shows_in_the_state_field(self):
        win = [r for r in rr.build_rows(self.probes()) if r["target"] == "h/win"][0]
        self.assertEqual(win["fields"]["state"], "dialog: Autodesk Revit 2026")

    def test_unreachable_targets_contribute_no_rows(self):
        self.assertFalse([r for r in rr.build_rows(self.probes()) if r["target"] == "h/cid21"])

    def test_cache_round_trip_and_missing_cache(self):
        with tempfile.TemporaryDirectory() as d:
            rr.CACHE = os.path.join(d, "sub", "rows.json")
            self.assertEqual(rr.read_cache()["rows"], [])
            rows = rr.build_rows(self.probes())
            rr.write_cache(rows, self.probes())
            c = rr.read_cache()
            self.assertEqual(c["rows"], rows)
            self.assertEqual(c["errors"], [{"target": "h/cid21", "error": "cannot reach: x"}])
            self.assertGreater(c["time"], 0)
            with open(rr.CACHE, "w") as f:
                f.write("{broken")
            self.assertEqual(rr.read_cache()["rows"], [])


class Invoke(unittest.TestCase):
    T = {"name": "t", "kind": "ssh", "ssh": "t"}

    def test_policy_values(self):
        self.assertIsNone(rr.policy_param(None))
        self.assertIsNone(rr.policy_param("report"))
        self.assertEqual(rr.policy_param("fail"), "fail")
        self.assertEqual(rr.policy_param("respond:OK"), {"respond": "OK"})
        with self.assertRaises(ValueError):
            rr.policy_param("explode")

    def test_invoke_sends_the_parameters_and_returns_the_job(self):
        job = {"id": "j-1", "state": "done", "ok": True}
        run = FakeRunner(default=(0, json.dumps(job) + "\n", ""))
        res = rr.invoke(self.T, 23717, "ExportToPDF", wait=5000, policy="respond:OK", timeout_ms=9000, runner=run)
        self.assertEqual(res, job)
        sent = run.calls[0]["stdin"]
        params = json.loads(sent.split("\n")[0][len("$P = '"):-len("' | ConvertFrom-Json")])
        self.assertEqual(params, {"port": 23717, "command": "ExportToPDF", "wait": 5000, "timeoutMs": 9000, "dialogPolicy": {"respond": "OK"}})
        self.assertGreaterEqual(run.calls[0]["timeout"], 50)

    def test_invoke_without_optional_parameters(self):
        run = FakeRunner(default=(0, '{"state":"queued"}', ""))
        rr.invoke(self.T, 1, "X", runner=run)
        params = json.loads(run.calls[0]["stdin"].split("\n")[0][len("$P = '"):-len("' | ConvertFrom-Json")])
        self.assertEqual(params, {"port": 1, "command": "X", "wait": 0})

    def test_transport_failure_becomes_an_error_object(self):
        res = rr.invoke(self.T, 1, "X", runner=FakeRunner(default=(255, "", "Permission denied (publickey).\n")))
        self.assertEqual(res["error"]["kind"], "transport")
        self.assertIn("Permission denied", res["error"]["message"])

    def test_find_target_exact_then_unique_substring(self):
        ts = [{"name": "host2/cid20"}, {"name": "host2/cid21"}, {"name": "host2/revit-win"}]
        self.assertEqual(rr.find_target(ts, "host2/cid21")["name"], "host2/cid21")
        self.assertEqual(rr.find_target(ts, "win")["name"], "host2/revit-win")
        self.assertIsNone(rr.find_target(ts, "host2"))     # ambiguous
        self.assertIsNone(rr.find_target(ts, "nope"))


class Scripts(unittest.TestCase):
    """The PowerShell files are not run here (no PowerShell on the dev machine); these guard what is easy to break."""
    def read(self, name):
        with open(os.path.join(HERE, "..", "..", "tools", "remote", name)) as f:
            return f.read()

    def test_probe_ends_with_a_single_json_line_and_scans_the_agent_port_range(self):
        s = self.read("probe.ps1")
        self.assertIn("23717", s); self.assertIn("23817", s)
        self.assertIn("ConvertTo-Json -Depth 8 -Compress", s)
        self.assertIn("/v1/commands", s)
        self.assertIn("$c -and $c.commands", s)

    def test_invoke_escapes_the_command_name_and_removes_its_temp_file(self):
        s = self.read("invoke.ps1")
        self.assertIn("EscapeDataString", s)
        self.assertIn("Remove-Item $tmp", s)

    def test_no_here_strings_in_the_scripts(self):
        for n in ("probe.ps1", "invoke.ps1"):
            self.assertNotIn("@'", self.read(n))


if __name__ == "__main__":
    unittest.main()
