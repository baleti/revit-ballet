using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using RevitBallet.Core.Query;

namespace RevitBallet.Core.Remote
{
    /// <summary>What probing one target found.</summary>
    public sealed class ProbeResult
    {
        public string Target;
        public string Error;
        public List<JObject> Instances = new List<JObject>();
        /// <summary>Revit year -> the add-in commands that year's agent lists.</summary>
        public Dictionary<string, List<JObject>> Commands = new Dictionary<string, List<JObject>>();

        public static ProbeResult Failed(string target, string error)
        {
            return new ProbeResult { Target = target, Error = error };
        }

        /// <summary>Parses the one JSON line the probe prints. Noise before it (a banner) is ignored; a failed transport is an error result, never an exception.</summary>
        public static ProbeResult Parse(string target, int exitCode, string stdout, string stderr)
        {
            string line = (stdout ?? "").Split('\n').Select(l => l.Trim()).LastOrDefault(l => l.StartsWith("{"));
            if (exitCode != 0 || line == null)
            {
                string why = exitCode == 124 ? "timed out" : ((stderr ?? "").Trim().Split('\n').LastOrDefault() ?? "no answer").Trim();
                return Failed(target, "cannot reach: " + (why.Length == 0 ? "no answer" : why));
            }
            JObject o;
            try { o = JObject.Parse(line); }
            catch (Exception) { return Failed(target, "unreadable answer"); }
            var r = new ProbeResult { Target = target, Error = (string)o["error"] };
            // PowerShell's ConvertTo-Json turns a one-element array into the element itself
            r.Instances = AsObjects(o["instances"]);
            var commands = o["commands"] as JObject;
            if (commands != null)
                foreach (var p in commands.Properties())
                    // an agent without /v1/commands (an older build) has no catalogue: no commands, not broken entries
                    r.Commands[p.Name] = AsObjects(p.Value).Where(c => !string.IsNullOrEmpty((string)c["name"])).ToList();
            return r;
        }

        private static List<JObject> AsObjects(JToken t)
        {
            if (t == null || t.Type == JTokenType.Null) return new List<JObject>();
            var arr = t as JArray;
            if (arr != null) return arr.OfType<JObject>().ToList();
            var obj = t as JObject;
            return obj != null ? new List<JObject> { obj } : new List<JObject>();
        }
    }

    /// <summary>One line of the picker: a command in one Revit instance.</summary>
    public sealed class RemoteRow
    {
        public string Id;
        public string Preview;
        public string Target;
        public int Port;
        public string Command;
        public Dictionary<string, object> Fields = new Dictionary<string, object>();
    }

    public static class RemoteRows
    {
        public const string FieldCommand = "command", FieldClass = "class", FieldInput = "input", FieldOutput = "output", FieldTarget = "target",
            FieldInstance = "instance", FieldYear = "year", FieldPid = "pid", FieldState = "state";

        /// <summary>The picker's vocabulary. `command` is the row's own text so only input and instance are columns by default.
        /// Free text covers display name, class, target and year only (never `input`, so "any" does not match every command that needs a selection).</summary>
        public static QuerySchema Schema()
        {
            var s = new QuerySchema();
            s.AddField(FieldCommand, false).AddField(FieldClass, false).AddField(FieldInput).AddField(FieldOutput, false).AddField(FieldTarget, false)
             .AddField(FieldInstance).AddField(FieldYear, false).AddField(FieldPid, false).AddField(FieldState, false);
            s.Haystack = row => ((string)row["haystack"] ?? "");
            return s;
        }

        public static string InstanceLabel(string target, JObject inst)
        {
            return target + " · " + (string)inst["year"] + " · " + (string)inst["pid"];
        }

        /// <summary>One row per (command x instance), using the catalogue of that instance's Revit year.</summary>
        public static List<RemoteRow> Build(IEnumerable<ProbeResult> probes)
        {
            var rows = new List<RemoteRow>();
            foreach (var pr in probes)
            {
                foreach (var inst in pr.Instances)
                {
                    string year = (string)inst["year"];
                    string state = (string)inst["state"] ?? "";
                    string dialog = (string)inst["dialog"];
                    if (!string.IsNullOrEmpty(dialog)) state = "dialog: " + dialog;
                    List<JObject> commands;
                    if (!pr.Commands.TryGetValue(year, out commands)) continue;
                    foreach (var c in commands)
                    {
                        string name = (string)c["name"], display = (string)c["displayName"] ?? name;
                        var row = new RemoteRow { Id = pr.Target + "|" + (int)inst["port"] + "|" + name, Preview = display, Target = pr.Target, Port = (int)inst["port"], Command = name };
                        row.Fields[FieldCommand] = display; row.Fields[FieldClass] = name;
                        row.Fields[FieldInput] = (string)c["input"] ?? ""; row.Fields[FieldOutput] = (string)c["output"] ?? "";
                        row.Fields[FieldTarget] = pr.Target; row.Fields[FieldInstance] = InstanceLabel(pr.Target, inst);
                        row.Fields[FieldYear] = year; row.Fields[FieldPid] = (string)inst["pid"]; row.Fields[FieldState] = state;
                        row.Fields["haystack"] = (display + " " + name + " " + pr.Target + " " + year).ToLowerInvariant();
                        rows.Add(row);
                    }
                }
            }
            return rows;
        }
    }

    /// <summary>How a finished invoke ended, in words for a notification.</summary>
    public static class RemoteOutcome
    {
        public static void Describe(JObject job, out string word, out string detail)
        {
            var err = job["error"] as JObject;
            string state = (string)job["state"];
            string kind = err == null ? null : (string)err["kind"];
            if (state == "done") { word = "done"; detail = ((string)job["output"] ?? "").Trim(); return; }
            if (state == "cancelled" || kind == "cancelled") { word = "cancelled"; detail = ""; return; }
            if (state == "queued" || state == "compiling" || state == "running") { word = "still running"; detail = "it did not finish while the picker waited"; return; }
            if (kind == "blocked_by_dialog")
            {
                var d = err["dialog"] as JObject;
                word = "blocked by a dialog";
                detail = (d == null ? "" : (string)d["title"]) + ": " + (d == null ? "" : ((string)d["text"] ?? "").Trim());
                return;
            }
            word = "failed";
            detail = (err == null ? null : (string)err["message"]) ?? ((string)job["output"] ?? "").Trim();
            if (string.IsNullOrEmpty(detail)) detail = "unknown error";
        }
    }
}
