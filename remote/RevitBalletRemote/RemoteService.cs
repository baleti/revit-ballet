using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RevitBallet.Core.Remote;

namespace RevitBalletRemote
{
    /// <summary>Finds the Revit instances on the configured targets, caches what it found and runs commands in them.</summary>
    internal sealed class RemoteService
    {
        public static string ConfigPath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "revit-ballet", "remote.json"); } }
        private static string CachePath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "revit-ballet", "remote-rows.json"); } }

        public RemoteConfig Config { get; private set; }

        public RemoteService() { Reload(); }

        public void Reload()
        {
            string json = null;
            try { if (File.Exists(ConfigPath)) json = File.ReadAllText(ConfigPath); } catch (Exception) { }
            Config = RemoteConfig.Parse(json);
        }

        public sealed class Snapshot
        {
            public List<RemoteRow> Rows = new List<RemoteRow>();
            public List<string> Errors = new List<string>();
            public DateTime Time;
        }

        public Snapshot LoadCache()
        {
            try
            {
                var o = JObject.Parse(File.ReadAllText(CachePath));
                var snap = new Snapshot { Time = (DateTime?)o["time"] ?? DateTime.MinValue };
                foreach (var r in (JArray)o["rows"])
                {
                    var row = new RemoteRow { Id = (string)r["id"], Preview = (string)r["preview"], Target = (string)r["target"], Port = (int)r["port"], Command = (string)r["command"] };
                    foreach (var p in ((JObject)r["fields"]).Properties()) row.Fields[p.Name] = (string)p.Value;
                    snap.Rows.Add(row);
                }
                return snap;
            }
            catch (Exception) { return new Snapshot(); }
        }

        private void SaveCache(Snapshot snap)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(CachePath));
                var o = new JObject { ["time"] = snap.Time, ["rows"] = new JArray(snap.Rows.Select(r => new JObject { ["id"] = r.Id, ["preview"] = r.Preview, ["target"] = r.Target, ["port"] = r.Port, ["command"] = r.Command, ["fields"] = JObject.FromObject(r.Fields) })) };
                File.WriteAllText(CachePath + ".tmp", o.ToString(Formatting.None));
                if (File.Exists(CachePath)) File.Delete(CachePath);
                File.Move(CachePath + ".tmp", CachePath);
            }
            catch (Exception) { /* the cache is an optimisation */ }
        }

        /// <summary>Probes every target in parallel (a few seconds for ssh targets) and returns the merged rows.</summary>
        public Snapshot Refresh()
        {
            Reload();
            var targets = Config.AllTargets();
            var tasks = targets.Select(t => Task.Run(() => Probe(t))).ToArray();
            Task.WaitAll(tasks);
            var probes = tasks.Select(t => t.Result).ToList();
            var snap = new Snapshot { Time = DateTime.UtcNow, Rows = RemoteRows.Build(probes), Errors = probes.Where(p => !string.IsNullOrEmpty(p.Error)).Select(p => p.Target + ": " + p.Error).ToList() };
            SaveCache(snap);
            return snap;
        }

        private static ProbeResult Probe(RemoteTarget t)
        {
            if (t.Unreachable != null) return ProbeResult.Failed(t.Name, t.Unreachable);
            string stdout, stderr;
            int code = ProcessRunner.Run(RemoteTransport.Command(t), RemoteScripts.WithParams("probe.ps1", new JObject()), 30000, out stdout, out stderr);
            return ProbeResult.Parse(t.Name, code, stdout, stderr);
        }

        /// <summary>Runs one command in one instance and waits for the job to end (a command may sit in a picker or dialog for minutes).</summary>
        public JObject Invoke(RemoteRow row, int waitMs = 600000)
        {
            Reload();
            var target = Config.AllTargets().FirstOrDefault(t => t.Name == row.Target);
            if (target == null) return Error("no_target", "target " + row.Target + " is no longer in remote.json");
            if (target.Unreachable != null) return Error("no_target", target.Unreachable);
            var p = new JObject { ["port"] = row.Port, ["command"] = row.Command, ["wait"] = waitMs };
            string stdout, stderr;
            int code = ProcessRunner.Run(RemoteTransport.Command(target), RemoteScripts.WithParams("invoke.ps1", p), waitMs + 60000, out stdout, out stderr);
            string line = (stdout ?? "").Split('\n').Select(l => l.Trim()).LastOrDefault(l => l.StartsWith("{"));
            if (line == null) return Error("transport", code == 124 ? "timed out" : ((stderr ?? "").Trim().Split('\n').LastOrDefault() ?? "no answer"));
            try { return JObject.Parse(line); } catch (Exception) { return Error("transport", "unreadable answer"); }
        }

        private static JObject Error(string kind, string message)
        {
            return new JObject { ["error"] = new JObject { ["kind"] = kind, ["message"] = message } };
        }
    }
}
