using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace RevitBallet.Core.Remote
{
    /// <summary>One machine whose Revit instances the remote picker can reach.</summary>
    public sealed class RemoteTarget
    {
        public string Name;
        /// <summary>"local" (this machine) or "ssh" (a Windows machine with OpenSSH, reached with ssh.exe).</summary>
        public string Kind = "ssh";
        public string Ssh;
        /// <summary>Why this target cannot be used (unsupported kind, bad destination); it is listed but never probed.</summary>
        public string Unreachable;
    }

    /// <summary>
    /// %APPDATA%\revit-ballet\remote.json, the same file format as the Linux tool (~/.config/revit-ballet/remote.json):
    /// { "hotkey": "Win+`", "local": true, "targets": [ { "name": "build-vm", "kind": "ssh", "ssh": "user@build-vm" } ] }
    /// </summary>
    public sealed class RemoteConfig
    {
        public string Hotkey = "Win+`";
        public bool Local = true;
        public List<RemoteTarget> Targets = new List<RemoteTarget>();

        public static RemoteConfig Parse(string json)
        {
            var cfg = new RemoteConfig();
            if (string.IsNullOrWhiteSpace(json)) return cfg;
            JObject o;
            try { o = JObject.Parse(json); }
            catch (Exception) { return cfg; }
            var hk = (string)o["hotkey"];
            if (!string.IsNullOrWhiteSpace(hk)) cfg.Hotkey = hk.Trim();
            var local = o["local"];
            if (local != null && local.Type == JTokenType.Boolean) cfg.Local = (bool)local;
            var targets = o["targets"] as JArray;
            if (targets != null)
            {
                foreach (var t in targets.OfType<JObject>())
                {
                    var target = new RemoteTarget { Name = (string)t["name"], Kind = ((string)t["kind"] ?? "ssh").ToLowerInvariant(), Ssh = (string)t["ssh"] };
                    if (string.IsNullOrWhiteSpace(target.Name)) continue;
                    if (target.Kind == "ssh")
                    {
                        if (string.IsNullOrWhiteSpace(target.Ssh)) target.Unreachable = "no \"ssh\" destination given";
                        else if (target.Ssh.StartsWith("-")) target.Unreachable = "the ssh destination must not start with '-'";
                        else if (target.Ssh.Any(char.IsWhiteSpace)) target.Unreachable = "the ssh destination must not contain spaces";
                    }
                    else if (target.Kind == "local") { }
                    else target.Unreachable = "kind \"" + target.Kind + "\" is not supported on Windows (ssh and local are)";
                    cfg.Targets.Add(target);
                }
            }
            return cfg;
        }

        /// <summary>The targets to probe: this machine first (unless "local": false), then the configured ones.</summary>
        public List<RemoteTarget> AllTargets()
        {
            var all = new List<RemoteTarget>();
            if (Local) all.Add(new RemoteTarget { Name = "local", Kind = "local" });
            all.AddRange(Targets.Where(t => !(Local && t.Kind == "local")));
            return all;
        }
    }

    /// <summary>The program and arguments that run a PowerShell script read from stdin on a target.</summary>
    public sealed class ProcessSpec
    {
        public string FileName;
        public string Arguments;
    }

    public static class RemoteTransport
    {
        private const string PowerShell = "powershell -NoProfile -ExecutionPolicy Bypass -Command -";

        public static ProcessSpec Command(RemoteTarget t)
        {
            if (t.Unreachable != null) throw new InvalidOperationException(t.Name + ": " + t.Unreachable);
            if (t.Kind == "local")
                return new ProcessSpec { FileName = "powershell.exe", Arguments = "-NoProfile -ExecutionPolicy Bypass -Command -" };
            if (t.Kind == "ssh")
                return new ProcessSpec { FileName = "ssh.exe", Arguments = "-o BatchMode=yes -o ConnectTimeout=8 -o ServerAliveInterval=5 -o ServerAliveCountMax=3 " + t.Ssh + " " + PowerShell };
            throw new InvalidOperationException("unknown target kind " + t.Kind);
        }
    }
}
