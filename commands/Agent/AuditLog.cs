using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace RevitBallet.Commands.Agent
{
    /// <summary>
    /// One JSON file per finished job under %APPDATA%\revit-ballet\runtime\agent\jobs: code, output, value, error, timings, client name.
    /// The answer to "what did the agent run against my model?". Never contains headers or the token. Pruned by age and total size.
    /// </summary>
    internal static class AuditLog
    {
        private const int MaxDays = 14;
        private const long MaxBytes = 200L * 1024 * 1024;
        private static string Dir { get { return PathHelper.EnsureRuntimeSubdirectoryExists(Path.Combine("agent", "jobs")); } }

        public static void Write(Job job)
        {
            try
            {
                var pid = System.Diagnostics.Process.GetCurrentProcess().Id;
                var j = job.ToJson(full: true);
                j["pid"] = pid;
                var name = $"{job.Submitted:yyyyMMdd-HHmmss}-{pid}-{job.Id}.json";
                File.WriteAllText(Path.Combine(Dir, name), j.ToString(Newtonsoft.Json.Formatting.Indented));
            }
            catch (Exception ex) { RevitBalletServer.LogToRevit("[AGENT] audit log write failed: " + ex.Message); }
        }

        public static void Prune()
        {
            try
            {
                var files = new DirectoryInfo(Dir).GetFiles("*.json").OrderBy(f => f.LastWriteTimeUtc).ToList();
                var cutoff = DateTime.UtcNow.AddDays(-MaxDays);
                foreach (var f in files.Where(f => f.LastWriteTimeUtc < cutoff).ToList()) { try { f.Delete(); files.Remove(f); } catch { } }
                long total = files.Sum(f => f.Length);
                foreach (var f in files) { if (total <= MaxBytes) break; try { total -= f.Length; f.Delete(); } catch { } }
            }
            catch { }
        }
    }
}
