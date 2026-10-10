using System;
using System.IO;
using System.Reflection;
using System.Text;
using Newtonsoft.Json.Linq;

namespace RevitBallet.Core.Remote
{
    /// <summary>The PowerShell that runs on a target (embedded copies of tools/remote/probe.ps1 and invoke.ps1).</summary>
    public static class RemoteScripts
    {
        public static string Load(string name)
        {
            var asm = typeof(RemoteScripts).Assembly;
            using (var s = asm.GetManifestResourceStream("remote/" + name))
            {
                if (s == null) throw new InvalidOperationException("embedded script not found: " + name);
                using (var r = new StreamReader(s, Encoding.UTF8)) return r.ReadToEnd();
            }
        }

        /// <summary>`$P = '&lt;json&gt;' | ConvertFrom-Json` on one line (no here-strings, nothing to quote), then the script.</summary>
        public static string WithParams(string scriptName, JObject parameters)
        {
            string literal = parameters.ToString(Newtonsoft.Json.Formatting.None).Replace("'", "''");
            return "$P = '" + literal + "' | ConvertFrom-Json\n" + Load(scriptName) + "\n";
        }
    }
}
