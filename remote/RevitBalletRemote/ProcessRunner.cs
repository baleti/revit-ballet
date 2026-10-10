using System;
using System.Diagnostics;
using System.Text;
using System.Threading;
using RevitBallet.Core.Remote;

namespace RevitBalletRemote
{
    internal static class ProcessRunner
    {
        /// <summary>Runs a program feeding it `stdin` and returns (exit code, stdout, stderr). Exit code 124 = killed after the timeout.</summary>
        public static int Run(ProcessSpec spec, string stdin, int timeoutMs, out string stdout, out string stderr)
        {
            var psi = new ProcessStartInfo(spec.FileName, spec.Arguments)
            {
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                UseShellExecute = false, CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8,
            };
            var outSb = new StringBuilder(); var errSb = new StringBuilder();
            try
            {
                using (var p = new Process { StartInfo = psi })
                {
                    p.OutputDataReceived += (s, e) => { if (e.Data != null) lock (outSb) outSb.AppendLine(e.Data); };
                    p.ErrorDataReceived += (s, e) => { if (e.Data != null) lock (errSb) errSb.AppendLine(e.Data); };
                    p.Start();
                    p.BeginOutputReadLine(); p.BeginErrorReadLine();
                    // powershell reads a script from stdin as UTF-8 without a BOM when we write the bytes ourselves
                    var bytes = new UTF8Encoding(false).GetBytes(stdin ?? "");
                    try { p.StandardInput.BaseStream.Write(bytes, 0, bytes.Length); p.StandardInput.BaseStream.Flush(); p.StandardInput.Close(); } catch (Exception) { /* the process may already have exited */ }
                    if (!p.WaitForExit(timeoutMs))
                    {
                        try { p.Kill(); } catch (Exception) { }
                        p.WaitForExit(2000);
                        lock (outSb) stdout = outSb.ToString(); lock (errSb) stderr = "timed out after " + timeoutMs / 1000 + "s";
                        return 124;
                    }
                    p.WaitForExit();            // flush the async readers
                    lock (outSb) stdout = outSb.ToString(); lock (errSb) stderr = errSb.ToString();
                    return p.ExitCode;
                }
            }
            catch (Exception ex)
            {
                stdout = ""; stderr = ex.Message;
                return 127;                      // e.g. ssh.exe not installed
            }
        }
    }
}
