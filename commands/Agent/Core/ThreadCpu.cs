using System;
using System.IO;
using System.Runtime.InteropServices;

namespace RevitBallet.Commands.Agent
{
    /// <summary>
    /// CPU time used by one thread, read from another thread. Lets the watchdog tell a compile that is working (CPU advancing) from one
    /// that is deadlocked or waiting forever (no CPU). Returns -1 where the platform gives no answer, then only the hard cap applies.
    /// </summary>
    internal sealed class ThreadCpu : IDisposable
    {
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll")] private static extern IntPtr OpenThread(uint access, bool inherit, uint threadId);
        [DllImport("kernel32.dll")] private static extern bool GetThreadTimes(IntPtr h, out long create, out long exit, out long kernel, out long user);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr h);

        private IntPtr handle; private string procStat;

        /// <summary>Call on the thread to be watched.</summary>
        public static ThreadCpu CaptureCurrent()
        {
            var t = new ThreadCpu();
            try
            {
#if NET8_0_OR_GREATER
                if (OperatingSystem.IsWindows()) t.handle = OpenThread(0x0800, false, GetCurrentThreadId());   // THREAD_QUERY_LIMITED_INFORMATION
                else if (File.Exists("/proc/thread-self/stat"))
                {
                    var tid = File.ReadAllText("/proc/thread-self/stat").Split(' ')[0];
                    t.procStat = "/proc/self/task/" + tid + "/stat";
                }
#else
                if (Environment.OSVersion.Platform == PlatformID.Win32NT) t.handle = OpenThread(0x0800, false, GetCurrentThreadId());
#endif
            }
            catch { }
            return t;
        }

        /// <summary>Milliseconds of CPU the thread has used (user + kernel), or -1 when unknown.</summary>
        public double ReadMs()
        {
            try
            {
                if (handle != IntPtr.Zero)
                {
                    long c, e, k, u;
                    if (GetThreadTimes(handle, out c, out e, out k, out u)) return (k + u) / 10000.0;   // 100 ns units
                }
                else if (procStat != null)
                {
                    var s = File.ReadAllText(procStat); var f = s.Substring(s.LastIndexOf(')') + 2).Split(' ');
                    return (long.Parse(f[11]) + long.Parse(f[12])) * 10.0;   // utime + stime, 100 ticks/s
                }
            }
            catch { }
            return -1;
        }

        public void Dispose() { if (handle != IntPtr.Zero) { CloseHandle(handle); handle = IntPtr.Zero; } }
    }
}
