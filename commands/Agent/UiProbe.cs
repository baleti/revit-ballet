using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace RevitBallet.Commands.Agent
{
    /// <summary>Is Revit's message pump alive, and is a modal dialog open? Both are answered without touching the Revit API (safe from any thread).</summary>
    internal static class UiProbe
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hWnd, uint cmd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int max);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
        private const uint GW_OWNER = 4, WM_NULL = 0, SMTO_ABORTIFHUNG = 2;

        private static IntPtr MainWindow() => Process.GetCurrentProcess().MainWindowHandle;

        /// <summary>True when the main window answers a message within the timeout.</summary>
        public static bool IsResponsive(int timeoutMs = 1500)
        {
            var h = MainWindow();
            if (h == IntPtr.Zero) return true;
            IntPtr r;
            return SendMessageTimeout(h, WM_NULL, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, (uint)timeoutMs, out r) != IntPtr.Zero;
        }

        /// <summary>Title of an open modal dialog (a visible window owned by the disabled main window), or null.</summary>
        public static string ModalDialogTitle()
        {
            var main = MainWindow();
            if (main == IntPtr.Zero || IsWindowEnabled(main)) return null;
            uint myPid = (uint)Process.GetCurrentProcess().Id;
            string found = null;
            EnumWindows((h, _) =>
            {
                uint pid; GetWindowThreadProcessId(h, out pid);
                if (pid != myPid || !IsWindowVisible(h) || h == main) return true;
                if (GetWindow(h, GW_OWNER) != main) return true;
                var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
                found = sb.Length > 0 ? sb.ToString() : "(untitled dialog)";
                return false;
            }, IntPtr.Zero);
            return found ?? "(modal state, dialog not found)";
        }
    }
}
