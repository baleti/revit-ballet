using System;
using System.Collections.Generic;
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

        /// <summary>The open modal dialog window (a visible window owned by the disabled main window), or IntPtr.Zero.</summary>
        public static IntPtr FindModalWindow(out string title)
        {
            title = null;
            var main = MainWindow();
            if (main == IntPtr.Zero || IsWindowEnabled(main)) return IntPtr.Zero;
            uint myPid = (uint)Process.GetCurrentProcess().Id;
            IntPtr found = IntPtr.Zero; string foundTitle = null;
            EnumWindows((h, _) =>
            {
                uint pid; GetWindowThreadProcessId(h, out pid);
                if (pid != myPid || !IsWindowVisible(h) || h == main) return true;
                if (GetWindow(h, GW_OWNER) != main) return true;
                var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
                found = h; foundTitle = sb.Length > 0 ? sb.ToString() : "(untitled dialog)";
                return false;
            }, IntPtr.Zero);
            title = found == IntPtr.Zero ? "(modal state, dialog not found)" : foundTitle;
            return found;
        }

        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int max);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr w, IntPtr l);
        [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
        private const uint BM_CLICK = 0x00F5;

        /// <summary>Classic Win32 buttons of a dialog (caption, handle): MessageBox, standard dialogs. Not the DirectUI buttons of a TaskDialog.</summary>
        public static System.Collections.Generic.List<KeyValuePair<string, IntPtr>> Win32Buttons(IntPtr dialog)
        {
            var list = new System.Collections.Generic.List<KeyValuePair<string, IntPtr>>();
            EnumChildWindows(dialog, (h, _) =>
            {
                var cls = new StringBuilder(64); GetClassName(h, cls, 64);
                if (cls.ToString().Equals("Button", StringComparison.OrdinalIgnoreCase) && IsWindowVisible(h))
                {
                    var t = new StringBuilder(128); GetWindowText(h, t, 128);
                    if (t.Length > 0) list.Add(new KeyValuePair<string, IntPtr>(t.ToString().Replace("&", ""), h));
                }
                return true;
            }, IntPtr.Zero);
            return list;
        }

        public static bool ClickWin32Button(IntPtr button) { SendMessage(button, BM_CLICK, IntPtr.Zero, IntPtr.Zero); return true; }
    }
}
