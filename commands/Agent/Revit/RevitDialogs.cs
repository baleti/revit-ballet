using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace RevitBallet.Commands.Agent
{
    /// <summary>
    /// What is the modal dialog that holds Revit up, and how do we press a button of it. Detection is Win32 (a visible window owned by the disabled main
    /// window), text and buttons come from UI Automation, and Revit's own DialogBoxShowing event adds the TaskDialog id and message. Everything here
    /// is callable from any thread while the UI thread sits in the dialog's message loop.
    /// </summary>
    internal static class RevitDialogs
    {
        private sealed class Shown { public string Id; public string Message; public string Kind; public DateTime At; }
        private static volatile Shown lastShown;
        private static UIControlledApplication app;

        /// <summary>Call once from a Revit API context (OnStartup).</summary>
        public static void Attach(UIControlledApplication application)
        {
            if (app != null) return;
            app = application;
            application.DialogBoxShowing += OnShowing;
        }

        public static void Detach()
        {
            if (app == null) return;
            try { app.DialogBoxShowing -= OnShowing; } catch { }
            app = null;
        }

        private static void OnShowing(object sender, DialogBoxShowingEventArgs e)
        {
            try
            {
                var task = e as TaskDialogShowingEventArgs;
                if (task != null) lastShown = new Shown { Id = task.DialogId, Message = task.Message, Kind = "task", At = DateTime.UtcNow };
                else
                {
                    var box = e as MessageBoxShowingEventArgs;
                    lastShown = new Shown { Id = "MessageBox:" + (box != null ? box.DialogType.ToString() : "?"), Message = box != null ? box.Message : null, Kind = "win32", At = DateTime.UtcNow };
                }
            }
            catch { }
        }

        /// <summary>The open dialog, or null. Gives up on UI Automation after 2 s and reports what Win32 knows (the title).</summary>
        public static DialogInfo Current()
        {
            string title;
            var hwnd = UiProbe.FindModalWindow(out title);
            if (title == null) return null;
            var d = new DialogInfo { Title = title, Kind = "win32" };
            var shown = lastShown;
            if (shown != null && (DateTime.UtcNow - shown.At).TotalSeconds < 60) { d.Kind = shown.Kind; d.DialogId = shown.Id; d.Text = shown.Message; }
            if (hwnd == IntPtr.Zero) return d;
            try
            {
                var read = Task.Run(() => ReadWithAutomation(hwnd));
                if (read.Wait(2000))
                {
                    var r = read.Result;
                    if (!string.IsNullOrWhiteSpace(r.Item1)) d.Text = r.Item1;
                    d.Buttons.AddRange(r.Item2);
                }
            }
            catch { }
            if (d.Buttons.Count == 0) d.Buttons.AddRange(UiProbe.Win32Buttons(hwnd).Select(b => b.Key));
            return d;
        }

        private static Tuple<string, List<string>> ReadWithAutomation(IntPtr hwnd)
        {
            var root = AutomationElement.FromHandle(hwnd);
            var texts = new List<string>();
            var buttons = new List<string>();
            foreach (AutomationElement t in root.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)))
            {
                var n = t.Current.Name; if (!string.IsNullOrWhiteSpace(n) && !texts.Contains(n)) texts.Add(n.Trim());
            }
            foreach (AutomationElement b in root.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)))
            {
                var parent = TreeWalker.ControlViewWalker.GetParent(b);
                if (parent != null && parent.Current.ControlType == ControlType.TitleBar) continue;   // the caption's own minimise/close buttons
                var n = (b.Current.Name ?? "").Replace("&", "").Trim();
                if (n.Length > 0 && !buttons.Contains(n)) buttons.Add(n);
            }
            return Tuple.Create(string.Join("\n", texts), buttons);
        }

        [DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
        [DllImport("user32.dll")] private static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);

        /// <summary>Press the named button: UI Automation Invoke, then a Win32 BM_CLICK, then a real mouse click; each checked by seeing the dialog go.</summary>
        public static bool TryRespond(string button, out string error)
        {
            error = null;
            string title;
            var hwnd = UiProbe.FindModalWindow(out title);
            if (title == null) { error = "no modal dialog is open"; return false; }
            if (hwnd == IntPtr.Zero) { error = "the dialog window could not be found"; return false; }
            var wanted = (button ?? "").Replace("&", "").Trim();
            try
            {
                AutomationElement target = null;
                var found = Task.Run(() =>
                {
                    var root = AutomationElement.FromHandle(hwnd);
                    foreach (AutomationElement b in root.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)))
                    {
                        var parent = TreeWalker.ControlViewWalker.GetParent(b);
                        if (parent != null && parent.Current.ControlType == ControlType.TitleBar) continue;
                        if (string.Equals((b.Current.Name ?? "").Replace("&", "").Trim(), wanted, StringComparison.OrdinalIgnoreCase)) return b;
                    }
                    return null;
                });
                if (found.Wait(2000)) target = found.Result;

                if (target != null)
                {
                    try { object p; if (target.TryGetCurrentPattern(InvokePattern.Pattern, out p)) Task.Run(() => ((InvokePattern)p).Invoke()); } catch { }
                    if (GoneWithin(hwnd, 1500)) return true;
                }
                var w32 = UiProbe.Win32Buttons(hwnd).FirstOrDefault(b => string.Equals(b.Key, wanted, StringComparison.OrdinalIgnoreCase));
                if (w32.Value != IntPtr.Zero)
                {
                    UiProbe.ClickWin32Button(w32.Value);
                    if (GoneWithin(hwnd, 1500)) return true;
                }
                if (target != null)
                {
                    var r = target.Current.BoundingRectangle;
                    if (!r.IsEmpty)
                    {
                        SetCursorPos((int)(r.Left + r.Width / 2), (int)(r.Top + r.Height / 2));
                        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero); mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);   // left down, up
                        if (GoneWithin(hwnd, 1500)) return true;
                    }
                }
                var have = Current();
                error = (target == null && w32.Value == IntPtr.Zero)
                    ? $"the dialog '{title}' has no button '{button}' (buttons: {(have == null || have.Buttons.Count == 0 ? "none found" : string.Join(", ", have.Buttons))})"
                    : $"pressed '{button}' but the dialog '{title}' is still open";
                return false;
            }
            catch (Exception ex) { error = ex.GetType().Name + ": " + ex.Message; return false; }
        }

        private static bool GoneWithin(IntPtr hwnd, int ms)
        {
            var end = DateTime.UtcNow.AddMilliseconds(ms);
            while (DateTime.UtcNow < end)
            {
                string t; var now = UiProbe.FindModalWindow(out t);
                if (t == null || now != hwnd) return true;
                Thread.Sleep(50);
            }
            return false;
        }
    }
}
