using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

namespace RevitBalletInstaller
{
    /// <summary>--remote-picker: Ask = a dialog (interactive default); On/Off = decide without asking; Keep = leave it as it is (quiet default).</summary>
    public enum RemotePickerMode { Ask, On, Off, Keep }

    /// <summary>
    /// The optional remote command picker (remote/RevitBalletRemote): a small tray program with a global hotkey (Win+`) that lists Revit Ballet's commands
    /// from outside Revit and runs one in any Revit instance on this machine or on machines listed in remote.json. Installed per user (no admin),
    /// started with Windows through the user's Run key. Nothing is installed unless the user opts in.
    /// </summary>
    internal static class RemotePickerComponent
    {
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunValue = "RevitBalletRemote";
        private static readonly string[] Files = { "RevitBalletRemote.exe", "RevitBallet.Core.dll", "Newtonsoft.Json.dll" };

        public static string InstallDir
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "revit-ballet", "remote-picker"); }
        }

        public static string ExePath { get { return Path.Combine(InstallDir, "RevitBalletRemote.exe"); } }

        public static bool IsInstalled()
        {
            using (var k = Registry.CurrentUser.OpenSubKey(RunKeyPath))
                return k != null && k.GetValue(RunValue) != null && File.Exists(ExePath);
        }

        /// <summary>True when this installer carries the component (a build without the remote project has nothing to offer).</summary>
        public static bool Available(Assembly asm)
        {
            return asm.GetManifestResourceNames().Contains("remote/RevitBalletRemote.exe");
        }

        public static void StopRunning()
        {
            foreach (var p in Process.GetProcessesByName("RevitBalletRemote"))
            {
                try { p.Kill(); p.WaitForExit(3000); } catch { }
            }
        }

        public static void Install(Assembly asm, bool startNow, Action<string> log)
        {
            StopRunning();
            Directory.CreateDirectory(InstallDir);
            foreach (var f in Files)
            {
                using (var s = asm.GetManifestResourceStream("remote/" + f))
                {
                    if (s == null) throw new InvalidOperationException("this installer does not contain " + f);
                    using (var o = new FileStream(Path.Combine(InstallDir, f), FileMode.Create, FileAccess.Write)) s.CopyTo(o);
                }
            }
            using (var k = Registry.CurrentUser.CreateSubKey(RunKeyPath)) k.SetValue(RunValue, "\"" + ExePath + "\"");
            log("remote picker installed in " + InstallDir + " (starts with Windows; hotkey Win+`, change it in %APPDATA%\\revit-ballet\\remote.json)");
            if (startNow)
            {
                try { Process.Start(new ProcessStartInfo(ExePath) { UseShellExecute = false, WorkingDirectory = InstallDir }); log("remote picker started"); }
                catch (Exception ex) { log("could not start the remote picker: " + ex.Message); }
            }
            else log("remote picker will start at the next sign-in (or run " + ExePath + ")");
        }

        public static void Remove(Action<string> log)
        {
            bool had = IsInstalled() || Directory.Exists(InstallDir);
            StopRunning();
            try { using (var k = Registry.CurrentUser.OpenSubKey(RunKeyPath, true)) { if (k != null) k.DeleteValue(RunValue, false); } } catch { }
            try { if (Directory.Exists(InstallDir)) Directory.Delete(InstallDir, true); } catch { /* a locked file goes with the next run */ }
            if (had) log("remote picker removed");
        }

        /// <summary>The opt-in dialog. Returns true to install/keep it, false to remove/leave it out. Default: unchecked unless it is already installed.</summary>
        public static bool Ask(bool installedNow)
        {
            using (var form = new Form
            {
                Text = "Revit Ballet - optional component",
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                ClientSize = new Size(560, 270),
            })
            {
                var title = new Label { Text = "Remote command picker", Font = new Font(SystemFonts.DialogFont.FontFamily, 12f, FontStyle.Bold), Left = 16, Top = 14, Width = 520, Height = 26 };
                var body = new Label
                {
                    Left = 16, Top = 46, Width = 528, Height = 130,
                    Text = "A small background program (tray icon) that lists Revit Ballet's commands from outside Revit and runs one in any Revit window: " +
                           "on this computer, or on other computers you list in remote.json (reached with ssh).\r\n\r\n" +
                           "Press Win + ` (backtick) to open it; the same query language as the grids in Revit. It starts with Windows for your account only, " +
                           "opens no network port, and is removed by running this installer again with the box unchecked, or by uninstalling Revit Ballet.",
                };
                var box = new CheckBox { Text = "Install the remote command picker (Win + `)", Left = 16, Top = 190, Width = 520, Checked = installedNow };
                var ok = new Button { Text = "Continue", Left = 440, Top = 228, Width = 104, Height = 30, DialogResult = DialogResult.OK };
                form.Controls.AddRange(new Control[] { title, body, box, ok });
                form.AcceptButton = ok;
                form.ShowDialog();
                return box.Checked;
            }
        }
    }

    public partial class BalletInstaller
    {
        /// <summary>Set from --remote-picker.</summary>
        public static RemotePickerMode RemotePicker = RemotePickerMode.Ask;

        private bool installRemotePicker;
        private bool applyRemotePicker;

        /// <summary>Decides (asking in the interactive installer) whether the remote command picker should be installed, kept or removed.</summary>
        private bool PrepareRemotePickerChoice(bool quiet)
        {
            var asm = Assembly.GetExecutingAssembly();
            bool installedNow = RemotePickerComponent.IsInstalled();
            var mode = RemotePicker;
            if (mode == RemotePickerMode.Ask && quiet) mode = RemotePickerMode.Keep;
            switch (mode)
            {
                case RemotePickerMode.On: installRemotePicker = true; applyRemotePicker = true; break;
                case RemotePickerMode.Off: installRemotePicker = false; applyRemotePicker = installedNow; break;
                case RemotePickerMode.Keep: installRemotePicker = installedNow; applyRemotePicker = installedNow; break;      // refresh the files of an installed one
                default:
                    if (!RemotePickerComponent.Available(asm)) { applyRemotePicker = false; break; }
                    installRemotePicker = RemotePickerComponent.Ask(installedNow);
                    applyRemotePicker = installRemotePicker || installedNow;
                    break;
            }
            if (installRemotePicker && !RemotePickerComponent.Available(asm))
            {
                if (quiet) Console.WriteLine("this installer was built without the remote command picker; skipping it");
                installRemotePicker = false; applyRemotePicker = false;
            }
            return true;
        }

        private void ApplyRemotePicker(bool quiet)
        {
            if (!applyRemotePicker) return;
            Action<string> log = m => Log(1, m);
            try
            {
                if (installRemotePicker) RemotePickerComponent.Install(Assembly.GetExecutingAssembly(), !quiet, log);
                else RemotePickerComponent.Remove(log);
            }
            catch (Exception ex)
            {
                string msg = "remote command picker: " + ex.Message;
                if (quiet) Console.Error.WriteLine(msg); else MessageBox.Show(msg, "Revit Ballet", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
