using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using RevitBallet.Core.Remote;

namespace RevitBalletRemote
{
    /// <summary>Tray icon + global hotkey + the picker. Runs for the whole Windows session (started by the installer's opt-in Run entry).</summary>
    internal sealed class TrayApp : ApplicationContext
    {
        private readonly NotifyIcon icon = new NotifyIcon();
        private readonly HotkeyWindow hotkey = new HotkeyWindow();
        private readonly RemoteService service = new RemoteService();
        private readonly PickerForm picker;
        private readonly Control ui = new Control();            // lives on the UI thread, so background tasks can BeginInvoke back

        public TrayApp(bool showAtStart)
        {
            ui.CreateControl(); var _ = ui.Handle;
            picker = new PickerForm(service);
            picker.Invoked += OnInvoked;

            icon.Icon = LoadIcon();
            icon.Text = "Revit Ballet - remote commands";
            var menu = new ContextMenuStrip();
            menu.Items.Add("Open command picker", null, (s, e) => TogglePicker(true));
            menu.Items.Add("Edit remote.json", null, (s, e) => EditConfig());
            menu.Items.Add("Quit", null, (s, e) => Exit());
            icon.ContextMenuStrip = menu;
            icon.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) TogglePicker(true); };
            icon.Visible = true;

            hotkey.Pressed += () => TogglePicker(false);
            HotkeySpec spec;
            if (!HotkeySpec.TryParse(service.Config.Hotkey, out spec)) { Notify("Remote commands", "\"" + service.Config.Hotkey + "\" in remote.json is not a hotkey; using Win+`.", ToolTipIcon.Warning); HotkeySpec.TryParse("Win+`", out spec); }
            if (!hotkey.Register(spec))
                Notify("Remote commands", "Could not register " + spec + ": " + hotkey.Error + ". The tray icon still opens the picker; change \"hotkey\" in remote.json.", ToolTipIcon.Warning);

            if (showAtStart) TogglePicker(true);
        }

        /// <summary>Another start of the program asked this one to open the picker.</summary>
        public void Show() { ui.BeginInvoke((Action)(() => TogglePicker(true))); }

        private void TogglePicker(bool forceOpen)
        {
            if (!forceOpen && picker.Visible) { picker.Hide(); return; }
            picker.Open();
        }

        private void OnInvoked(RemoteRow row)
        {
            Notify(row.Preview + " (" + row.Fields["instance"] + ")", "running...", ToolTipIcon.Info);
            Task.Run(() => service.Invoke(row)).ContinueWith(t =>
            {
                string word, detail;
                if (t.Exception != null) { word = "failed"; detail = t.Exception.GetBaseException().Message; }
                else RemoteOutcome.Describe(t.Result, out word, out detail);
                ui.BeginInvoke((Action)(() => Notify(row.Preview + " (" + row.Fields["instance"] + "): " + word, detail.Length > 250 ? detail.Substring(0, 247) + "..." : detail,
                    word == "done" || word == "cancelled" ? ToolTipIcon.Info : ToolTipIcon.Error)));
            });
        }

        private void Notify(string title, string text, ToolTipIcon kind)
        {
            icon.ShowBalloonTip(8000, title, string.IsNullOrEmpty(text) ? " " : text, kind);
        }

        private void EditConfig()
        {
            try
            {
                string path = RemoteService.ConfigPath;
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                if (!File.Exists(path)) File.WriteAllText(path, "{\n  \"hotkey\": \"Win+`\",\n  \"local\": true,\n  \"targets\": [\n  ]\n}\n");
                Process.Start(new ProcessStartInfo("notepad.exe", "\"" + path + "\"") { UseShellExecute = true });
            }
            catch (Exception ex) { Notify("Remote commands", ex.Message, ToolTipIcon.Error); }
        }

        private static Icon LoadIcon()
        {
            try { var exe = System.Reflection.Assembly.GetExecutingAssembly().Location; var i = Icon.ExtractAssociatedIcon(exe); if (i != null) return i; } catch (Exception) { }
            return SystemIcons.Application;
        }

        private void Exit()
        {
            icon.Visible = false; icon.Dispose(); hotkey.Dispose();
            ExitThread();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { icon.Dispose(); hotkey.Dispose(); picker.Dispose(); ui.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
