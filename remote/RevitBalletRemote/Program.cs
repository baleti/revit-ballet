using System;
using System.Threading;
using System.Windows.Forms;

namespace RevitBalletRemote
{
    internal static class Program
    {
        /// <summary>
        /// RevitBalletRemote.exe [--show]
        /// Runs in the tray, owns the global hotkey (default Win+`, change "hotkey" in %APPDATA%\revit-ballet\remote.json) and shows the picker.
        /// A second start does not run a second copy: with --show it asks the first one to open the picker.
        /// </summary>
        [STAThread]
        private static void Main(string[] args)
        {
            bool show = Array.IndexOf(args, "--show") >= 0;
            bool first;
            using (var mutex = new Mutex(true, @"Local\RevitBalletRemote", out first))
            using (var showEvent = new EventWaitHandle(false, EventResetMode.AutoReset, @"Local\RevitBalletRemote.show"))
            {
                if (!first)
                {
                    if (show) showEvent.Set();
                    return;
                }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                var app = new TrayApp(show);
                var watcher = new Thread(() => { while (true) { showEvent.WaitOne(); app.Show(); } }) { IsBackground = true, Name = "show-requests" };
                watcher.Start();
                Application.Run(app);
            }
        }
    }
}
