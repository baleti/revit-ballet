using System;
using System.IO;
using System.Windows.Forms;

namespace RevitBallet.Commands
{
    /// <summary>The modern (Vista-style) folder picker used by the export commands. Wraps Ookii.Dialogs (BSD-3-Clause).</summary>
    internal static class FolderPicker
    {
        private sealed class Owner : IWin32Window
        {
            public Owner(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; }
        }

        /// <returns>The chosen folder, or null when the user cancelled.</returns>
        public static string Pick(IntPtr ownerWindow, string title, string initialDirectory)
        {
            using (var dialog = new Ookii.Dialogs.WinForms.VistaFolderBrowserDialog
            {
                Description = title,
                UseDescriptionForTitle = true,
                ShowNewFolderButton = true,
            })
            {
                if (!string.IsNullOrEmpty(initialDirectory) && Directory.Exists(initialDirectory))
                    dialog.SelectedPath = initialDirectory;
                var result = dialog.ShowDialog(ownerWindow == IntPtr.Zero ? null : new Owner(ownerWindow));
                return result == DialogResult.OK ? dialog.SelectedPath : null;
            }
        }
    }
}
