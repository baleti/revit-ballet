using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OpenMcdf;
using RevitBallet.Commands;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;

// Opens one or more .rvt files, each in a new Revit process, using settings stored in
// %appdata%/revit-ballet/OpenRvtDocumentInNewSession.toml.
// Per CLAUDE.md policy: not registered in addin manifest or keyboard shortcuts.
[Transaction(TransactionMode.Manual)]
[CommandMeta("")]
public class OpenRvtDocumentInNewSession : IExternalCommand
{
    private static readonly string SettingsPath = Path.Combine(
        PathHelper.RevitBalletDirectory, "OpenRvtDocumentInNewSession.toml");

    private static readonly string PendingOpenDir = Path.Combine(
        PathHelper.RevitBalletDirectory, "pending-open");

    // ── Entry point ──────────────────────────────────────────────────────────

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uiApp = commandData.Application;
        var settings = TomlHelper.Read(SettingsPath);

        string lastFolder        = TomlHelper.GetString(settings, "LastFolder", "");
        bool   createNewLocal    = TomlHelper.GetBool(settings, "CreateNewLocal", true);
        string worksetOption     = TomlHelper.GetString(settings, "WorksetOption", "LastViewed");
        bool   useFileVersion    = TomlHelper.GetBool(settings, "UseFileRevitVersion", true);
        int    delaySeconds      = TomlHelper.GetInt(settings, "DelayBetweenOpens", 20);

        // ── File picker ──────────────────────────────────────────────────────

        var dlg = new OpenFileDialog
        {
            Title       = "Select Revit Files to Open in New Sessions",
            Filter      = "Revit Project Files (*.rvt)|*.rvt",
            Multiselect = true,
            InitialDirectory = Directory.Exists(lastFolder) ? lastFolder : ""
        };

        var owner = new RevitWindow(Helpers.GetMainWindowHandle(uiApp));
        if (dlg.ShowDialog(owner) != DialogResult.OK || dlg.FileNames.Length == 0)
            return Result.Cancelled;

        string[] selectedFiles = dlg.FileNames;

        // Persist last folder immediately.
        settings["LastFolder"] = Path.GetDirectoryName(selectedFiles[0]) ?? "";
        TomlHelper.Write(SettingsPath, settings);

        // ── Read metadata from each .rvt without opening Revit ───────────────

        var fileInfos = selectedFiles
            .Select(f => ReadRvtFileInfo(f))
            .ToList();

        // ── Workset selection DataGrid (only when "Specify" is active) ───────

        // worksetsByFile: filePath → list of workset name strings selected to open
        var worksetsByFile = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        if (worksetOption == "Specify")
        {
            var worksharedInfos = fileInfos.Where(fi => fi.IsWorkshared).ToList();
            if (worksharedInfos.Count > 0)
            {
                bool ok = GatherAndSelectWorksets(uiApp, worksharedInfos, worksetsByFile);
                if (!ok) return Result.Cancelled;
            }
        }

        // ── Launch one Revit process per file ────────────────────────────────

        Directory.CreateDirectory(PendingOpenDir);

        // Run on a background thread so the staggered delay doesn't freeze the UI.
        var captures = new { fileInfos, worksetsByFile, createNewLocal, worksetOption,
                             useFileVersion, delaySeconds };

        System.Threading.Tasks.Task.Run(() =>
        {
            for (int i = 0; i < captures.fileInfos.Count; i++)
            {
                if (i > 0)
                    Thread.Sleep(captures.delaySeconds * 1000);

                LaunchForFile(captures.fileInfos[i], captures.worksetsByFile,
                              captures.createNewLocal, captures.worksetOption,
                              captures.useFileVersion);
            }
        });

        return Result.Succeeded;
    }

    // ── Metadata reading (no Revit API / no Revit process) ──────────────────

    private sealed class RvtFileInfo
    {
        public string FilePath     { get; set; }
        public bool   IsWorkshared { get; set; }
        public bool   IsCentral    { get; set; }
        public string RevitYear    { get; set; } // e.g. "2024", or null if unknown
    }

    private static RvtFileInfo ReadRvtFileInfo(string filePath)
    {
        var info = new RvtFileInfo { FilePath = filePath };
        try
        {
            // BasicFileInfo.Extract reads the OLE stream without opening a full document.
            var bfi = BasicFileInfo.Extract(filePath);
            info.IsWorkshared = bfi.IsWorkshared;
            info.IsCentral    = bfi.IsWorkshared && bfi.IsCentral;
        }
        catch { /* non-workshared or unreadable */ }

        info.RevitYear = ReadRevitYearFromStream(filePath);
        return info;
    }

    // Reads the BasicFileInfo binary stream via OpenMcdf and extracts the year from
    // "Autodesk Revit YYYY" text embedded in the stream.
    private static string ReadRevitYearFromStream(string filePath)
    {
        try
        {
            CompoundFile cf = null;
            try
            {
                cf = new CompoundFile(filePath);
                var stream = cf.RootStorage.GetStream("BasicFileInfo");
                byte[] bytes = stream.GetData();

                // The stream header is binary; the metadata text is UTF-16 LE after it.
                // Try both decodings — Unicode first (most reliable).
                foreach (var enc in new[] { System.Text.Encoding.Unicode, System.Text.Encoding.ASCII })
                {
                    string text = enc.GetString(bytes);
                    var m = Regex.Match(text, @"Autodesk Revit\s+(\d{4})");
                    if (m.Success) return m.Groups[1].Value;
                }
            }
            finally
            {
                try { cf?.Close(); } catch { }
            }
        }
        catch { }
        return null;
    }

    // ── Workset gathering (opens files temporarily in current session) ────────

    // Returns false if the user cancelled the DataGrid.
    private static bool GatherAndSelectWorksets(
        UIApplication uiApp,
        List<RvtFileInfo> worksharedInfos,
        Dictionary<string, List<string>> worksetsByFile)
    {
        // worksetName → set of file paths that have it
        var worksetToFiles = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        bool multiFile = worksharedInfos.Count > 1;

        foreach (var fi in worksharedInfos)
        {
            List<string> names = OpenAndReadWorksetNames(uiApp, fi.FilePath);
            foreach (var name in names)
            {
                if (!worksetToFiles.TryGetValue(name, out var set))
                    worksetToFiles[name] = set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                set.Add(fi.FilePath);
            }
        }

        if (worksetToFiles.Count == 0)
        {
            TaskDialog.Show("Worksets", "No user worksets found in the selected workshared files.");
            return true; // Not a cancellation — just nothing to specify.
        }

        var rows = worksetToFiles
            .OrderBy(kv => kv.Key)
            .Select(kv =>
            {
                var row = new Dictionary<string, object>
                {
                    { "Workset", kv.Key },
                    { "In Files", kv.Value.Count }
                };
                if (multiFile)
                    row["Files"] = string.Join(", ",
                        kv.Value.Select(p => Path.GetFileName(p)).OrderBy(n => n));
                return row;
            })
            .ToList();

        var columns = multiFile
            ? new List<string> { "Workset", "In Files", "Files" }
            : new List<string> { "Workset" };

        var picked = CustomGUIs.DataGrid(rows, columns, false);
        if (picked == null || picked.Count == 0) return false;

        var selectedNames = new HashSet<string>(
            picked
                .Select(r => r.TryGetValue("Workset", out var v) ? v as string : null)
                .Where(n => n != null),
            StringComparer.OrdinalIgnoreCase);

        // Map selected workset names back to each file.
        foreach (var fi in worksharedInfos)
        {
            var forFile = worksetToFiles
                .Where(kv => selectedNames.Contains(kv.Key) && kv.Value.Contains(fi.FilePath))
                .Select(kv => kv.Key)
                .ToList();
            worksetsByFile[fi.FilePath] = forFile;
        }

        return true;
    }

    // Opens a file detached with all worksets closed (fastest way to get workset names),
    // reads the user workset names, closes the document, returns the names.
    private static List<string> OpenAndReadWorksetNames(UIApplication uiApp, string filePath)
    {
        Document tempDoc = null;
        try
        {
            var opts = new OpenOptions();
            opts.DetachFromCentralOption =
                DetachFromCentralOption.DetachAndPreserveWorksets;
            opts.SetOpenWorksetsConfiguration(
                new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));

            var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            tempDoc = uiApp.Application.OpenDocumentFile(modelPath, opts);

            return new FilteredWorksetCollector(tempDoc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .Select(ws => ws.Name)
                .OrderBy(n => n)
                .ToList();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[OpenRvtDocumentInNewSession] Cannot read worksets from {filePath}: {ex.Message}");
            return new List<string>();
        }
        finally
        {
            try { tempDoc?.Close(false); } catch { }
        }
    }

    // ── Launching ────────────────────────────────────────────────────────────

    private static void LaunchForFile(
        RvtFileInfo fi,
        Dictionary<string, List<string>> worksetsByFile,
        bool createNewLocal,
        string worksetOption,
        bool useFileVersion)
    {
        string revitExe = FindRevitExe(fi.RevitYear, useFileVersion);
        if (string.IsNullOrEmpty(revitExe) || !File.Exists(revitExe))
        {
            Debug.WriteLine($"[OpenRvtDocumentInNewSession] Revit.exe not found for year {fi.RevitYear}");
            return;
        }

        Process proc = null;
        try
        {
            var psi = new ProcessStartInfo(revitExe) { UseShellExecute = true };
            proc = Process.Start(psi);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[OpenRvtDocumentInNewSession] Failed to start Revit: {ex.Message}");
            return;
        }

        if (proc == null) return;

        // Build and write the pending-open config keyed by the new process's PID.
        worksetsByFile.TryGetValue(fi.FilePath, out var wsNames);
        var cfg = new Dictionary<string, string>
        {
            ["FilePath"]      = fi.FilePath,
            ["IsWorkshared"]  = fi.IsWorkshared ? "true" : "false",
            ["IsCentral"]     = fi.IsCentral    ? "true" : "false",
            ["CreateNewLocal"] = createNewLocal  ? "true" : "false",
            ["WorksetOption"] = worksetOption,
            ["WorksetsToOpen"] = wsNames != null && wsNames.Count > 0
                                 ? string.Join("|", wsNames)
                                 : ""
        };

        string configPath = Path.Combine(PendingOpenDir, $"{proc.Id}.toml");
        try { TomlHelper.Write(configPath, cfg); }
        catch (Exception ex)
        {
            Debug.WriteLine($"[OpenRvtDocumentInNewSession] Failed to write pending config: {ex.Message}");
        }
    }

    // ── Revit.exe discovery ──────────────────────────────────────────────────

    private static string FindRevitExe(string yearString, bool useFileVersion)
    {
        // If the user doesn't want version matching, use the currently running Revit.
        if (!useFileVersion || string.IsNullOrEmpty(yearString))
            return Process.GetCurrentProcess().MainModule?.FileName;

        // 1. Registry lookup (64-bit hive, then WOW64 fallback).
        string[] regBases = {
            $@"SOFTWARE\Autodesk\Revit\Autodesk Revit {yearString}",
            $@"SOFTWARE\WOW6432Node\Autodesk\Revit\Autodesk Revit {yearString}"
        };
        foreach (var regBase in regBases)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(regBase))
                {
                    if (key == null) continue;
                    // Enumerate sub-keys (e.g. "REVIT-x:\") looking for InstallationPath.
                    foreach (var subName in key.GetSubKeyNames())
                    {
                        using (var sub = key.OpenSubKey(subName))
                        {
                            if (sub == null) continue;
                            string installPath = sub.GetValue("InstallationPath") as string;
                            if (!string.IsNullOrEmpty(installPath))
                            {
                                string exe = Path.Combine(installPath, "Revit.exe");
                                if (File.Exists(exe)) return exe;
                            }
                        }
                    }
                }
            }
            catch { }
        }

        // 2. Common install paths.
        string[] commonPaths = {
            $@"C:\Program Files\Autodesk\Revit {yearString}\Revit.exe",
            $@"C:\Program Files (x86)\Autodesk\Revit {yearString}\Revit.exe"
        };
        foreach (var path in commonPaths)
            if (File.Exists(path)) return path;

        // 3. Fall back to the currently running Revit if the version-matched one isn't found.
        Debug.WriteLine($"[OpenRvtDocumentInNewSession] Revit {yearString} not found; falling back to current.");
        return Process.GetCurrentProcess().MainModule?.FileName;
    }
}
