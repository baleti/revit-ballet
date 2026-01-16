using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SwitchDocumentInNetwork : IExternalCommand
{
    // Windows API imports for bringing window to foreground
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetActiveWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    private const int SW_RESTORE = 9;

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;

        using (var executionLog = CommandExecutionLogger.Start("SwitchDocumentInNetwork", commandData))
        using (var diagnostics = CommandDiagnostics.StartCommand("SwitchDocumentInNetwork", uiApp))
        {
            try
            {
                // Read documents file from runtime directory
                string documentsFilePath = Path.Combine(PathHelper.RuntimeDirectory, "documents");

                if (!File.Exists(documentsFilePath))
                {
                    TaskDialog.Show("Error", "No active documents found. Document registry file does not exist.");
                    diagnostics.LogError("Documents file not found");
                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }

                // Parse documents file
                var documents = ParseDocumentsFile(documentsFilePath);

                if (documents.Count == 0)
                {
                    TaskDialog.Show("Error", "No active documents found in registry.");
                    diagnostics.LogError("No documents in registry");
                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }

                // Get current session ID for highlighting
                string currentSessionId = RevitBallet.RevitBallet.SessionId;

                // Prepare data for DataGrid
                var gridData = new List<Dictionary<string, object>>();
                var columns = new List<string> { "Document", "Last Sync", "Session ID", "Port", "Hostname", "Last Heartbeat" };

                foreach (var doc in documents)
                {
                    var row = new Dictionary<string, object>
                    {
                        ["Document"] = string.IsNullOrWhiteSpace(doc.DocumentTitle) ? "(Untitled)" : doc.DocumentTitle,
                        ["Last Sync"] = FormatLastSync(doc.LastSync),
                        ["Session ID"] = doc.SessionId,
                        ["Port"] = doc.Port,
                        ["Hostname"] = doc.Hostname,
                        ["Last Heartbeat"] = FormatHeartbeat(doc.LastHeartbeat),
                        ["_ProcessId"] = doc.ProcessId, // Hidden field for later use
                        ["_IsCurrent"] = doc.SessionId == currentSessionId
                    };
                    gridData.Add(row);
                }

                // Sort by Document column
                gridData = gridData.OrderBy(row => row["Document"].ToString()).ToList();

                // Find index of current session for initial selection
                int currentSessionIndex = gridData.FindIndex(row => Convert.ToBoolean(row["_IsCurrent"]));
                List<int> initialSelection = currentSessionIndex >= 0 ? new List<int> { currentSessionIndex } : null;

                // Show selection dialog with current session pre-selected
                var selectedRows = CustomGUIs.DataGrid(gridData, columns, false, initialSelection);

                if (selectedRows == null || selectedRows.Count == 0)
                {
                    diagnostics.Log("User cancelled selection");
                    executionLog.SetResult(Result.Cancelled);
                    return Result.Cancelled;
                }

                // Get selected session's process ID
                var selectedRow = selectedRows[0];
                int targetProcessId = Convert.ToInt32(selectedRow["_ProcessId"]);
                bool isCurrent = Convert.ToBoolean(selectedRow["_IsCurrent"]);

                if (isCurrent)
                {
                    TaskDialog.Show("Info", "Selected session is the current session.");
                    executionLog.SetResult(Result.Succeeded);
                    return Result.Succeeded;
                }

                // Find and bring window to foreground
                IntPtr targetWindow = FindMainWindow(targetProcessId);

                if (targetWindow == IntPtr.Zero)
                {
                    TaskDialog.Show("Error", $"Could not find window for process ID {targetProcessId}.");
                    diagnostics.LogError($"Window not found for PID {targetProcessId}");
                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }

                // Restore if minimized
                if (IsIconic(targetWindow))
                {
                    ShowWindow(targetWindow, SW_RESTORE);
                }

                // Bring to foreground
                bool success = SetForegroundWindow(targetWindow);

                if (success)
                {
                    diagnostics.Log($"Successfully switched to session {selectedRow["Session ID"]} (PID: {targetProcessId})");
                    executionLog.SetResult(Result.Succeeded);
                    return Result.Succeeded;
                }
                else
                {
                    TaskDialog.Show("Warning", "Window was found but could not be brought to foreground. It may be on another desktop or locked.");
                    diagnostics.LogError($"SetForegroundWindow failed for PID {targetProcessId}");
                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to switch session: {ex.Message}");
                diagnostics.LogError($"Exception: {ex}");
                executionLog.SetResult(Result.Failed);
                return Result.Failed;
            }
        }
    }

    private List<DocumentInfo> ParseDocumentsFile(string filePath)
    {
        var documents = new List<DocumentInfo>();

        try
        {
            var lines = File.ReadAllLines(filePath);

            foreach (var line in lines)
            {
                // Skip comments and empty lines
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                    continue;

                var parts = line.Split(',');
                // Format: DocumentTitle,DocumentPath,SessionId,Port,Hostname,ProcessId,RegisteredAt,LastHeartbeat,LastSync
                if (parts.Length < 8)
                    continue;

                var doc = new DocumentInfo
                {
                    DocumentTitle = parts[0].Trim(),
                    DocumentPath = parts[1].Trim(),
                    SessionId = parts[2].Trim(),
                    Port = parts[3].Trim(),
                    Hostname = parts[4].Trim(),
                    ProcessId = int.Parse(parts[5].Trim()),
                    RegisteredAt = DateTime.Parse(parts[6].Trim()),
                    LastHeartbeat = DateTime.Parse(parts[7].Trim())
                };

                // LastSync is optional (index 8)
                if (parts.Length > 8 && !string.IsNullOrWhiteSpace(parts[8]))
                {
                    DateTime lastSyncParsed;
                    if (DateTime.TryParse(parts[8].Trim(), out lastSyncParsed))
                    {
                        doc.LastSync = lastSyncParsed;
                    }
                }

                // Filter out stale documents (no heartbeat for > 2 minutes)
                if ((DateTime.Now - doc.LastHeartbeat).TotalSeconds < 120)
                {
                    documents.Add(doc);
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to parse documents file: {ex.Message}", ex);
        }

        return documents;
    }

    private IntPtr FindMainWindow(int processId)
    {
        IntPtr mainWindow = IntPtr.Zero;

        EnumWindows((hWnd, lParam) =>
        {
            uint windowProcessId;
            GetWindowThreadProcessId(hWnd, out windowProcessId);

            if (windowProcessId == processId && IsWindowVisible(hWnd))
            {
                mainWindow = hWnd;
                return false; // Stop enumeration
            }

            return true; // Continue enumeration
        }, IntPtr.Zero);

        return mainWindow;
    }

    private string FormatHeartbeat(DateTime heartbeat)
    {
        var timeAgo = DateTime.Now - heartbeat;

        if (timeAgo.TotalSeconds < 60)
            return "Just now";
        else if (timeAgo.TotalMinutes < 60)
            return $"{(int)timeAgo.TotalMinutes}m ago";
        else if (timeAgo.TotalHours < 24)
            return $"{(int)timeAgo.TotalHours}h ago";
        else
            return $"{(int)timeAgo.TotalDays}d ago";
    }

    private string FormatLastSync(DateTime? lastSync)
    {
        if (!lastSync.HasValue)
            return "-";

        var timeAgo = DateTime.Now - lastSync.Value;

        if (timeAgo.TotalSeconds < 60)
            return "Just now";
        else if (timeAgo.TotalMinutes < 60)
            return $"{(int)timeAgo.TotalMinutes}m ago";
        else if (timeAgo.TotalHours < 24)
            return $"{(int)timeAgo.TotalHours}h ago";
        else
            return $"{(int)timeAgo.TotalDays}d ago";
    }

    private class DocumentInfo
    {
        public string DocumentTitle { get; set; }
        public string DocumentPath { get; set; }
        public string SessionId { get; set; }
        public string Port { get; set; }
        public string Hostname { get; set; }
        public int ProcessId { get; set; }
        public DateTime RegisteredAt { get; set; }
        public DateTime LastHeartbeat { get; set; }
        public DateTime? LastSync { get; set; }
    }
}
