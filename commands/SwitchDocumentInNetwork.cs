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
    /// <summary>
    /// Marks this command as usable outside Revit context via network.
    /// </summary>
    public static bool IsNetworkCommand => true;

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
                // Get active documents from registry
                var documents = DocumentRegistry.GetActiveDocuments();

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
                var columns = new List<string> { "Document", "Last Transaction", "Last Sync", "Session ID", "Port", "Hostname", "Last Heartbeat" };

                foreach (var doc in documents)
                {
                    var row = new Dictionary<string, object>
                    {
                        ["Document"] = string.IsNullOrWhiteSpace(doc.DocumentTitle) ? "(Untitled)" : doc.DocumentTitle,
                        ["Last Transaction"] = FormatLastTransaction(doc.LastTransaction),
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

    private string FormatLastTransaction(DateTime? lastTransaction)
    {
        if (!lastTransaction.HasValue)
            return "-";

        var timeAgo = DateTime.Now - lastTransaction.Value;

        if (timeAgo.TotalSeconds < 60)
            return "Just now";
        else if (timeAgo.TotalMinutes < 60)
            return $"{(int)timeAgo.TotalMinutes}m ago";
        else if (timeAgo.TotalHours < 24)
            return $"{(int)timeAgo.TotalHours}h ago";
        else
            return $"{(int)timeAgo.TotalDays}d ago";
    }

}
