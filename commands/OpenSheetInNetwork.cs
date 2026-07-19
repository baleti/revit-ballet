using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

namespace RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
[CommandMeta("")]
public class OpenSheetInNetwork : IExternalCommand
{
    public static bool IsNetworkCommand => true;

    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc proc, IntPtr lParam);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    private const int SW_RESTORE = 9;

    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;

        string token = NetworkClient.GetSharedToken();
        if (token == null)
        {
            TaskDialog.Show("Error", "Network token not found. Ensure Revit Ballet server is running.");
            return Result.Failed;
        }

        var sessions = DocumentRegistry.GetActiveDocuments();
        if (sessions.Count == 0)
        {
            TaskDialog.Show("Error", "No active sessions found in registry.");
            return Result.Failed;
        }

        string currentSessionId = RevitBalletApplication.SessionId;

        var gridData = new List<Dictionary<string, object>>();

        string sheetQuery = @"
var sheets = new FilteredElementCollector(Doc)
    .OfClass(typeof(ViewSheet)).Cast<ViewSheet>()
    .ToList();

foreach (var s in sheets)
{
    Console.WriteLine(""SHEET|"" + s.Id.IntegerValue + ""|"" + s.SheetNumber + ""|"" + s.Name);
}
";

        foreach (var session in sessions)
        {
            if (session.SessionId == currentSessionId)
            {
                Document localDoc = null;
                foreach (Document doc in uiApp.Application.Documents)
                {
                    if (!doc.IsLinked && !doc.IsFamilyDocument &&
                        (doc.Title == session.DocumentTitle || doc.PathName == session.DocumentPath))
                    {
                        localDoc = doc;
                        break;
                    }
                }

                if (localDoc != null)
                {
                    var sheets = new FilteredElementCollector(localDoc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .ToList();

                    foreach (var s in sheets)
                    {
                        gridData.Add(new Dictionary<string, object>
                        {
                            ["Document"] = session.DocumentTitle,
                            ["SheetNumber"] = s.SheetNumber,
                            ["Name"] = s.Name,
                            ["_SessionId"] = session.SessionId,
                            ["_ProcessId"] = session.ProcessId,
#if REVIT2024 || REVIT2025 || REVIT2026
                            ["_ElementId"] = (int)s.Id.Value,
#else
                            ["_ElementId"] = s.Id.IntegerValue,
#endif
                            ["_IsLocal"] = true
                        });
                    }
                }
            }
            else
            {
                try
                {
                    var response = SendRoslynQuery(session.Port.ToString(), token, sheetQuery);
                    if (response != null && response.Success && !string.IsNullOrEmpty(response.Output))
                    {
                        foreach (var line in response.Output.Split('\n'))
                        {
                            if (!line.StartsWith("SHEET|")) continue;
                            var parts = line.Split('|');
                            if (parts.Length < 4) continue;

                            gridData.Add(new Dictionary<string, object>
                            {
                                ["Document"] = session.DocumentTitle,
                                ["SheetNumber"] = parts[2],
                                ["Name"] = parts[3].TrimEnd('\r'),
                                ["_SessionId"] = session.SessionId,
                                ["_ProcessId"] = session.ProcessId,
                                ["_ElementId"] = int.Parse(parts[1]),
                                ["_IsLocal"] = false
                            });
                        }
                    }
                }
                catch { /* skip unreachable sessions */ }
            }
        }

        if (gridData.Count == 0)
        {
            TaskDialog.Show("Info", "No sheets found in any network sessions.");
            return Result.Failed;
        }

        gridData = gridData.OrderBy(r => r["Document"].ToString())
                           .ThenBy(r => r["SheetNumber"].ToString())
                           .ToList();

        var columns = new List<string> { "Document", "SheetNumber", "Name" };

        CustomGUIs.SetCurrentUIDocument(uiApp.ActiveUIDocument);
        var selectedRows = CustomGUIs.DataGrid(gridData, columns, false);

        if (selectedRows == null || selectedRows.Count == 0)
            return Result.Succeeded;

        foreach (var row in selectedRows)
        {
            bool isLocal = Convert.ToBoolean(row["_IsLocal"]);
            int elementIdValue = Convert.ToInt32(row["_ElementId"]);
            int processId = Convert.ToInt32(row["_ProcessId"]);

            if (isLocal)
            {
                Document localDoc = null;
                string docTitle = row["Document"].ToString();
                foreach (Document doc in uiApp.Application.Documents)
                {
                    if (!doc.IsLinked && !doc.IsFamilyDocument && doc.Title == docTitle)
                    {
                        localDoc = doc;
                        break;
                    }
                }
                if (localDoc != null)
                {
                    ViewSheet sheet = localDoc.GetElement(elementIdValue.ToElementId()) as ViewSheet;
                    if (sheet != null)
                        uiApp.ActiveUIDocument.RequestViewChange(sheet);
                }
            }
            else
            {
                string openScript = $@"
var sheet = Doc.GetElement(new ElementId({elementIdValue})) as ViewSheet;
if (sheet != null) UIDoc.RequestViewChange(sheet);
Console.WriteLine(""OPENED|"" + {elementIdValue});
";
                try { SendRoslynQuery(sessions.First(s => s.ProcessId == processId).Port.ToString(), token, openScript); }
                catch { }

                IntPtr window = FindMainWindow(processId);
                if (window != IntPtr.Zero)
                {
                    if (IsIconic(window)) ShowWindow(window, SW_RESTORE);
                    SetForegroundWindow(window);
                }
            }
        }

        return Result.Succeeded;
    }

    private IntPtr FindMainWindow(int processId)
    {
        IntPtr result = IntPtr.Zero;
        EnumWindows((hWnd, _) =>
        {
            uint pid;
            GetWindowThreadProcessId(hWnd, out pid);
            if (pid == (uint)processId && IsWindowVisible(hWnd))
            {
                result = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }

    private RoslynResult SendRoslynQuery(string port, string token, string script)
    {
        return NetworkClient.ExecuteScript(int.Parse(port), script, token);
    }

}
