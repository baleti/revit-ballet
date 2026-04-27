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

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
[CommandMeta("")]
public class OpenViewInNetwork : IExternalCommand
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

        string tokenPath = Path.Combine(PathHelper.RuntimeDirectory, "network", "token");
        if (!File.Exists(tokenPath))
        {
            TaskDialog.Show("Error", "Network token not found. Ensure Revit Ballet server is running.");
            return Result.Failed;
        }
        string token = File.ReadAllText(tokenPath).Trim();

        var sessions = DocumentRegistry.GetActiveDocuments();
        if (sessions.Count == 0)
        {
            TaskDialog.Show("Error", "No active sessions found in registry.");
            return Result.Failed;
        }

        string currentSessionId = RevitBallet.RevitBallet.SessionId;

        // Query each session for its non-sheet views
        var gridData = new List<Dictionary<string, object>>();

        string viewQuery = @"
var viewIdToViewport = new FilteredElementCollector(Doc)
    .OfClass(typeof(Viewport)).Cast<Viewport>()
    .GroupBy(vp => vp.ViewId).ToDictionary(g => g.Key, g => g.First());

var views = new FilteredElementCollector(Doc)
    .OfClass(typeof(View)).Cast<View>()
    .Where(v => !(v is ViewSheet) && !v.IsTemplate &&
                v.ViewType != ViewType.ProjectBrowser &&
                v.ViewType != ViewType.SystemBrowser)
    .ToList();

foreach (var v in views)
{
    string sheetNum = """";
    string sheetTitle = """";
    Viewport vp;
    if (viewIdToViewport.TryGetValue(v.Id, out vp))
    {
        var sheet = Doc.GetElement(vp.SheetId) as ViewSheet;
        if (sheet != null) { sheetNum = sheet.SheetNumber; sheetTitle = sheet.Name; }
    }
    Console.WriteLine(""VIEW|"" + v.Id.IntegerValue + ""|"" + v.Name + ""|"" + v.ViewType + ""|"" + sheetNum + ""|"" + sheetTitle);
}
";

        foreach (var session in sessions)
        {
            if (session.SessionId == currentSessionId)
            {
                // Query local session directly
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
                    var viewIdToViewport = new FilteredElementCollector(localDoc)
                        .OfClass(typeof(Viewport)).Cast<Viewport>()
                        .GroupBy(vp => vp.ViewId).ToDictionary(g => g.Key, g => g.First());

                    var views = new FilteredElementCollector(localDoc)
                        .OfClass(typeof(View)).Cast<View>()
                        .Where(v => !(v is ViewSheet) && !v.IsTemplate &&
                                    v.ViewType != ViewType.ProjectBrowser &&
                                    v.ViewType != ViewType.SystemBrowser)
                        .ToList();

                    foreach (var v in views)
                    {
                        string sheetNum = "";
                        string sheetTitle = "";
                        Viewport vp;
                        if (viewIdToViewport.TryGetValue(v.Id, out vp))
                        {
                            ViewSheet s = localDoc.GetElement(vp.SheetId) as ViewSheet;
                            if (s != null) { sheetNum = s.SheetNumber; sheetTitle = s.Name; }
                        }

                        gridData.Add(new Dictionary<string, object>
                        {
                            ["Document"] = session.DocumentTitle,
                            ["Name"] = v.Name,
                            ["ViewType"] = v.ViewType.ToString(),
                            ["Sheet Number"] = sheetNum,
                            ["Sheet Title"] = sheetTitle,
                            ["_SessionId"] = session.SessionId,
                            ["_ProcessId"] = session.ProcessId,
    #if REVIT2024 || REVIT2025 || REVIT2026
                        ["_ElementId"] = (int)v.Id.Value,
#else
                        ["_ElementId"] = v.Id.IntegerValue,
#endif
                            ["_IsLocal"] = true
                        });
                    }
                }
            }
            else
            {
                // Query remote session via Roslyn
                try
                {
                    var response = SendRoslynQuery(session.Port.ToString(), token, viewQuery);
                    if (response != null && response.Success && !string.IsNullOrEmpty(response.Output))
                    {
                        foreach (var line in response.Output.Split('\n'))
                        {
                            if (!line.StartsWith("VIEW|")) continue;
                            var parts = line.Split('|');
                            if (parts.Length < 6) continue;

                            gridData.Add(new Dictionary<string, object>
                            {
                                ["Document"] = session.DocumentTitle,
                                ["Name"] = parts[2],
                                ["ViewType"] = parts[3],
                                ["Sheet Number"] = parts[4],
                                ["Sheet Title"] = parts[5].TrimEnd('\r'),
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
            TaskDialog.Show("Info", "No views found in any network sessions.");
            return Result.Failed;
        }

        gridData = gridData.OrderBy(r => r["Document"].ToString())
                           .ThenBy(r => r["Name"].ToString())
                           .ToList();

        var columns = new List<string> { "Document", "Name", "ViewType", "Sheet Number", "Sheet Title" };

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
                // Open directly in current session
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
                    View view = localDoc.GetElement(elementIdValue.ToElementId()) as View;
                    if (view != null)
                        uiApp.ActiveUIDocument.RequestViewChange(view);
                }
            }
            else
            {
                // Send open command to remote session, then focus its window
                string openScript = $@"
var view = Doc.GetElement(new ElementId({elementIdValue})) as View;
if (view != null) UIDoc.RequestViewChange(view);
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

    private RoslynResponse SendRoslynQuery(string port, string token, string script)
    {
        string url = $"https://127.0.0.1:{port}/roslyn";
#if NET8_0_OR_GREATER
        using (var handler = new System.Net.Http.HttpClientHandler())
        {
            handler.ServerCertificateCustomValidationCallback = (m, c, ch, e) => true;
            using (var client = new System.Net.Http.HttpClient(handler))
            {
                client.Timeout = TimeSpan.FromSeconds(30);
                client.DefaultRequestHeaders.Add("X-Auth-Token", token);
                var content = new System.Net.Http.StringContent(script, Encoding.UTF8, "text/plain");
                var resp = client.PostAsync(url, content).Result;
                var text = resp.Content.ReadAsStringAsync().Result;
                return Newtonsoft.Json.JsonConvert.DeserializeObject<RoslynResponse>(text);
            }
        }
#else
        var request = (HttpWebRequest)WebRequest.Create(url);
        request.Method = "POST";
        request.ContentType = "text/plain";
        request.Headers.Add("X-Auth-Token", token);
        request.Timeout = 30000;
        request.ServerCertificateValidationCallback = (s, c, ch, e) => true;
        byte[] body = Encoding.UTF8.GetBytes(script);
        request.ContentLength = body.Length;
        using (var stream = request.GetRequestStream())
            stream.Write(body, 0, body.Length);
        using (var response = request.GetResponse())
        using (var reader = new StreamReader(response.GetResponseStream()))
            return Newtonsoft.Json.JsonConvert.DeserializeObject<RoslynResponse>(reader.ReadToEnd());
#endif
    }

    private class RoslynResponse
    {
        public bool Success { get; set; }
        public string Output { get; set; }
        public string Error { get; set; }
    }
}
