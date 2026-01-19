using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;

/// <summary>
/// Lists selected elements across all sessions in the network.
/// Uses SelectionStorage for cross-session selection tracking.
/// </summary>
[Transaction(TransactionMode.Manual)]
public class ListSelectedInNetwork : IExternalCommand
{
    /// <summary>
    /// Marks this command as usable outside Revit context via network.
    /// </summary>
    public static bool IsNetworkCommand => true;

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document activeDoc = uidoc.Document;

            // Load selection from SelectionStorage
            var storedSelection = SelectionStorage.LoadSelection();

            if (!storedSelection.Any())
            {
                TaskDialog.Show("Info", "No elements in network selection.\n\nUse network-scope commands like SelectByCategoriesInNetwork to build a cross-network selection first.");
                return Result.Cancelled;
            }

            // DEBUG: Log what we loaded
            System.Diagnostics.Debug.WriteLine($"Loaded {storedSelection.Count} elements from SelectionStorage");
            foreach (var item in storedSelection.Take(5))
            {
                System.Diagnostics.Debug.WriteLine($"  - Doc: {item.DocumentTitle}, SessionId: {item.SessionId ?? "NULL"}, UniqueId: {item.UniqueId}");
            }

            // Read network token
            string tokenPath = Path.Combine(PathHelper.RuntimeDirectory, "network", "token");
            if (!File.Exists(tokenPath))
            {
                TaskDialog.Show("Error", "Network token not found.");
                return Result.Failed;
            }
            string token = File.ReadAllText(tokenPath).Trim();

            string currentSessionId = RevitBallet.RevitBallet.SessionId;

            // Collect element data from all sessions
            List<Dictionary<string, object>> elementData = new List<Dictionary<string, object>>();

            using (var progress = new CancellableProgressDialog("Collecting element data"))
            {
                progress.Start();
                progress.SetTotal(storedSelection.Count);

                // Group by SessionId for efficient processing
                var bySession = storedSelection.GroupBy(s => s.SessionId ?? currentSessionId);

                System.Diagnostics.Debug.WriteLine($"Grouped into {bySession.Count()} session groups");
                foreach (var sessionGroup in bySession)
                {
                    string sessionId = sessionGroup.Key;
                    System.Diagnostics.Debug.WriteLine($"Processing session {sessionId} with {sessionGroup.Count()} elements (current: {currentSessionId})");

                    if (sessionId == currentSessionId)
                    {
                        // Process current session locally
                        var byDocument = sessionGroup.GroupBy(s => s.DocumentTitle);

                        foreach (var docGroup in byDocument)
                        {
                            // Find matching open document
                            Document doc = null;
                            foreach (Document d in app.Documents)
                            {
                                if (d.IsLinked) continue;
                                if (d.Title == docGroup.Key)
                                {
                                    doc = d;
                                    break;
                                }
                            }

                            if (doc == null)
                            {
                                System.Diagnostics.Debug.WriteLine($"  Document '{docGroup.Key}' not found in current session");
                                foreach (var _ in docGroup)
                                {
                                    progress.IncrementProgress();
                                }
                                continue;
                            }

                            System.Diagnostics.Debug.WriteLine($"  Found document '{doc.Title}' in current session");

                            // Process elements in this document
                            foreach (var selItem in docGroup)
                            {
                                progress.CheckAndShow();

                                if (progress.IsCancelled)
                                    throw new OperationCanceledException("Operation cancelled by user.");

                                try
                                {
                                    Element elem = doc.GetElement(selItem.UniqueId);
                                    if (elem != null)
                                    {
                                        var data = BuildElementDataDictionary(elem, doc);
                                        elementData.Add(data);
                                    }
                                }
                                catch { /* Skip problematic elements */ }

                                progress.IncrementProgress();
                            }
                        }
                    }
                    else
                    {
                        // Process remote session via Roslyn
                        System.Diagnostics.Debug.WriteLine($"  Querying remote session {sessionId}");
                        var remoteData = QueryRemoteSessionElements(sessionGroup.ToList(), token);
                        System.Diagnostics.Debug.WriteLine($"  Got {remoteData.Count} elements from remote session");
                        elementData.AddRange(remoteData);

                        foreach (var _ in sessionGroup)
                        {
                            progress.IncrementProgress();
                        }
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine($"Total element data collected: {elementData.Count}");

            if (!elementData.Any())
            {
                // Write diagnostic file
                try
                {
                    string diagnosticPath = Path.Combine(
                        PathHelper.RuntimeDirectory,
                        "diagnostics",
                        $"ListSelectedInNetwork-NoElements-{DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");
                    Directory.CreateDirectory(Path.GetDirectoryName(diagnosticPath));

                    var lines = new List<string>();
                    lines.Add($"ListSelectedInNetwork - No elements found");
                    lines.Add($"Current Session ID: {currentSessionId}");
                    lines.Add($"Total items in SelectionStorage: {storedSelection.Count}");
                    lines.Add("");
                    lines.Add("Session groups:");
                    foreach (var group in storedSelection.GroupBy(s => s.SessionId ?? currentSessionId))
                    {
                        lines.Add($"  Session {group.Key}: {group.Count()} elements");
                        foreach (var doc in group.GroupBy(e => e.DocumentTitle))
                        {
                            lines.Add($"    Document '{doc.Key}': {doc.Count()} elements");
                        }
                    }

                    File.WriteAllLines(diagnosticPath, lines);
                }
                catch { }

                TaskDialog.Show("Info", "None of the selected elements are available.");
                return Result.Cancelled;
            }

            // Get ALL property names from ALL elements
            var allPropertyNames = new HashSet<string>();
            foreach (var data in elementData)
            {
                foreach (var key in data.Keys)
                {
                    if (!key.EndsWith("Object"))
                    {
                        allPropertyNames.Add(key);
                    }
                }
            }

            // Build ordered list with standard columns
            var orderedProps = new List<string> { "Name" };
            if (allPropertyNames.Contains("Document")) orderedProps.Add("Document");
            if (allPropertyNames.Contains("Type Name")) orderedProps.Add("Type Name");
            if (allPropertyNames.Contains("Family")) orderedProps.Add("Family");
            orderedProps.Add("Category");
            if (allPropertyNames.Contains("Group")) orderedProps.Add("Group");
            if (allPropertyNames.Contains("OwnerView")) orderedProps.Add("OwnerView");
            orderedProps.Add("Id");

            var remainingProps = allPropertyNames.Except(orderedProps).OrderBy(p => p);
            var propertyNames = orderedProps.Where(p => allPropertyNames.Contains(p))
                .Concat(remainingProps)
                .ToList();

            // Set the current UIDocument for edit operations (only works for local session)
            CustomGUIs.SetCurrentUIDocument(uidoc);

            var chosenRows = CustomGUIs.DataGrid(elementData, propertyNames, false);

            // Apply edits (only affects local session elements)
            if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
            {
                CustomGUIs.ApplyCellEditsToEntities();
            }

            if (chosenRows.Count == 0)
            {
                SelectionStorage.ClearSelection();
                return Result.Cancelled;
            }

            // Update SelectionStorage with filtered selection
            var newSelection = new List<SelectionItem>();

            foreach (var row in chosenRows)
            {
                if (row.TryGetValue("UniqueId", out var uniqueIdObj) && uniqueIdObj is string uniqueId &&
                    row.TryGetValue("DocumentTitle", out var docTitleObj) && docTitleObj is string docTitle &&
                    row.TryGetValue("DocumentPath", out var docPathObj) && docPathObj is string docPath &&
                    row.TryGetValue("Id", out var idObj))
                {
                    long idValue = 0;
                    if (idObj is int intId)
                        idValue = intId;
                    else if (idObj is long longId)
                        idValue = longId;

                    string sessionId = null;
                    if (row.TryGetValue("SessionId", out var sessionIdObj) && sessionIdObj is string sid)
                    {
                        sessionId = sid;
                    }

                    newSelection.Add(new SelectionItem
                    {
                        DocumentTitle = docTitle,
                        DocumentPath = docPath,
                        UniqueId = uniqueId,
                        ElementIdValue = (int)idValue,
                        SessionId = sessionId
                    });
                }
            }

            SelectionStorage.SaveSelection(newSelection);

            return Result.Succeeded;
        }
        catch (OperationCanceledException)
        {
            message = "Operation cancelled by user.";
            return Result.Cancelled;
        }
        catch (Exception ex)
        {
            message = $"Unexpected error: {ex.Message}";
            return Result.Failed;
        }
    }

    private List<Dictionary<string, object>> QueryRemoteSessionElements(List<SelectionItem> elements, string token)
    {
        var result = new List<Dictionary<string, object>>();

        if (!elements.Any()) return result;

        // Group by document (should all be same session but might be different documents)
        var byDocument = elements.GroupBy(e => new { e.DocumentTitle, e.SessionId });

        using (var handler = new HttpClientHandler())
        {
            handler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;

            using (var client = new HttpClient(handler))
            {
                client.Timeout = TimeSpan.FromSeconds(120);

                foreach (var docGroup in byDocument)
                {
                    try
                    {
                        // Find document info from documents registry
                        string documentsPath = Path.Combine(PathHelper.RuntimeDirectory, "documents");
                        var documents = ParseDocumentsFile(documentsPath);
                        // Find any document entry with matching session ID (they all have same port/hostname)
                        var docInfo = documents.FirstOrDefault(d => d.SessionId == docGroup.Key.SessionId);

                        if (docInfo == null) continue;

                        // Build query to get element data
                        var uniqueIds = docGroup.Select(e => $"\"{e.UniqueId}\"").ToList();
                        var uniqueIdsArray = "{ " + string.Join(", ", uniqueIds) + " }";
                        var escapedDocTitle = docGroup.Key.DocumentTitle.Replace("\\", "\\\\").Replace("\"", "\\\"");

                        // Must find specific document by title - Doc may point to different active document
                        var query = $@"var docTitle = ""{escapedDocTitle}"";
var uniqueIds = new string[] {uniqueIdsArray};
Document targetDoc = null;
foreach (Document d in UIApp.Application.Documents)
{{
    if (!d.IsLinked && d.Title == docTitle)
    {{
        targetDoc = d;
        break;
    }}
}}
if (targetDoc == null)
{{
    Console.WriteLine(""ERROR|Document not found: "" + docTitle);
}}
else
{{
    foreach (var uid in uniqueIds)
    {{
        try
        {{
            var elem = targetDoc.GetElement(uid);
            if (elem != null)
            {{
                var parts = new List<string>();
                parts.Add(""ELEMENT|"" + uid);
                parts.Add(""Name|"" + (elem.Name ?? """"));
                parts.Add(""Category|"" + (elem.Category?.Name ?? """"));
                parts.Add(""Id|"" + elem.Id.IntegerValue);

                // Get Type Name and Family
                var typeId = elem.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {{
                    var typeElem = targetDoc.GetElement(typeId);
                    if (typeElem != null)
                    {{
                        parts.Add(""Type Name|"" + (typeElem.Name ?? """"));
                        if (typeElem is FamilySymbol fs)
                        {{
                            parts.Add(""Family|"" + (fs.FamilyName ?? """"));
                        }}
                    }}
                }}

                Console.WriteLine(string.Join(""||"", parts));
            }}
        }}
        catch {{ }}
    }}
}}";

                        var content = new StringContent(query, Encoding.UTF8, "text/plain");
                        var request = new HttpRequestMessage(HttpMethod.Post, $"https://127.0.0.1:{docInfo.Port}/roslyn");
                        request.Headers.Add("X-Auth-Token", token);
                        request.Content = content;

                        var response = client.SendAsync(request).Result;
                        var responseText = response.Content.ReadAsStringAsync().Result;

                        var jsonResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<RoslynResponse>(responseText);

                        if (jsonResponse.Success && !string.IsNullOrWhiteSpace(jsonResponse.Output))
                        {
                            result.AddRange(ParseRemoteElementData(jsonResponse.Output, docGroup.Key.DocumentTitle, docInfo.DocumentPath, docInfo.SessionId));
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to query remote session: {ex.Message}");
                    }
                }
            }
        }

        return result;
    }

    private List<Dictionary<string, object>> ParseRemoteElementData(string output, string documentTitle, string documentPath, string sessionId)
    {
        var result = new List<Dictionary<string, object>>();

        foreach (var line in output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (line.StartsWith("ELEMENT|"))
            {
                var data = new Dictionary<string, object>
                {
                    ["Document"] = documentTitle,
                    ["DocumentTitle"] = documentTitle,
                    ["DocumentPath"] = documentPath ?? documentTitle,
                    ["SessionId"] = sessionId
                };

                var parts = line.Split(new[] { "||" }, StringSplitOptions.None);
                foreach (var part in parts)
                {
                    var keyValue = part.Split(new[] { '|' }, 2);
                    if (keyValue.Length == 2)
                    {
                        string key = keyValue[0];
                        string value = keyValue[1];

                        if (key == "ELEMENT")
                        {
                            data["UniqueId"] = value;
                        }
                        else if (key == "Id")
                        {
                            if (long.TryParse(value, out long id))
                            {
                                data["Id"] = id;
                            }
                        }
                        else
                        {
                            data[key] = value;
                        }
                    }
                }

                result.Add(data);
            }
        }

        return result;
    }

    private List<DocumentInfo> ParseDocumentsFile(string filePath)
    {
        var documents = new List<DocumentInfo>();

        try
        {
            var lines = File.ReadAllLines(filePath);

            foreach (var line in lines)
            {
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

                // Filter out stale documents (no heartbeat for > 2 minutes)
                if ((DateTime.Now - doc.LastHeartbeat).TotalSeconds <= 120)
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

    /// <summary>
    /// Build element data dictionary for local session elements
    /// </summary>
    private Dictionary<string, object> BuildElementDataDictionary(Element element, Document doc)
    {
        string groupName = string.Empty;
        if (element.GroupId != null && element.GroupId != ElementId.InvalidElementId && element.GroupId.AsLong() != -1)
        {
            if (doc.GetElement(element.GroupId) is Group g)
                groupName = g.Name;
        }

        string ownerViewName = string.Empty;
        if (element.OwnerViewId != null && element.OwnerViewId != ElementId.InvalidElementId)
        {
            if (doc.GetElement(element.OwnerViewId) is View v)
                ownerViewName = v.Name;
        }

        var data = new Dictionary<string, object>
        {
            ["Name"] = element.Name,
            ["Category"] = element.Category?.Name ?? string.Empty,
            ["Group"] = groupName,
            ["OwnerView"] = ownerViewName,
            ["Id"] = element.Id.AsLong(),
            ["ElementIdObject"] = element.Id,
            ["UniqueId"] = element.UniqueId,
            ["Document"] = doc.Title,
            ["DocumentTitle"] = doc.Title,
            ["DocumentPath"] = doc.PathName ?? "",
            ["SessionId"] = RevitBallet.RevitBallet.SessionId
        };

        // Add Type Name and Family
        ElementId typeId = element.GetTypeId();
        if (typeId != null && typeId != ElementId.InvalidElementId)
        {
            Element typeElement = doc.GetElement(typeId);
            if (typeElement != null)
            {
                data["Type Name"] = typeElement.Name;

                if (typeElement is FamilySymbol familySymbol)
                {
                    data["Family"] = familySymbol.FamilyName;
                }
                else
                {
                    Parameter familyParam = typeElement.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                    data["Family"] = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
                        ? familyParam.AsString()
                        : "System Type";
                }
            }
        }

        return data;
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
    }

    private class RoslynResponse
    {
        public bool Success { get; set; }
        public string Output { get; set; }
        public string Error { get; set; }
    }
}
