using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
[CommandMeta("")]
public class SelectByWorksetsInNetwork : IExternalCommand
{
    /// <summary>
    /// Marks this command as usable outside Revit context via network.
    /// </summary>
    public static bool IsNetworkCommand => true;

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;

        try
        {
            string token = NetworkClient.GetSharedToken();
            if (token == null)
            {
                TaskDialog.Show("Error", "Network token not found. Ensure at least one Revit session is running.");
                return Result.Failed;
            }

            // Get active documents from registry
            var documents = DocumentRegistry.GetActiveDocuments();
            if (documents.Count == 0)
            {
                TaskDialog.Show("Error", "No active documents found in registry.");
                return Result.Failed;
            }

            // Step 1: Show documents DataGrid - let user select which documents to query
            var documentGridData = new List<Dictionary<string, object>>();
            foreach (var doc in documents)
            {
                var row = new Dictionary<string, object>
                {
                    ["Document"] = doc.DocumentTitle,
                    ["Session ID"] = doc.SessionId,
                    ["Port"] = doc.Port,
                    ["Hostname"] = doc.Hostname,
                    ["Last Heartbeat"] = FormatHeartbeat(doc.LastHeartbeat),
                    ["_Document"] = doc // Hidden field
                };
                documentGridData.Add(row);
            }

            // Sort by Document column
            documentGridData = documentGridData.OrderBy(row => row["Document"].ToString()).ToList();

            var documentColumns = new List<string> { "Document", "Session ID", "Port", "Hostname", "Last Heartbeat" };
            var selectedDocuments = CustomGUIs.DataGrid(documentGridData, documentColumns, false);

            if (selectedDocuments == null || selectedDocuments.Count == 0)
                return Result.Cancelled;

            // Extract selected document objects
            var documentsToQuery = selectedDocuments.Select(row => (DocumentEntry)row["_Document"]).ToList();

            // Step 2: Query selected documents for workset COUNTS only (fast)
            string currentSessionId = RevitBalletApplication.SessionId;
            var worksetCounts = QueryDocumentsForWorksetCounts(documentsToQuery, token, currentSessionId, uiapp);

            if (worksetCounts.Count == 0)
            {
                TaskDialog.Show("No Worksets", "No worksets found in selected documents (documents may not be workshared).");
                return Result.Cancelled;
            }

            // Step 3: Build DataGrid with worksets as rows and documents as columns
            var documentTitles = documentsToQuery.Select(d => d.DocumentTitle).Distinct().ToList();

            var worksetList = new List<Dictionary<string, object>>();

            foreach (var workset in worksetCounts.Keys.OrderBy(w => GetTypeSortOrder(w.Type)).ThenBy(w => w.Name))
            {
                var entry = new Dictionary<string, object>
                {
                    { "Type", workset.Type },
                    { "Workset", workset.Name },
                    { "_WorksetKey", workset }
                };

                // Add count for each document
                foreach (string docTitle in documentTitles)
                {
                    int count = 0;
                    if (worksetCounts[workset].ContainsKey(docTitle))
                    {
                        count = worksetCounts[workset][docTitle];
                    }
                    entry[docTitle] = count;
                }

                worksetList.Add(entry);
            }

            // Define properties to display
            var propertyNames = new List<string> { "Type", "Workset" };
            propertyNames.AddRange(documentTitles);

            // Step 4: Show workset DataGrid
            List<Dictionary<string, object>> selectedWorksets = CustomGUIs.DataGrid(worksetList, propertyNames, false);
            if (selectedWorksets == null || selectedWorksets.Count == 0)
                return Result.Cancelled;

            // Step 5: Query for actual elements in selected worksets
            var selectedWorksetKeys = selectedWorksets
                .Where(w => w.ContainsKey("_WorksetKey"))
                .Select(w => (WorksetKey)w["_WorksetKey"])
                .ToList();

            var worksetElements = QueryElementsForWorksets(documentsToQuery, selectedWorksetKeys, token, currentSessionId, uiapp);

            // Step 6: Gather selection items from query results
            List<SelectionItem> selectionItems = new List<SelectionItem>();

            foreach (var worksetEntry in worksetElements)
            {
                foreach (var docEntry in worksetEntry.Value)
                {
                    string docTitle = docEntry.Key;
                    foreach (var elemInfo in docEntry.Value)
                    {
                        selectionItems.Add(new SelectionItem
                        {
                            DocumentTitle = docTitle,
                            DocumentPath = elemInfo.DocumentPath,
                            UniqueId = elemInfo.UniqueId,
                            ElementIdValue = elemInfo.ElementIdValue,
                            SessionId = elemInfo.SessionId
                        });
                    }
                }
            }

            // Load existing selection and merge
            var existingSelection = SelectionStorage.LoadSelection();
            var existingUniqueIds = new HashSet<string>(existingSelection.Select(s => $"{s.DocumentTitle}|{s.UniqueId}"));

            // Add new items that don't already exist
            foreach (var item in selectionItems)
            {
                string key = $"{item.DocumentTitle}|{item.UniqueId}";
                if (!existingUniqueIds.Contains(key))
                {
                    existingSelection.Add(item);
                }
            }

            // Save selection
            SelectionStorage.SaveSelection(existingSelection);

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Failed to query network sessions: {ex.Message}");
            return Result.Failed;
        }
    }

    private int GetTypeSortOrder(string type)
    {
        switch (type)
        {
            case "User": return 1;
            case "Standard": return 2;
            case "Family": return 3;
            case "View": return 4;
            default: return 5;
        }
    }

    private Dictionary<WorksetKey, Dictionary<string, int>> QueryDocumentsForWorksetCounts(
        List<DocumentEntry> documents, string token, string currentSessionId, UIApplication uiapp)
    {
        // WorksetKey -> Document Title -> Element Count
        var result = new Dictionary<WorksetKey, Dictionary<string, int>>();

        // Process current session locally
        foreach (var docInfo in documents.Where(d => d.SessionId == currentSessionId))
        {
            try
            {
                var app = uiapp.Application;
                foreach (Document doc in app.Documents)
                {
                    if (doc.IsLinked || doc.Title != docInfo.DocumentTitle) continue;
                    if (!doc.IsWorkshared) continue;

                    // Collect elements by workset
                    var collector = new FilteredElementCollector(doc).WhereElementIsNotElementType();
                    var worksetElementCounts = new Dictionary<WorksetId, int>();

                    foreach (Element e in collector)
                    {
                        WorksetId wsId = e.WorksetId;
                        if (wsId == WorksetId.InvalidWorksetId) continue;
                        if (!worksetElementCounts.ContainsKey(wsId))
                            worksetElementCounts[wsId] = 0;
                        worksetElementCounts[wsId]++;
                    }

                    foreach (var pair in worksetElementCounts)
                    {
                        Workset ws = doc.GetWorksetTable().GetWorkset(pair.Key);
                        if (ws == null) continue;

                        var key = new WorksetKey { Name = ws.Name, Type = GetWorksetTypeString(ws.Kind) };
                        if (!result.ContainsKey(key))
                            result[key] = new Dictionary<string, int>();

                        result[key][doc.Title] = pair.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warn("SelectByWorksetsInNetwork", $"Failed to query local document: {ex.Message}");
            }
        }

        // Process remote documents via Roslyn - parallel requests through the shared client
        var remoteDocuments = documents.Where(d => d.SessionId != currentSessionId && !string.IsNullOrWhiteSpace(d.DocumentTitle)).ToList();
        if (remoteDocuments.Count > 0)
        {
            var responses = NetworkClient.ExecuteOnDocuments(remoteDocuments, docInfo =>
            {
                var escapedTitle = NetworkClient.EscapeForScript(docInfo.DocumentTitle);
                return $@"var docTitle = ""{escapedTitle}"";
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
else if (!targetDoc.IsWorkshared)
{{
    Console.WriteLine(""INFO|Document is not workshared"");
}}
else
{{
    var collector = new FilteredElementCollector(targetDoc).WhereElementIsNotElementType();
    var worksetElementCounts = new Dictionary<WorksetId, int>();
    foreach (Element e in collector)
    {{
        WorksetId wsId = e.WorksetId;
        if (wsId == WorksetId.InvalidWorksetId) continue;
        if (!worksetElementCounts.ContainsKey(wsId)) worksetElementCounts[wsId] = 0;
        worksetElementCounts[wsId]++;
    }}
    foreach (var pair in worksetElementCounts)
    {{
        Workset ws = targetDoc.GetWorksetTable().GetWorkset(pair.Key);
        if (ws == null) continue;
        string wsType = """";
        switch (ws.Kind)
        {{
            case WorksetKind.UserWorkset: wsType = ""User""; break;
            case WorksetKind.StandardWorkset: wsType = ""Standard""; break;
            case WorksetKind.FamilyWorkset: wsType = ""Family""; break;
            case WorksetKind.ViewWorkset: wsType = ""View""; break;
            default: wsType = ws.Kind.ToString(); break;
        }}
        Console.WriteLine(""WORKSET|"" + ws.Name + ""|"" + wsType + ""|"" + pair.Value);
    }}
}}";
            }, token);

            foreach (var (docInfo, response) in responses)
            {
                if (response != null && response.Success && !string.IsNullOrWhiteSpace(response.Output))
                {
                    ParseWorksetCountResponse(response.Output, docInfo, result);
                }
            }
        }

        return result;
    }

    private string GetWorksetTypeString(WorksetKind kind)
    {
        switch (kind)
        {
            case WorksetKind.UserWorkset: return "User";
            case WorksetKind.StandardWorkset: return "Standard";
            case WorksetKind.FamilyWorkset: return "Family";
            case WorksetKind.ViewWorkset: return "View";
            default: return kind.ToString();
        }
    }

    private void ParseWorksetCountResponse(string output, DocumentEntry docInfo,
        Dictionary<WorksetKey, Dictionary<string, int>> result)
    {
        foreach (var line in output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (line.StartsWith("WORKSET|"))
            {
                var parts = line.Substring(8).Split('|');
                if (parts.Length >= 3)
                {
                    var key = new WorksetKey
                    {
                        Name = parts[0],
                        Type = parts[1]
                    };
                    int count = int.Parse(parts[2]);

                    if (!result.ContainsKey(key))
                        result[key] = new Dictionary<string, int>();

                    result[key][docInfo.DocumentTitle] = count;
                }
            }
        }
    }

    private Dictionary<WorksetKey, Dictionary<string, List<ElementInfo>>> QueryElementsForWorksets(
        List<DocumentEntry> documents, List<WorksetKey> worksetKeys, string token, string currentSessionId, UIApplication uiapp)
    {
        var result = new Dictionary<WorksetKey, Dictionary<string, List<ElementInfo>>>();

        foreach (var key in worksetKeys)
            result[key] = new Dictionary<string, List<ElementInfo>>();

        var worksetNames = new HashSet<string>(worksetKeys.Select(w => w.Name));

        // Process current session locally
        foreach (var docInfo in documents.Where(d => d.SessionId == currentSessionId))
        {
            try
            {
                var app = uiapp.Application;
                foreach (Document doc in app.Documents)
                {
                    if (doc.IsLinked || doc.Title != docInfo.DocumentTitle) continue;
                    if (!doc.IsWorkshared) continue;

                    // Build workset name to key mapping
                    var worksetNameToKey = new Dictionary<string, WorksetKey>();
                    foreach (var key in worksetKeys)
                    {
                        worksetNameToKey[key.Name] = key;
                    }

                    var collector = new FilteredElementCollector(doc).WhereElementIsNotElementType();

                    foreach (Element e in collector)
                    {
                        WorksetId wsId = e.WorksetId;
                        if (wsId == WorksetId.InvalidWorksetId) continue;

                        Workset ws = doc.GetWorksetTable().GetWorkset(wsId);
                        if (ws == null || !worksetNames.Contains(ws.Name)) continue;

                        var key = worksetNameToKey[ws.Name];

                        if (!result[key].ContainsKey(doc.Title))
                            result[key][doc.Title] = new List<ElementInfo>();

                        result[key][doc.Title].Add(new ElementInfo
                        {
                            UniqueId = e.UniqueId,
#if REVIT2024 || REVIT2025 || REVIT2026
                            ElementIdValue = (int)e.Id.Value,
#else
                            ElementIdValue = e.Id.IntegerValue,
#endif
                            DocumentPath = doc.PathName ?? doc.Title,
                            SessionId = docInfo.SessionId
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warn("SelectByWorksetsInNetwork", $"Failed to query elements for local document: {ex.Message}");
            }
        }

        // Process remote documents via Roslyn - parallel requests through the shared client
        var remoteDocuments = documents.Where(d => d.SessionId != currentSessionId && !string.IsNullOrWhiteSpace(d.DocumentTitle)).ToList();
        if (remoteDocuments.Count > 0)
        {
            // Build workset names filter for query
            var worksetNamesArray = "new HashSet<string> { " + string.Join(", ", worksetKeys.Select(w => $"\"{NetworkClient.EscapeForScript(w.Name)}\"")) + " }";

            var responses = NetworkClient.ExecuteOnDocuments(remoteDocuments, docInfo =>
            {
                var escapedTitle = NetworkClient.EscapeForScript(docInfo.DocumentTitle);
                return $@"var docTitle = ""{escapedTitle}"";
var worksetNames = {worksetNamesArray};
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
else if (!targetDoc.IsWorkshared)
{{
    Console.WriteLine(""INFO|Document is not workshared"");
}}
else
{{
    var collector = new FilteredElementCollector(targetDoc).WhereElementIsNotElementType();
    foreach (Element e in collector)
    {{
        WorksetId wsId = e.WorksetId;
        if (wsId == WorksetId.InvalidWorksetId) continue;
        Workset ws = targetDoc.GetWorksetTable().GetWorkset(wsId);
        if (ws == null || !worksetNames.Contains(ws.Name)) continue;
        Console.WriteLine(""ELEMENT|"" + ws.Name + ""|"" + e.UniqueId + ""|"" + e.Id.IntegerValue);
    }}
}}";
            }, token);

            foreach (var (docInfo, response) in responses)
            {
                if (response != null && response.Success && !string.IsNullOrWhiteSpace(response.Output))
                {
                    ParseElementsResponse(response.Output, docInfo, worksetKeys, result);
                }
            }
        }

        return result;
    }

    private void ParseElementsResponse(string output, DocumentEntry docInfo, List<WorksetKey> worksetKeys,
        Dictionary<WorksetKey, Dictionary<string, List<ElementInfo>>> result)
    {
        foreach (var line in output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (line.StartsWith("ELEMENT|"))
            {
                var parts = line.Substring(8).Split('|');
                if (parts.Length >= 3)
                {
                    string worksetName = parts[0];

                    var matchingKey = worksetKeys.FirstOrDefault(k => k.Name == worksetName);
                    if (matchingKey != null)
                    {
                        if (!result[matchingKey].ContainsKey(docInfo.DocumentTitle))
                            result[matchingKey][docInfo.DocumentTitle] = new List<ElementInfo>();

                        result[matchingKey][docInfo.DocumentTitle].Add(new ElementInfo
                        {
                            UniqueId = parts[1],
                            ElementIdValue = int.Parse(parts[2]),
                            DocumentPath = docInfo.DocumentPath ?? docInfo.DocumentTitle,
                            SessionId = docInfo.SessionId
                        });
                    }
                }
            }
        }
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

    private class WorksetKey
    {
        public string Name { get; set; }
        public string Type { get; set; }

        public override bool Equals(object obj)
        {
            if (obj is WorksetKey other)
                return Name == other.Name && Type == other.Type;
            return false;
        }

        public override int GetHashCode()
        {
            return (Name ?? "").GetHashCode() ^ (Type ?? "").GetHashCode();
        }
    }

    private class ElementInfo
    {
        public string UniqueId { get; set; }
        public int ElementIdValue { get; set; }
        public string DocumentPath { get; set; }
        public string SessionId { get; set; }
    }
}
