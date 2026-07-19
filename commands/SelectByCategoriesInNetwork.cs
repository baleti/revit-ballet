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
public class SelectByCategoriesInNetwork : IExternalCommand
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
            var selectedDocuments = CustomGUIs.DataGrid(documentGridData, documentColumns, false); // Don't span all screens

            if (selectedDocuments == null || selectedDocuments.Count == 0)
                return Result.Cancelled;

            // Extract selected document objects
            var documentsToQuery = selectedDocuments.Select(row => (DocumentEntry)row["_Document"]).ToList();

            // Step 2: Query selected documents for category COUNTS only (fast)
            // Use local API for current session, Roslyn for remote sessions
            string currentSessionId = RevitBalletApplication.SessionId;
            var categoryCounts = QueryDocumentsForCategoryCounts(documentsToQuery, token, currentSessionId, uiapp);

            if (categoryCounts.Count == 0)
            {
                TaskDialog.Show("No Categories", "No categories found in selected documents.");
                return Result.Cancelled;
            }

            // Step 3: Build DataGrid with categories as rows and documents as columns
            var documentTitles = documentsToQuery.Select(d => d.DocumentTitle).Distinct().ToList();

            var categoryList = new List<Dictionary<string, object>>();

            foreach (var category in categoryCounts.Keys.OrderBy(c => c))
            {
                var entry = new Dictionary<string, object>
                {
                    { "Category", category },
                    { "CategoryName", category }
                };

                // Add count for each document
                foreach (string docTitle in documentTitles)
                {
                    int count = 0;
                    if (categoryCounts[category].ContainsKey(docTitle))
                    {
                        count = categoryCounts[category][docTitle];
                    }
                    entry[docTitle] = count;
                }

                categoryList.Add(entry);
            }

            // Define properties to display
            var propertyNames = new List<string> { "Category" };
            propertyNames.AddRange(documentTitles);

            // Step 4: Show category DataGrid
            List<Dictionary<string, object>> selectedCategories = CustomGUIs.DataGrid(categoryList, propertyNames, false);
            if (selectedCategories == null || selectedCategories.Count == 0)
                return Result.Cancelled;

            // Step 5: Query for actual elements in selected categories
            var selectedCategoryNames = selectedCategories.Select(c => (string)c["CategoryName"]).ToList();

            var categoryElements = QueryElementsForCategories(documentsToQuery, selectedCategoryNames, token, currentSessionId, uiapp);

            // Step 6: Gather selection items from query results
            List<SelectionItem> selectionItems = new List<SelectionItem>();

            foreach (var categoryEntry in categoryElements)
            {
                foreach (var docEntry in categoryEntry.Value)
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

    private Dictionary<string, Dictionary<string, int>> QueryDocumentsForCategoryCounts(
        List<DocumentEntry> documents, string token, string currentSessionId, UIApplication uiapp)
    {
        // Category -> Document Title -> Count
        var result = new Dictionary<string, Dictionary<string, int>>();

        // Process current session locally
        foreach (var docInfo in documents.Where(d => d.SessionId == currentSessionId))
        {
            try
            {
                var app = uiapp.Application;
                foreach (Document doc in app.Documents)
                {
                    if (doc.IsLinked || doc.Title != docInfo.DocumentTitle) continue;

                    var collector = new FilteredElementCollector(doc);
                    var elements = collector.WhereElementIsNotElementType();
                    var categoryGroups = elements.Where(e => e.Category != null).GroupBy(e => e.Category.Name);

                    foreach (var group in categoryGroups)
                    {
                        string categoryName = group.Key;
                        int count = group.Count();

                        if (!result.ContainsKey(categoryName))
                        {
                            result[categoryName] = new Dictionary<string, int>();
                        }

                        result[categoryName][doc.Title] = count;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warn("SelectByCategoriesInNetwork", $"Failed to query local document: {ex.Message}");
            }
        }

        // Process remote documents via Roslyn - parallel requests through the shared client
        var remoteDocuments = documents.Where(d => d.SessionId != currentSessionId && !string.IsNullOrWhiteSpace(d.DocumentTitle)).ToList();
        if (remoteDocuments.Count > 0)
        {
            var responses = NetworkClient.ExecuteOnDocuments(remoteDocuments, docInfo =>
            {
                // Must find specific document by title - Doc may point to a different active document
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
else
{{
    var collector = new FilteredElementCollector(targetDoc);
    var elements = collector.WhereElementIsNotElementType();
    var categoryGroups = elements.Where(e => e.Category != null).GroupBy(e => e.Category.Name).OrderBy(g => g.Key);
    foreach (var group in categoryGroups)
    {{
        Console.WriteLine(""CATEGORY|"" + group.Key + ""|"" + group.Count());
    }}
}}";
            }, token);

            foreach (var (docInfo, response) in responses)
            {
                if (response != null && response.Success && !string.IsNullOrWhiteSpace(response.Output))
                {
                    ParseCategoryCountResponse(response.Output, docInfo, result);
                }
                else
                {
                    Log.Warn("SelectByCategoriesInNetwork",
                        $"Query returned no data for {docInfo.DocumentTitle}: Success={response?.Success}, Error={response?.Error}");
                }
            }
        }

        return result;
    }

    private void ParseCategoryCountResponse(string output, DocumentEntry docInfo,
        Dictionary<string, Dictionary<string, int>> result)
    {
        foreach (var line in output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (line.StartsWith("CATEGORY|"))
            {
                var parts = line.Substring(9).Split('|');
                if (parts.Length >= 2)
                {
                    string categoryName = parts[0];
                    int count = int.Parse(parts[1]);

                    if (!result.ContainsKey(categoryName))
                    {
                        result[categoryName] = new Dictionary<string, int>();
                    }

                    result[categoryName][docInfo.DocumentTitle] = count;
                }
            }
        }
    }

    private Dictionary<string, Dictionary<string, List<ElementInfo>>> QueryElementsForCategories(
        List<DocumentEntry> documents, List<string> categoryNames, string token, string currentSessionId, UIApplication uiapp)
    {
        // Category -> Document Title -> List of ElementInfo
        var result = new Dictionary<string, Dictionary<string, List<ElementInfo>>>();

        foreach (var categoryName in categoryNames)
        {
            result[categoryName] = new Dictionary<string, List<ElementInfo>>();
        }

        var categoryNamesSet = new HashSet<string>(categoryNames);

        // Process current session locally
        foreach (var docInfo in documents.Where(d => d.SessionId == currentSessionId))
        {
            try
            {
                var app = uiapp.Application;
                foreach (Document doc in app.Documents)
                {
                    if (doc.IsLinked || doc.Title != docInfo.DocumentTitle) continue;

                    var collector = new FilteredElementCollector(doc);
                    var elements = collector.WhereElementIsNotElementType()
                        .Where(e => e.Category != null && categoryNamesSet.Contains(e.Category.Name));

                    foreach (var elem in elements)
                    {
                        string categoryName = elem.Category.Name;

                        if (!result[categoryName].ContainsKey(doc.Title))
                        {
                            result[categoryName][doc.Title] = new List<ElementInfo>();
                        }

                        result[categoryName][doc.Title].Add(new ElementInfo
                        {
                            UniqueId = elem.UniqueId,
#if REVIT2024 || REVIT2025 || REVIT2026
                            ElementIdValue = (int)elem.Id.Value,
#else
                            ElementIdValue = elem.Id.IntegerValue,
#endif
                            DocumentPath = doc.PathName ?? doc.Title,
                            SessionId = docInfo.SessionId
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warn("SelectByCategoriesInNetwork", $"Failed to query elements for local document: {ex.Message}");
            }
        }

        // Process remote documents via Roslyn - parallel requests through the shared client
        var remoteDocuments = documents.Where(d => d.SessionId != currentSessionId && !string.IsNullOrWhiteSpace(d.DocumentTitle)).ToList();
        if (remoteDocuments.Count > 0)
        {
            var categoriesArray = "{ \"" + string.Join("\", \"", categoryNames.Select(c => c.Replace("\"", "\\\""))) + "\" }";

            var responses = NetworkClient.ExecuteOnDocuments(remoteDocuments, docInfo =>
            {
                var escapedTitle = NetworkClient.EscapeForScript(docInfo.DocumentTitle);

                // Must find specific document by title - Doc may point to a different active document
                // Use version-agnostic ElementId access (IntegerValue works on all versions)
                return $@"var docTitle = ""{escapedTitle}"";
var categories = new string[] {categoriesArray};
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
    var collector = new FilteredElementCollector(targetDoc);
    var elements = collector.WhereElementIsNotElementType().Where(e => e.Category != null && categories.Contains(e.Category.Name));
    var categoryGroups = elements.GroupBy(e => e.Category.Name);
    foreach (var group in categoryGroups)
    {{
        Console.WriteLine(""CATEGORY|"" + group.Key);
        foreach (var elem in group)
        {{
            Console.WriteLine(""ELEMENT|"" + elem.UniqueId + ""|"" + elem.Id.IntegerValue);
        }}
    }}
}}";
            }, token);

            foreach (var (docInfo, response) in responses)
            {
                if (response != null && response.Success && !string.IsNullOrWhiteSpace(response.Output))
                {
                    ParseElementsResponse(response.Output, docInfo, result);
                }
                else
                {
                    Log.Warn("SelectByCategoriesInNetwork",
                        $"Element query failed for {docInfo.DocumentTitle}: Success={response?.Success}, Error={response?.Error}");
                }
            }
        }

        return result;
    }

    private void ParseElementsResponse(string output, DocumentEntry docInfo,
        Dictionary<string, Dictionary<string, List<ElementInfo>>> result)
    {
        string currentCategory = null;

        foreach (var line in output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (line.StartsWith("CATEGORY|"))
            {
                currentCategory = line.Substring(9);

                if (result.ContainsKey(currentCategory) && !result[currentCategory].ContainsKey(docInfo.DocumentTitle))
                {
                    result[currentCategory][docInfo.DocumentTitle] = new List<ElementInfo>();
                }
            }
            else if (line.StartsWith("ELEMENT|") && currentCategory != null)
            {
                var parts = line.Substring(8).Split('|');
                if (parts.Length >= 2 && result.ContainsKey(currentCategory))
                {
                    if (!result[currentCategory].ContainsKey(docInfo.DocumentTitle))
                    {
                        result[currentCategory][docInfo.DocumentTitle] = new List<ElementInfo>();
                    }

                    var elemInfo = new ElementInfo
                    {
                        UniqueId = parts[0],
                        ElementIdValue = int.Parse(parts[1]),
                        DocumentPath = docInfo.DocumentPath ?? docInfo.DocumentTitle,
                        SessionId = docInfo.SessionId
                    };

                    result[currentCategory][docInfo.DocumentTitle].Add(elemInfo);
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

    private class ElementInfo
    {
        public string UniqueId { get; set; }
        public int ElementIdValue { get; set; }
        public string DocumentPath { get; set; }
        public string SessionId { get; set; }
    }
}
