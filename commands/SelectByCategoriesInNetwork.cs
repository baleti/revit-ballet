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
using System.Threading.Tasks;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
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
            // Read network token
            string tokenPath = Path.Combine(PathHelper.RuntimeDirectory, "network", "token");
            if (!File.Exists(tokenPath))
            {
                TaskDialog.Show("Error", "Network token not found. Ensure at least one Revit session is running.");
                return Result.Failed;
            }
            string token = File.ReadAllText(tokenPath).Trim();

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
            var documentsToQuery = selectedDocuments.Select(row => (DocumentInfo)row["_Document"]).ToList();

            // Step 2: Query selected documents for category COUNTS only (fast)
            // Use local API for current session, Roslyn for remote sessions
            string currentSessionId = RevitBallet.RevitBallet.SessionId;
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

            // DIAGNOSTIC: Write comprehensive debug info
            var diagnosticLines = new List<string>();
            diagnosticLines.Add($"=== SelectByCategoriesInNetwork Diagnostic at {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ===");
            diagnosticLines.Add($"Current Session ID: {currentSessionId}");
            diagnosticLines.Add($"Documents to query: {documentsToQuery.Count}");
            foreach (var docInfo in documentsToQuery)
            {
                diagnosticLines.Add($"  - SessionId: {docInfo.SessionId}, Doc: {docInfo.DocumentTitle}, Port: {docInfo.Port}");
            }
            diagnosticLines.Add($"Selected categories: {string.Join(", ", selectedCategoryNames)}");
            diagnosticLines.Add("");

            var categoryElements = QueryElementsForCategories(documentsToQuery, selectedCategoryNames, token, currentSessionId, uiapp, diagnosticLines);

            // DIAGNOSTIC: Log what QueryElementsForCategories returned
            diagnosticLines.Add($"QueryElementsForCategories returned {categoryElements.Count} categories");
            foreach (var cat in categoryElements)
            {
                diagnosticLines.Add($"  Category '{cat.Key}': {cat.Value.Count} documents");
                foreach (var doc in cat.Value)
                {
                    diagnosticLines.Add($"    Document '{doc.Key}': {doc.Value.Count} elements");
                    if (doc.Value.Count > 0)
                    {
                        var firstElem = doc.Value[0];
                        diagnosticLines.Add($"      First element: SessionId={firstElem.SessionId}, UniqueId={firstElem.UniqueId}");
                    }
                }
            }
            diagnosticLines.Add("");

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

            // DEBUG: Log selection items before saving
            diagnosticLines.Add($"Gathered {selectionItems.Count} selection items");
            var bySession = selectionItems.GroupBy(s => s.SessionId);
            foreach (var sessGroup in bySession)
            {
                diagnosticLines.Add($"  Session {sessGroup.Key}: {sessGroup.Count()} items");
                var byDoc = sessGroup.GroupBy(s => s.DocumentTitle);
                foreach (var docGroup in byDoc)
                {
                    diagnosticLines.Add($"    Document '{docGroup.Key}': {docGroup.Count()} items");
                }
            }
            diagnosticLines.Add("");

            System.Diagnostics.Debug.WriteLine($"SelectByCategoriesInNetwork: Collected {selectionItems.Count} selection items");
            if (selectionItems.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine($"First item: Doc={selectionItems[0].DocumentTitle}, SessionId={selectionItems[0].SessionId}, UniqueId={selectionItems[0].UniqueId}");
            }

            // Load existing selection and merge
            var existingSelection = SelectionStorage.LoadSelection();
            var existingUniqueIds = new HashSet<string>(existingSelection.Select(s => $"{s.DocumentTitle}|{s.UniqueId}"));

            System.Diagnostics.Debug.WriteLine($"Existing selection count: {existingSelection.Count}");

            // Add new items that don't already exist
            foreach (var item in selectionItems)
            {
                string key = $"{item.DocumentTitle}|{item.UniqueId}";
                if (!existingUniqueIds.Contains(key))
                {
                    existingSelection.Add(item);
                }
            }

            System.Diagnostics.Debug.WriteLine($"Total selection after merge: {existingSelection.Count}");

            diagnosticLines.Add($"After merge: {existingSelection.Count} total items");
            diagnosticLines.Add("");

            // Save selection
            SelectionStorage.SaveSelection(existingSelection);

            System.Diagnostics.Debug.WriteLine($"Selection saved to storage");

            diagnosticLines.Add($"Selection saved to storage");
            diagnosticLines.Add("=== END DIAGNOSTIC ===");

            // Write final diagnostic
            try
            {
                string diagnosticPath = Path.Combine(PathHelper.RuntimeDirectory, "diagnostics", $"SelectByCategoriesInNetwork-Query-{DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");
                Directory.CreateDirectory(Path.GetDirectoryName(diagnosticPath));
                File.WriteAllLines(diagnosticPath, diagnosticLines);
            }
            catch { }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Failed to query network sessions: {ex.Message}");
            return Result.Failed;
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

                // Filter out stale documents (no heartbeat for > 2 minutes)
                if ((DateTime.Now - doc.LastHeartbeat).TotalSeconds < 120 &&
                    !string.IsNullOrWhiteSpace(doc.DocumentTitle))
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

    private Dictionary<string, Dictionary<string, int>> QueryDocumentsForCategoryCounts(
        List<DocumentInfo> documents, string token, string currentSessionId, UIApplication uiapp)
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
                System.Diagnostics.Debug.WriteLine($"Failed to query local document: {ex.Message}");
            }
        }

        // Process remote documents via Roslyn - PARALLEL requests
        var remoteDocuments = documents.Where(d => d.SessionId != currentSessionId && !string.IsNullOrWhiteSpace(d.DocumentTitle)).ToList();
        if (remoteDocuments.Count > 0)
        {
            using (var handler = new HttpClientHandler())
            {
                handler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;

                using (var client = new HttpClient(handler))
                {
                    client.Timeout = TimeSpan.FromSeconds(120);

                    // Build all requests and send them in parallel
                    var tasks = new List<Task<(DocumentInfo DocInfo, RoslynResponse Response)>>();

                    foreach (var docInfo in remoteDocuments)
                    {
                        // Must find specific document by title - Doc may point to a different active document
                        var escapedTitle = docInfo.DocumentTitle.Replace("\\", "\\\\").Replace("\"", "\\\"");
                        var query = $@"var docTitle = ""{escapedTitle}"";
Console.WriteLine(""DEBUG|Looking for document: "" + docTitle);
Console.WriteLine(""DEBUG|Available documents: "" + UIApp.Application.Documents.Size);
Document targetDoc = null;
foreach (Document d in UIApp.Application.Documents)
{{
    Console.WriteLine(""DEBUG|  Found: "" + d.Title + "" (IsLinked="" + d.IsLinked + "")"");
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
    Console.WriteLine(""DEBUG|Found "" + categoryGroups.Count() + "" categories"");
    foreach (var group in categoryGroups)
    {{
        Console.WriteLine(""CATEGORY|"" + group.Key + ""|"" + group.Count());
    }}
}}";

                        var task = SendRoslynRequestAsync(client, docInfo, query, token);
                        tasks.Add(task);
                    }

                    // Wait for all requests to complete in parallel
                    try
                    {
                        Task.WhenAll(tasks).Wait();
                    }
                    catch (AggregateException)
                    {
                        // Individual task exceptions are handled below
                    }

                    // Process all responses
                    foreach (var task in tasks)
                    {
                        try
                        {
                            if (task.Status == TaskStatus.RanToCompletion)
                            {
                                var (docInfo, jsonResponse) = task.Result;
                                if (jsonResponse != null && jsonResponse.Success && !string.IsNullOrWhiteSpace(jsonResponse.Output))
                                {
                                    ParseCategoryCountResponse(jsonResponse.Output, docInfo, result);
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"Query returned no data for {docInfo.DocumentTitle}: Success={jsonResponse?.Success}, Error={jsonResponse?.Error}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to process response: {ex.Message}");
                        }
                    }
                }
            }
        }

        return result;
    }

    private async Task<(DocumentInfo DocInfo, RoslynResponse Response)> SendRoslynRequestAsync(
        HttpClient client, DocumentInfo docInfo, string query, string token)
    {
        try
        {
            var content = new StringContent(query, Encoding.UTF8, "text/plain");
            var request = new HttpRequestMessage(HttpMethod.Post, $"https://127.0.0.1:{docInfo.Port}/roslyn");
            request.Headers.Add("X-Auth-Token", token);
            request.Content = content;

            var response = await client.SendAsync(request);
            var responseText = await response.Content.ReadAsStringAsync();

            var jsonResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<RoslynResponse>(responseText);
            return (docInfo, jsonResponse);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to query document {docInfo.SessionId} ({docInfo.DocumentTitle}): {ex.Message}");
            return (docInfo, null);
        }
    }

    private void ParseCategoryCountResponse(string output, DocumentInfo docInfo,
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
        List<DocumentInfo> documents, List<string> categoryNames, string token, string currentSessionId, UIApplication uiapp, List<string> diagnosticLines)
    {
        System.Diagnostics.Debug.WriteLine($"QueryElementsForCategories: documents={documents.Count}, categories={categoryNames.Count}, currentSessionId={currentSessionId}");
        diagnosticLines.Add($"--- QueryElementsForCategories Start ---");

        // Category -> Document Title -> List of ElementInfo
        var result = new Dictionary<string, Dictionary<string, List<ElementInfo>>>();

        foreach (var categoryName in categoryNames)
        {
            result[categoryName] = new Dictionary<string, List<ElementInfo>>();
        }

        var categoryNamesSet = new HashSet<string>(categoryNames);

        // Process current session locally
        var currentDocs = documents.Where(d => d.SessionId == currentSessionId).ToList();
        System.Diagnostics.Debug.WriteLine($"Current session document matches: {currentDocs.Count}");
        diagnosticLines.Add($"Local documents (SessionId={currentSessionId}): {currentDocs.Count}");

        foreach (var docInfo in currentDocs)
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
                System.Diagnostics.Debug.WriteLine($"Failed to query elements for local document: {ex.Message}");
            }
        }

        // Process remote documents via Roslyn - PARALLEL requests
        var remoteDocuments = documents.Where(d => d.SessionId != currentSessionId && !string.IsNullOrWhiteSpace(d.DocumentTitle)).ToList();
        System.Diagnostics.Debug.WriteLine($"Remote document matches: {remoteDocuments.Count}");
        diagnosticLines.Add($"Remote documents: {remoteDocuments.Count}");
        foreach (var rd in remoteDocuments)
        {
            System.Diagnostics.Debug.WriteLine($"  Remote document: {rd.SessionId}, Doc: {rd.DocumentTitle}, Port: {rd.Port}");
            diagnosticLines.Add($"  - SessionId={rd.SessionId}, Doc={rd.DocumentTitle}, Port={rd.Port}");
        }

        if (remoteDocuments.Count > 0)
        {
            using (var handler = new HttpClientHandler())
            {
                handler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;

                using (var client = new HttpClient(handler))
                {
                    client.Timeout = TimeSpan.FromSeconds(120);

                    // Build all requests and send them in parallel
                    var categoriesArray = "{ \"" + string.Join("\", \"", categoryNames.Select(c => c.Replace("\"", "\\\""))) + "\" }";
                    var tasks = new List<Task<(DocumentInfo DocInfo, RoslynResponse Response)>>();

                    diagnosticLines.Add($"  Sending {remoteDocuments.Count} requests in parallel...");

                    foreach (var docInfo in remoteDocuments)
                    {
                        var escapedTitle = docInfo.DocumentTitle.Replace("\\", "\\\\").Replace("\"", "\\\"");

                        // Must find specific document by title - Doc may point to a different active document
                        // Use version-agnostic ElementId access (IntegerValue works on all versions)
                        var query = $@"var docTitle = ""{escapedTitle}"";
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

                        var task = SendRoslynRequestAsync(client, docInfo, query, token);
                        tasks.Add(task);
                    }

                    // Wait for all requests to complete in parallel
                    try
                    {
                        Task.WhenAll(tasks).Wait();
                    }
                    catch (AggregateException)
                    {
                        // Individual task exceptions are handled below
                    }

                    diagnosticLines.Add($"  All {tasks.Count} requests completed");

                    // Process all responses
                    foreach (var task in tasks)
                    {
                        try
                        {
                            if (task.Status == TaskStatus.RanToCompletion)
                            {
                                var (docInfo, jsonResponse) = task.Result;

                                diagnosticLines.Add($"  Response from {docInfo.DocumentTitle}: Success={jsonResponse?.Success}, HasOutput={!string.IsNullOrWhiteSpace(jsonResponse?.Output)}");
                                if (jsonResponse != null && !string.IsNullOrWhiteSpace(jsonResponse.Output))
                                {
                                    diagnosticLines.Add($"  Output preview: {jsonResponse.Output.Substring(0, Math.Min(500, jsonResponse.Output.Length))}");
                                }
                                if (jsonResponse != null && !string.IsNullOrWhiteSpace(jsonResponse.Error))
                                {
                                    diagnosticLines.Add($"  Error: {jsonResponse.Error}");
                                }

                                System.Diagnostics.Debug.WriteLine($"  Remote query response: Success={jsonResponse?.Success}, HasOutput={!string.IsNullOrWhiteSpace(jsonResponse?.Output)}");

                                if (jsonResponse != null && jsonResponse.Success && !string.IsNullOrWhiteSpace(jsonResponse.Output))
                                {
                                    System.Diagnostics.Debug.WriteLine($"  Parsing elements response for document {docInfo.SessionId}");
                                    ParseElementsResponse(jsonResponse.Output, docInfo, result);
                                    System.Diagnostics.Debug.WriteLine($"  After parsing, result has {result.Count} categories");
                                    diagnosticLines.Add($"  After parsing: {result.Values.Sum(d => d.Values.Sum(l => l.Count))} elements total");
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"  Remote query failed or no output: Error={jsonResponse?.Error}");
                                }
                            }
                            else if (task.IsFaulted)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to query elements: {task.Exception?.Message}");
                                diagnosticLines.Add($"  EXCEPTION: {task.Exception?.Message}");
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to process response: {ex.Message}");
                            diagnosticLines.Add($"  EXCEPTION: {ex.Message}");
                        }
                    }
                }
            }
        }

        diagnosticLines.Add($"--- QueryElementsForCategories End ---");
        System.Diagnostics.Debug.WriteLine($"QueryElementsForCategories completed: {result.Count} categories");
        foreach (var cat in result)
        {
            System.Diagnostics.Debug.WriteLine($"  Category '{cat.Key}': {cat.Value.Count} documents");
            foreach (var doc in cat.Value)
            {
                System.Diagnostics.Debug.WriteLine($"    Doc '{doc.Key}': {doc.Value.Count} elements");
            }
        }

        return result;
    }

    private void ParseElementsResponse(string output, DocumentInfo docInfo,
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

    private class DocumentInfo
    {
        public string SessionId { get; set; }
        public string Port { get; set; }
        public string Hostname { get; set; }
        public int ProcessId { get; set; }
        public DateTime RegisteredAt { get; set; }
        public DateTime LastHeartbeat { get; set; }
        public string DocumentTitle { get; set; }
        public string DocumentPath { get; set; }
    }

    private class ElementInfo
    {
        public string UniqueId { get; set; }
        public int ElementIdValue { get; set; }
        public string DocumentPath { get; set; }
        public string SessionId { get; set; }
    }

    private class RoslynResponse
    {
        public bool Success { get; set; }
        public string Output { get; set; }
        public string Error { get; set; }
    }
}
