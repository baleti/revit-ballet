using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectByFamilyTypesInNetwork : IExternalCommand
{
    /// <summary>
    /// Marks this command as usable outside Revit context via network.
    /// </summary>
    public static bool IsNetworkCommand => true;

    // Diagnostic tracking
    private List<string> _diagnostics = new List<string>();
    private Stopwatch _overallStopwatch = new Stopwatch();
    private Dictionary<string, DocumentTiming> _documentTimings = new Dictionary<string, DocumentTiming>();

    private class DocumentTiming
    {
        public string DocumentTitle { get; set; }
        public string SessionId { get; set; }
        public bool IsLocal { get; set; }
        public double CountQueryMs { get; set; }
        public double ElementsQueryMs { get; set; }
        public int FamilyTypesFound { get; set; }
        public int ElementsFound { get; set; }
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        _overallStopwatch.Start();
        _diagnostics.Clear();
        _documentTimings.Clear();

        _diagnostics.Add($"=== SelectByFamilyTypesInNetwork Timing Diagnostics ===");
        _diagnostics.Add($"Started at: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
        _diagnostics.Add("");

        try
        {
            var stepWatch = Stopwatch.StartNew();

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

            stepWatch.Stop();
            _diagnostics.Add($"[STEP 0] Initialize and parse documents: {stepWatch.ElapsedMilliseconds}ms");
            _diagnostics.Add($"         Found {documents.Count} registered documents");
            _diagnostics.Add("");

            // Step 1: Show documents DataGrid - let user select which documents to query
            stepWatch.Restart();
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

            stepWatch.Stop();
            _diagnostics.Add($"[STEP 1] Document selection UI: {stepWatch.ElapsedMilliseconds}ms (includes user interaction)");

            if (selectedDocuments == null || selectedDocuments.Count == 0)
            {
                WriteDiagnostics("Cancelled at document selection");
                return Result.Cancelled;
            }

            // Extract selected document objects
            var documentsToQuery = selectedDocuments.Select(row => (DocumentInfo)row["_Document"]).ToList();
            _diagnostics.Add($"         Selected {documentsToQuery.Count} documents to query");
            _diagnostics.Add("");

            // Initialize timing entries for each document
            string currentSessionId = RevitBallet.RevitBallet.SessionId;
            foreach (var doc in documentsToQuery)
            {
                _documentTimings[doc.DocumentTitle] = new DocumentTiming
                {
                    DocumentTitle = doc.DocumentTitle,
                    SessionId = doc.SessionId,
                    IsLocal = doc.SessionId == currentSessionId
                };
            }

            // Step 2: Query selected documents for family type COUNTS only (fast)
            stepWatch.Restart();
            var familyTypeCounts = QueryDocumentsForFamilyTypeCounts(documentsToQuery, token, currentSessionId, uiapp);
            stepWatch.Stop();

            _diagnostics.Add($"[STEP 2] Query family type counts: {stepWatch.ElapsedMilliseconds}ms TOTAL");
            _diagnostics.Add($"         Found {familyTypeCounts.Count} unique family types");
            _diagnostics.Add("");
            _diagnostics.Add("         Per-document breakdown:");
            foreach (var timing in _documentTimings.Values.OrderBy(t => t.DocumentTitle))
            {
                string locality = timing.IsLocal ? "LOCAL" : "REMOTE";
                _diagnostics.Add($"           [{locality}] {timing.DocumentTitle}: {timing.CountQueryMs:F1}ms, {timing.FamilyTypesFound} types");
            }
            _diagnostics.Add("");

            if (familyTypeCounts.Count == 0)
            {
                TaskDialog.Show("No Family Types", "No family types found in selected documents.");
                WriteDiagnostics("No family types found");
                return Result.Cancelled;
            }

            // Step 3: Build DataGrid with family types as rows and documents as columns
            stepWatch.Restart();
            var documentTitles = documentsToQuery.Select(d => d.DocumentTitle).Distinct().ToList();

            var familyTypeList = new List<Dictionary<string, object>>();

            foreach (var familyType in familyTypeCounts.Keys.OrderBy(ft => ft.Category).ThenBy(ft => ft.Family).ThenBy(ft => ft.TypeName))
            {
                var entry = new Dictionary<string, object>
                {
                    { "Category", familyType.Category },
                    { "Family", familyType.Family },
                    { "Type Name", familyType.TypeName },
                    { "_FamilyTypeKey", familyType }
                };

                // Add count for each document
                foreach (string docTitle in documentTitles)
                {
                    int count = 0;
                    if (familyTypeCounts[familyType].ContainsKey(docTitle))
                    {
                        count = familyTypeCounts[familyType][docTitle];
                    }
                    entry[docTitle] = count;
                }

                familyTypeList.Add(entry);
            }

            // Define properties to display
            var propertyNames = new List<string> { "Category", "Family", "Type Name" };
            propertyNames.AddRange(documentTitles);

            stepWatch.Stop();
            _diagnostics.Add($"[STEP 3] Build family type grid data: {stepWatch.ElapsedMilliseconds}ms");
            _diagnostics.Add($"         {familyTypeList.Count} rows prepared");
            _diagnostics.Add("");

            // Step 4: Show family type DataGrid
            stepWatch.Restart();
            List<Dictionary<string, object>> selectedFamilyTypes = CustomGUIs.DataGrid(familyTypeList, propertyNames, false);
            stepWatch.Stop();
            _diagnostics.Add($"[STEP 4] Family type selection UI: {stepWatch.ElapsedMilliseconds}ms (includes user interaction)");

            if (selectedFamilyTypes == null || selectedFamilyTypes.Count == 0)
            {
                WriteDiagnostics("Cancelled at family type selection");
                return Result.Cancelled;
            }

            _diagnostics.Add($"         Selected {selectedFamilyTypes.Count} family types");
            _diagnostics.Add("");

            // Step 5: Query for actual elements of selected family types
            stepWatch.Restart();
            var selectedFamilyTypeKeys = selectedFamilyTypes
                .Where(ft => ft.ContainsKey("_FamilyTypeKey"))
                .Select(ft => (FamilyTypeKey)ft["_FamilyTypeKey"])
                .ToList();

            // Reset element counts for timing
            foreach (var timing in _documentTimings.Values)
            {
                timing.ElementsQueryMs = 0;
                timing.ElementsFound = 0;
            }

            var familyTypeElements = QueryElementsForFamilyTypes(documentsToQuery, selectedFamilyTypeKeys, token, currentSessionId, uiapp);
            stepWatch.Stop();

            int totalElements = familyTypeElements.Values.Sum(d => d.Values.Sum(l => l.Count));
            _diagnostics.Add($"[STEP 5] Query elements for selected types: {stepWatch.ElapsedMilliseconds}ms TOTAL");
            _diagnostics.Add($"         Found {totalElements} total elements");
            _diagnostics.Add("");
            _diagnostics.Add("         Per-document breakdown:");
            foreach (var timing in _documentTimings.Values.OrderBy(t => t.DocumentTitle))
            {
                string locality = timing.IsLocal ? "LOCAL" : "REMOTE";
                _diagnostics.Add($"           [{locality}] {timing.DocumentTitle}: {timing.ElementsQueryMs:F1}ms, {timing.ElementsFound} elements");
            }
            _diagnostics.Add("");

            // Step 6: Gather selection items from query results
            stepWatch.Restart();
            List<SelectionItem> selectionItems = new List<SelectionItem>();

            foreach (var familyTypeEntry in familyTypeElements)
            {
                foreach (var docEntry in familyTypeEntry.Value)
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

            int newItemsAdded = 0;
            // Add new items that don't already exist
            foreach (var item in selectionItems)
            {
                string key = $"{item.DocumentTitle}|{item.UniqueId}";
                if (!existingUniqueIds.Contains(key))
                {
                    existingSelection.Add(item);
                    newItemsAdded++;
                }
            }

            // Save selection
            SelectionStorage.SaveSelection(existingSelection);
            stepWatch.Stop();

            _diagnostics.Add($"[STEP 6] Merge and save selection: {stepWatch.ElapsedMilliseconds}ms");
            _diagnostics.Add($"         {newItemsAdded} new items added to selection");
            _diagnostics.Add($"         {existingSelection.Count} total items in selection storage");
            _diagnostics.Add("");

            _overallStopwatch.Stop();
            _diagnostics.Add($"=== TOTAL EXECUTION TIME: {_overallStopwatch.ElapsedMilliseconds}ms ===");

            // Summary statistics
            _diagnostics.Add("");
            _diagnostics.Add("=== PERFORMANCE SUMMARY ===");
            var localDocs = _documentTimings.Values.Where(t => t.IsLocal).ToList();
            var remoteDocs = _documentTimings.Values.Where(t => !t.IsLocal).ToList();

            if (localDocs.Any())
            {
                _diagnostics.Add($"Local documents ({localDocs.Count}):");
                _diagnostics.Add($"  Count query: avg={localDocs.Average(t => t.CountQueryMs):F1}ms, max={localDocs.Max(t => t.CountQueryMs):F1}ms");
                _diagnostics.Add($"  Elements query: avg={localDocs.Average(t => t.ElementsQueryMs):F1}ms, max={localDocs.Max(t => t.ElementsQueryMs):F1}ms");
            }

            if (remoteDocs.Any())
            {
                _diagnostics.Add($"Remote documents ({remoteDocs.Count}):");
                _diagnostics.Add($"  Count query: avg={remoteDocs.Average(t => t.CountQueryMs):F1}ms, max={remoteDocs.Max(t => t.CountQueryMs):F1}ms, min={remoteDocs.Min(t => t.CountQueryMs):F1}ms");
                _diagnostics.Add($"  Elements query: avg={remoteDocs.Average(t => t.ElementsQueryMs):F1}ms, max={remoteDocs.Max(t => t.ElementsQueryMs):F1}ms, min={remoteDocs.Min(t => t.ElementsQueryMs):F1}ms");

                // Identify slowest document
                var slowestCount = remoteDocs.OrderByDescending(t => t.CountQueryMs).First();
                var slowestElements = remoteDocs.OrderByDescending(t => t.ElementsQueryMs).First();
                _diagnostics.Add($"  Slowest count query: {slowestCount.DocumentTitle} ({slowestCount.CountQueryMs:F1}ms)");
                _diagnostics.Add($"  Slowest elements query: {slowestElements.DocumentTitle} ({slowestElements.ElementsQueryMs:F1}ms)");
            }

            WriteDiagnostics("Completed successfully");
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            _overallStopwatch.Stop();
            _diagnostics.Add($"[ERROR] Exception after {_overallStopwatch.ElapsedMilliseconds}ms: {ex.Message}");
            _diagnostics.Add(ex.StackTrace);
            WriteDiagnostics("Failed with exception");

            TaskDialog.Show("Error", $"Failed to query network sessions: {ex.Message}");
            return Result.Failed;
        }
    }

    private void WriteDiagnostics(string status)
    {
        // Diagnostic file writing disabled
        // try
        // {
        //     _diagnostics.Add("");
        //     _diagnostics.Add($"Status: {status}");
        //     _diagnostics.Add($"Completed at: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
        //
        //     string diagnosticPath = Path.Combine(
        //         PathHelper.RuntimeDirectory,
        //         "diagnostics",
        //         $"SelectByFamilyTypesInNetwork-{DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");
        //
        //     Directory.CreateDirectory(Path.GetDirectoryName(diagnosticPath));
        //     File.WriteAllLines(diagnosticPath, _diagnostics);
        // }
        // catch { }
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

    private Dictionary<FamilyTypeKey, Dictionary<string, int>> QueryDocumentsForFamilyTypeCounts(
        List<DocumentInfo> documents, string token, string currentSessionId, UIApplication uiapp)
    {
        // FamilyTypeKey -> Document Title -> Instance Count
        var result = new Dictionary<FamilyTypeKey, Dictionary<string, int>>();

        // Process current session locally
        foreach (var docInfo in documents.Where(d => d.SessionId == currentSessionId))
        {
            var localWatch = Stopwatch.StartNew();
            int typesFound = 0;

            try
            {
                var app = uiapp.Application;
                foreach (Document doc in app.Documents)
                {
                    if (doc.IsLinked || doc.Title != docInfo.DocumentTitle) continue;

                    var elementTypes = new FilteredElementCollector(doc)
                        .OfClass(typeof(ElementType))
                        .Cast<ElementType>()
                        .ToList();

                    // Count instances by type
                    Dictionary<ElementId, int> typeInstanceCounts = new Dictionary<ElementId, int>();
                    var allInstances = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .Where(x => x.GetTypeId() != null && x.GetTypeId() != ElementId.InvalidElementId)
                        .ToList();

                    foreach (var instance in allInstances)
                    {
                        ElementId typeId = instance.GetTypeId();
                        if (!typeInstanceCounts.ContainsKey(typeId))
                            typeInstanceCounts[typeId] = 0;
                        typeInstanceCounts[typeId]++;
                    }

                    foreach (var elementType in elementTypes)
                    {
                        string typeName = elementType.Name;
                        string familyName = "";
                        string categoryName = "";

                        if (elementType is FamilySymbol fs)
                        {
                            familyName = fs.Family.Name;
                            categoryName = fs.Category != null ? fs.Category.Name : "N/A";
                            if (categoryName.Contains("Import Symbol") || familyName.Contains("Import Symbol"))
                                continue;
                        }
                        else
                        {
                            Parameter familyParam = elementType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                            familyName = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
                                ? familyParam.AsString()
                                : "System Type";
                            categoryName = elementType.Category != null ? elementType.Category.Name : "N/A";
                            if (categoryName.Contains("Import Symbol") || familyName.Contains("Import Symbol"))
                                continue;
                        }

                        int count = typeInstanceCounts.ContainsKey(elementType.Id) ? typeInstanceCounts[elementType.Id] : 0;

                        var key = new FamilyTypeKey { Category = categoryName, Family = familyName, TypeName = typeName };
                        if (!result.ContainsKey(key))
                            result[key] = new Dictionary<string, int>();

                        if (!result[key].ContainsKey(doc.Title))
                            result[key][doc.Title] = 0;
                        result[key][doc.Title] += count;
                        typesFound++;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to query local document: {ex.Message}");
            }

            localWatch.Stop();
            if (_documentTimings.ContainsKey(docInfo.DocumentTitle))
            {
                _documentTimings[docInfo.DocumentTitle].CountQueryMs = localWatch.Elapsed.TotalMilliseconds;
                _documentTimings[docInfo.DocumentTitle].FamilyTypesFound = typesFound;
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

                    var tasks = new List<Task<(DocumentInfo DocInfo, RoslynResponse Response, double ElapsedMs)>>();

                    foreach (var docInfo in remoteDocuments)
                    {
                        var requestBody = Newtonsoft.Json.JsonConvert.SerializeObject(new Dictionary<string, string>
                        {
                            { "documentTitle", docInfo.DocumentTitle }
                        });

                        var task = SendQueryRequestWithTimingAsync(client, docInfo, "/query/familytypes/counts", requestBody, token);
                        tasks.Add(task);
                    }

                    try { Task.WhenAll(tasks).Wait(); }
                    catch (AggregateException) { }

                    foreach (var task in tasks)
                    {
                        try
                        {
                            if (task.Status == TaskStatus.RanToCompletion)
                            {
                                var (docInfo, jsonResponse, elapsedMs) = task.Result;

                                // Record timing
                                if (_documentTimings.ContainsKey(docInfo.DocumentTitle))
                                {
                                    _documentTimings[docInfo.DocumentTitle].CountQueryMs = elapsedMs;
                                }

                                if (jsonResponse != null && jsonResponse.Success && !string.IsNullOrWhiteSpace(jsonResponse.Output))
                                {
                                    int typesFound = ParseFamilyTypeCountResponse(jsonResponse.Output, docInfo, result);
                                    if (_documentTimings.ContainsKey(docInfo.DocumentTitle))
                                    {
                                        _documentTimings[docInfo.DocumentTitle].FamilyTypesFound = typesFound;
                                    }
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

    private async Task<(DocumentInfo DocInfo, RoslynResponse Response, double ElapsedMs)> SendQueryRequestWithTimingAsync(
        HttpClient client, DocumentInfo docInfo, string endpoint, string requestBody, string token)
    {
        var stopwatch = Stopwatch.StartNew();
        try
        {
            var content = new StringContent(requestBody, Encoding.UTF8, "application/json");
            var request = new HttpRequestMessage(HttpMethod.Post, $"https://127.0.0.1:{docInfo.Port}{endpoint}");
            request.Headers.Add("X-Auth-Token", token);
            request.Content = content;

            var response = await client.SendAsync(request);
            var responseText = await response.Content.ReadAsStringAsync();

            stopwatch.Stop();
            var jsonResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<RoslynResponse>(responseText);
            return (docInfo, jsonResponse, stopwatch.Elapsed.TotalMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            System.Diagnostics.Debug.WriteLine($"Failed to query document {docInfo.SessionId} ({docInfo.DocumentTitle}): {ex.Message}");
            return (docInfo, null, stopwatch.Elapsed.TotalMilliseconds);
        }
    }

    private int ParseFamilyTypeCountResponse(string output, DocumentInfo docInfo,
        Dictionary<FamilyTypeKey, Dictionary<string, int>> result)
    {
        int typesFound = 0;
        foreach (var line in output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (line.StartsWith("FAMILYTYPE|"))
            {
                var parts = line.Substring(11).Split('|');
                if (parts.Length >= 4)
                {
                    var key = new FamilyTypeKey
                    {
                        Category = parts[0],
                        Family = parts[1],
                        TypeName = parts[2]
                    };
                    int count = int.Parse(parts[3]);

                    if (!result.ContainsKey(key))
                        result[key] = new Dictionary<string, int>();

                    result[key][docInfo.DocumentTitle] = count;
                    typesFound++;
                }
            }
        }
        return typesFound;
    }

    private Dictionary<FamilyTypeKey, Dictionary<string, List<ElementInfo>>> QueryElementsForFamilyTypes(
        List<DocumentInfo> documents, List<FamilyTypeKey> familyTypeKeys, string token, string currentSessionId, UIApplication uiapp)
    {
        var result = new Dictionary<FamilyTypeKey, Dictionary<string, List<ElementInfo>>>();

        foreach (var key in familyTypeKeys)
            result[key] = new Dictionary<string, List<ElementInfo>>();

        // Process current session locally
        foreach (var docInfo in documents.Where(d => d.SessionId == currentSessionId))
        {
            var localWatch = Stopwatch.StartNew();
            int elementsFound = 0;

            try
            {
                var app = uiapp.Application;
                foreach (Document doc in app.Documents)
                {
                    if (doc.IsLinked || doc.Title != docInfo.DocumentTitle) continue;

                    // Build a set of type IDs matching our family type keys
                    var matchingTypeIds = new HashSet<ElementId>();
                    var typeIdToKey = new Dictionary<ElementId, FamilyTypeKey>();

                    var elementTypes = new FilteredElementCollector(doc)
                        .OfClass(typeof(ElementType))
                        .Cast<ElementType>()
                        .ToList();

                    foreach (var elementType in elementTypes)
                    {
                        string typeName = elementType.Name;
                        string familyName = "";
                        string categoryName = "";

                        if (elementType is FamilySymbol fs)
                        {
                            familyName = fs.Family.Name;
                            categoryName = fs.Category != null ? fs.Category.Name : "N/A";
                        }
                        else
                        {
                            Parameter familyParam = elementType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                            familyName = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
                                ? familyParam.AsString()
                                : "System Type";
                            categoryName = elementType.Category != null ? elementType.Category.Name : "N/A";
                        }

                        var key = new FamilyTypeKey { Category = categoryName, Family = familyName, TypeName = typeName };
                        if (familyTypeKeys.Any(k => k.Category == key.Category && k.Family == key.Family && k.TypeName == key.TypeName))
                        {
                            matchingTypeIds.Add(elementType.Id);
                            typeIdToKey[elementType.Id] = key;
                        }
                    }

                    // Collect instances of matching types
                    var instances = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .Where(x => x.GetTypeId() != null && matchingTypeIds.Contains(x.GetTypeId()))
                        .ToList();

                    foreach (var elem in instances)
                    {
                        var key = typeIdToKey[elem.GetTypeId()];
                        var matchingKey = familyTypeKeys.First(k => k.Category == key.Category && k.Family == key.Family && k.TypeName == key.TypeName);

                        if (!result[matchingKey].ContainsKey(doc.Title))
                            result[matchingKey][doc.Title] = new List<ElementInfo>();

                        result[matchingKey][doc.Title].Add(new ElementInfo
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
                        elementsFound++;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to query elements for local document: {ex.Message}");
            }

            localWatch.Stop();
            if (_documentTimings.ContainsKey(docInfo.DocumentTitle))
            {
                _documentTimings[docInfo.DocumentTitle].ElementsQueryMs = localWatch.Elapsed.TotalMilliseconds;
                _documentTimings[docInfo.DocumentTitle].ElementsFound = elementsFound;
            }
        }

        // Process remote documents via pre-compiled endpoint - PARALLEL requests
        var remoteDocuments = documents.Where(d => d.SessionId != currentSessionId && !string.IsNullOrWhiteSpace(d.DocumentTitle)).ToList();
        if (remoteDocuments.Count > 0)
        {
            // Build family type list for JSON request
            var familyTypesList = familyTypeKeys.Select(ft => new Dictionary<string, string>
            {
                { "category", ft.Category },
                { "family", ft.Family },
                { "typeName", ft.TypeName }
            }).ToList();

            using (var handler = new HttpClientHandler())
            {
                handler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;

                using (var client = new HttpClient(handler))
                {
                    client.Timeout = TimeSpan.FromSeconds(120);

                    var tasks = new List<Task<(DocumentInfo DocInfo, RoslynResponse Response, double ElapsedMs)>>();

                    foreach (var docInfo in remoteDocuments)
                    {
                        var requestBody = Newtonsoft.Json.JsonConvert.SerializeObject(new Dictionary<string, object>
                        {
                            { "documentTitle", docInfo.DocumentTitle },
                            { "familyTypes", Newtonsoft.Json.JsonConvert.SerializeObject(familyTypesList) }
                        });

                        var task = SendQueryRequestWithTimingAsync(client, docInfo, "/query/familytypes/elements", requestBody, token);
                        tasks.Add(task);
                    }

                    try { Task.WhenAll(tasks).Wait(); }
                    catch (AggregateException) { }

                    foreach (var task in tasks)
                    {
                        try
                        {
                            if (task.Status == TaskStatus.RanToCompletion)
                            {
                                var (docInfo, jsonResponse, elapsedMs) = task.Result;

                                // Record timing
                                if (_documentTimings.ContainsKey(docInfo.DocumentTitle))
                                {
                                    _documentTimings[docInfo.DocumentTitle].ElementsQueryMs = elapsedMs;
                                }

                                if (jsonResponse != null && jsonResponse.Success && !string.IsNullOrWhiteSpace(jsonResponse.Output))
                                {
                                    int elementsFound = ParseElementsResponse(jsonResponse.Output, docInfo, familyTypeKeys, result);
                                    if (_documentTimings.ContainsKey(docInfo.DocumentTitle))
                                    {
                                        _documentTimings[docInfo.DocumentTitle].ElementsFound = elementsFound;
                                    }
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

    private int ParseElementsResponse(string output, DocumentInfo docInfo, List<FamilyTypeKey> familyTypeKeys,
        Dictionary<FamilyTypeKey, Dictionary<string, List<ElementInfo>>> result)
    {
        int elementsFound = 0;
        foreach (var line in output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (line.StartsWith("ELEMENT|"))
            {
                var parts = line.Substring(8).Split('|');
                if (parts.Length >= 5)
                {
                    var key = new FamilyTypeKey
                    {
                        Category = parts[0],
                        Family = parts[1],
                        TypeName = parts[2]
                    };

                    var matchingKey = familyTypeKeys.FirstOrDefault(k => k.Category == key.Category && k.Family == key.Family && k.TypeName == key.TypeName);
                    if (matchingKey != null)
                    {
                        if (!result[matchingKey].ContainsKey(docInfo.DocumentTitle))
                            result[matchingKey][docInfo.DocumentTitle] = new List<ElementInfo>();

                        result[matchingKey][docInfo.DocumentTitle].Add(new ElementInfo
                        {
                            UniqueId = parts[3],
                            ElementIdValue = int.Parse(parts[4]),
                            DocumentPath = docInfo.DocumentPath ?? docInfo.DocumentTitle,
                            SessionId = docInfo.SessionId
                        });
                        elementsFound++;
                    }
                }
            }
        }
        return elementsFound;
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

    private class FamilyTypeKey
    {
        public string Category { get; set; }
        public string Family { get; set; }
        public string TypeName { get; set; }

        public override bool Equals(object obj)
        {
            if (obj is FamilyTypeKey other)
                return Category == other.Category && Family == other.Family && TypeName == other.TypeName;
            return false;
        }

        public override int GetHashCode()
        {
            return (Category ?? "").GetHashCode() ^ (Family ?? "").GetHashCode() ^ (TypeName ?? "").GetHashCode();
        }
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
