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
public class SelectByFamilyTypesInNetwork : IExternalCommand
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

            // Read documents file
            string documentsPath = Path.Combine(PathHelper.RuntimeDirectory, "documents");
            if (!File.Exists(documentsPath))
            {
                TaskDialog.Show("Error", "Documents file not found.");
                return Result.Failed;
            }

            var documents = ParseDocumentsFile(documentsPath);
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
            var documentsToQuery = selectedDocuments.Select(row => (DocumentInfo)row["_Document"]).ToList();

            // Step 2: Query selected documents for family type COUNTS only (fast)
            string currentSessionId = RevitBallet.RevitBallet.SessionId;
            var familyTypeCounts = QueryDocumentsForFamilyTypeCounts(documentsToQuery, token, currentSessionId, uiapp);

            if (familyTypeCounts.Count == 0)
            {
                TaskDialog.Show("No Family Types", "No family types found in selected documents.");
                return Result.Cancelled;
            }

            // Step 3: Build DataGrid with family types as rows and documents as columns
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

            // Step 4: Show family type DataGrid
            List<Dictionary<string, object>> selectedFamilyTypes = CustomGUIs.DataGrid(familyTypeList, propertyNames, false);
            if (selectedFamilyTypes == null || selectedFamilyTypes.Count == 0)
                return Result.Cancelled;

            // Step 5: Query for actual elements of selected family types
            var selectedFamilyTypeKeys = selectedFamilyTypes
                .Where(ft => ft.ContainsKey("_FamilyTypeKey"))
                .Select(ft => (FamilyTypeKey)ft["_FamilyTypeKey"])
                .ToList();

            var familyTypeElements = QueryElementsForFamilyTypes(documentsToQuery, selectedFamilyTypeKeys, token, currentSessionId, uiapp);

            // Step 6: Gather selection items from query results
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

                    var tasks = new List<Task<(DocumentInfo DocInfo, RoslynResponse Response)>>();

                    foreach (var docInfo in remoteDocuments)
                    {
                        var escapedTitle = docInfo.DocumentTitle.Replace("\\", "\\\\").Replace("\"", "\\\"");
                        var query = $@"var docTitle = ""{escapedTitle}"";
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
    var elementTypes = new FilteredElementCollector(targetDoc).OfClass(typeof(ElementType)).Cast<ElementType>().ToList();
    var typeInstanceCounts = new Dictionary<ElementId, int>();
    var allInstances = new FilteredElementCollector(targetDoc).WhereElementIsNotElementType()
        .Where(x => x.GetTypeId() != null && x.GetTypeId() != ElementId.InvalidElementId).ToList();
    foreach (var instance in allInstances)
    {{
        ElementId typeId = instance.GetTypeId();
        if (!typeInstanceCounts.ContainsKey(typeId)) typeInstanceCounts[typeId] = 0;
        typeInstanceCounts[typeId]++;
    }}
    foreach (var elementType in elementTypes)
    {{
        string typeName = elementType.Name;
        string familyName = """";
        string categoryName = """";
        var fs = elementType as FamilySymbol;
        if (fs != null)
        {{
            familyName = fs.Family.Name;
            categoryName = fs.Category != null ? fs.Category.Name : ""N/A"";
            if (categoryName.Contains(""Import Symbol"") || familyName.Contains(""Import Symbol"")) continue;
        }}
        else
        {{
            Parameter familyParam = elementType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
            familyName = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString())) ? familyParam.AsString() : ""System Type"";
            categoryName = elementType.Category != null ? elementType.Category.Name : ""N/A"";
            if (categoryName.Contains(""Import Symbol"") || familyName.Contains(""Import Symbol"")) continue;
        }}
        int count = typeInstanceCounts.ContainsKey(elementType.Id) ? typeInstanceCounts[elementType.Id] : 0;
        Console.WriteLine(""FAMILYTYPE|"" + categoryName + ""|"" + familyName + ""|"" + typeName + ""|"" + count);
    }}
}}";

                        var task = SendRoslynRequestAsync(client, docInfo, query, token);
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
                                var (docInfo, jsonResponse) = task.Result;
                                if (jsonResponse != null && jsonResponse.Success && !string.IsNullOrWhiteSpace(jsonResponse.Output))
                                {
                                    ParseFamilyTypeCountResponse(jsonResponse.Output, docInfo, result);
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

    private void ParseFamilyTypeCountResponse(string output, DocumentInfo docInfo,
        Dictionary<FamilyTypeKey, Dictionary<string, int>> result)
    {
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
                }
            }
        }
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
        if (remoteDocuments.Count > 0)
        {
            // Build family type filter for query
            var familyTypeFilters = new StringBuilder();
            familyTypeFilters.Append("var familyTypes = new List<Tuple<string, string, string>> {");
            bool first = true;
            foreach (var ft in familyTypeKeys)
            {
                if (!first) familyTypeFilters.Append(",");
                familyTypeFilters.Append($"Tuple.Create(\"{EscapeForCSharp(ft.Category)}\", \"{EscapeForCSharp(ft.Family)}\", \"{EscapeForCSharp(ft.TypeName)}\")");
                first = false;
            }
            familyTypeFilters.Append("};");

            using (var handler = new HttpClientHandler())
            {
                handler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;

                using (var client = new HttpClient(handler))
                {
                    client.Timeout = TimeSpan.FromSeconds(120);

                    var tasks = new List<Task<(DocumentInfo DocInfo, RoslynResponse Response)>>();

                    foreach (var docInfo in remoteDocuments)
                    {
                        var escapedTitle = docInfo.DocumentTitle.Replace("\\", "\\\\").Replace("\"", "\\\"");
                        var query = $@"var docTitle = ""{escapedTitle}"";
{familyTypeFilters}
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
    var elementTypes = new FilteredElementCollector(targetDoc).OfClass(typeof(ElementType)).Cast<ElementType>().ToList();
    var matchingTypeIds = new HashSet<ElementId>();
    var typeIdToKey = new Dictionary<ElementId, Tuple<string, string, string>>();
    foreach (var elementType in elementTypes)
    {{
        string typeName = elementType.Name;
        string familyName = """";
        string categoryName = """";
        var fs = elementType as FamilySymbol;
        if (fs != null)
        {{
            familyName = fs.Family.Name;
            categoryName = fs.Category != null ? fs.Category.Name : ""N/A"";
        }}
        else
        {{
            Parameter familyParam = elementType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
            familyName = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString())) ? familyParam.AsString() : ""System Type"";
            categoryName = elementType.Category != null ? elementType.Category.Name : ""N/A"";
        }}
        var key = Tuple.Create(categoryName, familyName, typeName);
        if (familyTypes.Any(ft => ft.Item1 == key.Item1 && ft.Item2 == key.Item2 && ft.Item3 == key.Item3))
        {{
            matchingTypeIds.Add(elementType.Id);
            typeIdToKey[elementType.Id] = key;
        }}
    }}
    var instances = new FilteredElementCollector(targetDoc).WhereElementIsNotElementType()
        .Where(x => x.GetTypeId() != null && matchingTypeIds.Contains(x.GetTypeId())).ToList();
    foreach (var elem in instances)
    {{
        var key = typeIdToKey[elem.GetTypeId()];
        Console.WriteLine(""ELEMENT|"" + key.Item1 + ""|"" + key.Item2 + ""|"" + key.Item3 + ""|"" + elem.UniqueId + ""|"" + elem.Id.IntegerValue);
    }}
}}";

                        var task = SendRoslynRequestAsync(client, docInfo, query, token);
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
                                var (docInfo, jsonResponse) = task.Result;
                                if (jsonResponse != null && jsonResponse.Success && !string.IsNullOrWhiteSpace(jsonResponse.Output))
                                {
                                    ParseElementsResponse(jsonResponse.Output, docInfo, familyTypeKeys, result);
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

    private string EscapeForCSharp(string value)
    {
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    private void ParseElementsResponse(string output, DocumentInfo docInfo, List<FamilyTypeKey> familyTypeKeys,
        Dictionary<FamilyTypeKey, Dictionary<string, List<ElementInfo>>> result)
    {
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
