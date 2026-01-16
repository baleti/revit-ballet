using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
#if NET8_0_OR_GREATER
using System.Net.Http;
#endif

/// <summary>
/// Copies type parameter values from selected types to matching types in documents
/// across all Revit sessions in the network.
/// </summary>
[Transaction(TransactionMode.Manual)]
public class CopyTypeParametersInNetwork : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document activeDoc = uidoc.Document;

        using (var executionLog = CommandExecutionLogger.Start("CopyTypeParametersInNetwork", commandData))
        using (var diagnostics = CommandDiagnostics.StartCommand("CopyTypeParametersInNetwork", uiapp))
        {
            try
            {
                // Read auth token
                string tokenPath = Path.Combine(PathHelper.RuntimeDirectory, "network", "token");
                if (!File.Exists(tokenPath))
                {
                    TaskDialog.Show("Error", "Network token not found. Ensure Revit Ballet server is running.");
                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }
                string authToken = File.ReadAllText(tokenPath).Trim();

                // Step 1: Get source types from selection or prompt user
                var sourceTypes = GetSourceTypes(uidoc, activeDoc, diagnostics);
                if (sourceTypes == null || sourceTypes.Count == 0)
                {
                    diagnostics.Log("No types selected");
                    executionLog.SetResult(Result.Cancelled);
                    return Result.Cancelled;
                }

                diagnostics.Log($"Selected {sourceTypes.Count} source type(s)");

                // Step 2: Collect parameter data from source types
                var sourceTypeData = CollectTypeParameterData(sourceTypes, activeDoc, diagnostics);
                if (sourceTypeData.Count == 0)
                {
                    TaskDialog.Show("Info", "No editable parameters found on selected types.");
                    executionLog.SetResult(Result.Cancelled);
                    return Result.Cancelled;
                }

                // Step 3: Read documents file
                string documentsFilePath = Path.Combine(PathHelper.RuntimeDirectory, "documents");
                if (!File.Exists(documentsFilePath))
                {
                    TaskDialog.Show("Error", "No active documents found. Document registry file does not exist.");
                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }

                var networkDocuments = ParseDocumentsFile(documentsFilePath);
                if (networkDocuments.Count == 0)
                {
                    TaskDialog.Show("Error", "No active documents found in registry.");
                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }

                string currentSessionId = RevitBallet.RevitBallet.SessionId;
                diagnostics.Log($"Current session: {currentSessionId}, Found {networkDocuments.Count} document(s) in network");

                // Step 4: Find sessions/documents with matching types
                // TODO: Re-enable after making DataGrid asynchronous
                // var targetDocuments = FindNetworkDocumentsWithMatchingTypes(
                //     app, activeDoc, sessions, currentSessionId, sourceTypeData, authToken, diagnostics);
                //
                // if (targetDocuments.Count == 0)
                // {
                //     TaskDialog.Show("Info", "No other documents in the network have matching types.");
                //     executionLog.SetResult(Result.Cancelled);
                //     return Result.Cancelled;
                // }

                // Temporary: Show all documents without checking for matching types
                var targetDocuments = new List<NetworkDocumentInfo>();
                foreach (var docEntry in networkDocuments)
                {
                    bool isCurrentSession = docEntry.SessionId == currentSessionId;

                    // Skip the active document
                    if (docEntry.DocumentTitle == activeDoc.Title &&
                        (string.IsNullOrEmpty(docEntry.DocumentPath) || docEntry.DocumentPath == activeDoc.PathName))
                        continue;

                    targetDocuments.Add(new NetworkDocumentInfo
                    {
                        SessionId = docEntry.SessionId,
                        DocumentTitle = docEntry.DocumentTitle,
                        DocumentPath = docEntry.DocumentPath,
                        Hostname = docEntry.Hostname,
                        Port = docEntry.Port,
                        MatchingTypeCount = 0, // Unknown - not checked
                        IsLocal = isCurrentSession
                    });
                }

                if (targetDocuments.Count == 0)
                {
                    TaskDialog.Show("Info", "No other documents found in the network.");
                    executionLog.SetResult(Result.Cancelled);
                    return Result.Cancelled;
                }

                // Step 5: Show document selection UI
                var selectedDocuments = ShowDocumentSelectionUI(targetDocuments, diagnostics);
                if (selectedDocuments == null || selectedDocuments.Count == 0)
                {
                    diagnostics.Log("User cancelled document selection");
                    executionLog.SetResult(Result.Cancelled);
                    return Result.Cancelled;
                }

                diagnostics.Log($"Selected {selectedDocuments.Count} target document(s)");

                // Step 6: Copy parameters to matching types in selected documents
                int successCount = 0;
                int failCount = 0;
                var errors = new List<string>();

                // Group by session for efficient processing
                var localDocs = selectedDocuments.Where(d => d.IsLocal).ToList();
                var remoteDocs = selectedDocuments.Where(d => !d.IsLocal).GroupBy(d => d.SessionId).ToList();

                // Process local documents
                foreach (var docInfo in localDocs)
                {
                    diagnostics.Log($"Processing local document: {docInfo.DocumentTitle}");

                    var result = CopyParametersToLocalDocument(app, docInfo.DocumentTitle, sourceTypeData, diagnostics);
                    successCount += result.SuccessCount;
                    failCount += result.FailCount;
                    errors.AddRange(result.Errors);
                }

                // Process remote documents
                foreach (var sessionGroup in remoteDocs)
                {
                    foreach (var docInfo in sessionGroup)
                    {
                        diagnostics.Log($"Processing remote document: {docInfo.DocumentTitle} in session {docInfo.SessionId}");

                        var result = CopyParametersToRemoteDocument(
                            docInfo, sourceTypeData, authToken, diagnostics);
                        successCount += result.SuccessCount;
                        failCount += result.FailCount;
                        errors.AddRange(result.Errors);
                    }
                }

                diagnostics.Log($"Copy complete: {successCount} succeeded, {failCount} failed");

                // Only show dialog on errors
                if (failCount > 0 && errors.Count > 0)
                {
                    TaskDialog.Show("Partial Success",
                        $"Copied parameters to {successCount} type(s).\n\nErrors ({failCount}):\n" +
                        string.Join("\n", errors.Take(10)) +
                        (errors.Count > 10 ? $"\n... and {errors.Count - 10} more" : ""));
                }

                executionLog.SetResult(Result.Succeeded);
                return Result.Succeeded;
            }
            catch (OperationCanceledException)
            {
                executionLog.SetResult(Result.Cancelled);
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to copy type parameters: {ex.Message}");
                diagnostics.LogError($"Exception: {ex}");
                executionLog.SetResult(Result.Failed);
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Gets source types from current selection or prompts user to select.
    /// </summary>
    private List<Element> GetSourceTypes(UIDocument uidoc, Document doc, CommandDiagnostics.DiagnosticSession diagnostics)
    {
        var result = new List<Element>();

        // Get current selection
        ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();

        if (currentSelection != null && currentSelection.Count > 0)
        {
            // Process selection - convert instances to types
            foreach (ElementId id in currentSelection)
            {
                Element elem = doc.GetElement(id);
                if (elem == null) continue;

                if (elem is ElementType)
                {
                    result.Add(elem);
                }
                else
                {
                    // Instance - get its type
                    ElementId typeId = elem.GetTypeId();
                    if (typeId != null && typeId != ElementId.InvalidElementId)
                    {
                        Element typeElem = doc.GetElement(typeId);
                        if (typeElem != null && !result.Any(e => e.Id == typeElem.Id))
                        {
                            result.Add(typeElem);
                        }
                    }
                }
            }

            if (result.Count > 0)
            {
                diagnostics.Log($"Got {result.Count} type(s) from selection");
                return result;
            }
        }

        // No types in selection - show type selection UI
        diagnostics.Log("No types in selection, showing type picker");
        return ShowTypeSelectionUI(uidoc, doc, diagnostics);
    }

    /// <summary>
    /// Shows type selection UI similar to SelectFamilyTypesInViews.
    /// </summary>
    private List<Element> ShowTypeSelectionUI(UIDocument uidoc, Document doc, CommandDiagnostics.DiagnosticSession diagnostics)
    {
        var typeEntries = new List<Dictionary<string, object>>();
        var typeElementMap = new Dictionary<string, Element>();
        var typeInstanceCounts = new Dictionary<string, int>();

        // Count instances per type
        var allInstances = new FilteredElementCollector(doc)
            .WhereElementIsNotElementType()
            .Where(e => e.GetTypeId() != null && e.GetTypeId() != ElementId.InvalidElementId)
            .ToList();

        var instanceCountByTypeId = new Dictionary<ElementId, int>();
        foreach (var instance in allInstances)
        {
            ElementId typeId = instance.GetTypeId();
            if (!instanceCountByTypeId.ContainsKey(typeId))
                instanceCountByTypeId[typeId] = 0;
            instanceCountByTypeId[typeId]++;
        }

        // Collect all element types
        var elementTypes = new FilteredElementCollector(doc)
            .OfClass(typeof(ElementType))
            .Cast<ElementType>()
            .Where(t => t.Category != null)
            .ToList();

        foreach (var elementType in elementTypes)
        {
            string typeName = elementType.Name;
            string familyName = "";
            string categoryName = elementType.Category?.Name ?? "N/A";

            if (elementType is FamilySymbol fs)
            {
                familyName = fs.Family?.Name ?? "";
                if (familyName.Contains("Import Symbol")) continue;
            }
            else
            {
                familyName = GetSystemFamilyName(elementType) ?? "System Type";
            }

            int count = instanceCountByTypeId.ContainsKey(elementType.Id)
                ? instanceCountByTypeId[elementType.Id]
                : 0;

            string uniqueKey = $"{categoryName}:{familyName}:{typeName}";
            if (!typeElementMap.ContainsKey(uniqueKey))
            {
                typeElementMap[uniqueKey] = elementType;
                typeInstanceCounts[uniqueKey] = count;
            }
        }

        // Build entries for DataGrid
        foreach (var kvp in typeElementMap)
        {
            string[] parts = kvp.Key.Split(':');
            var entry = new Dictionary<string, object>
            {
                { "Category", parts[0] },
                { "Family", parts[1] },
                { "Type Name", parts[2] },
                { "Count", typeInstanceCounts[kvp.Key] },
                { "ElementIdObject", kvp.Value.Id }
            };
            typeEntries.Add(entry);
        }

        typeEntries = typeEntries
            .OrderBy(e => e["Category"].ToString())
            .ThenBy(e => e["Family"].ToString())
            .ThenBy(e => e["Type Name"].ToString())
            .ToList();

        if (typeEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No types found in document.");
            return null;
        }

        CustomGUIs.SetCurrentUIDocument(uidoc);
        var propertyNames = new List<string> { "Category", "Family", "Type Name", "Count" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, false);

        if (selectedEntries == null || selectedEntries.Count == 0)
            return null;

        var selectedTypes = new List<Element>();
        foreach (var entry in selectedEntries)
        {
            if (entry.TryGetValue("ElementIdObject", out var idObj) && idObj is ElementId id)
            {
                Element elem = doc.GetElement(id);
                if (elem != null)
                    selectedTypes.Add(elem);
            }
        }

        return selectedTypes;
    }

    private string GetSystemFamilyName(Element typeElement)
    {
        Parameter familyParam = typeElement.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
        return (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
            ? familyParam.AsString()
            : null;
    }

    /// <summary>
    /// Collects parameter data from source types.
    /// </summary>
    private List<TypeParameterData> CollectTypeParameterData(List<Element> sourceTypes, Document doc, CommandDiagnostics.DiagnosticSession diagnostics)
    {
        var result = new List<TypeParameterData>();

        foreach (Element typeElem in sourceTypes)
        {
            var typeData = new TypeParameterData
            {
                CategoryName = typeElem.Category?.Name ?? "",
                FamilyName = GetFamilyNameFromType(typeElem),
                TypeName = typeElem.Name,
                Parameters = new List<ParameterValueData>()
            };

            foreach (Parameter param in typeElem.Parameters)
            {
                if (param.IsReadOnly) continue;
                if (!param.HasValue) continue;

                var paramData = new ParameterValueData
                {
                    Name = param.Definition.Name,
                    StorageType = param.StorageType,
                    IsShared = param.IsShared
                };

                switch (param.StorageType)
                {
                    case StorageType.String:
                        paramData.StringValue = param.AsString();
                        break;
                    case StorageType.Integer:
                        paramData.IntegerValue = param.AsInteger();
                        break;
                    case StorageType.Double:
                        paramData.DoubleValue = param.AsDouble();
                        break;
                    case StorageType.ElementId:
                        continue;
                    default:
                        continue;
                }

                typeData.Parameters.Add(paramData);
            }

            if (typeData.Parameters.Count > 0)
            {
                result.Add(typeData);
                diagnostics.Log($"Collected {typeData.Parameters.Count} parameters from {typeData.FamilyName}:{typeData.TypeName}");
            }
        }

        return result;
    }

    private string GetFamilyNameFromType(Element typeElement)
    {
        if (typeElement is FamilySymbol fs)
            return fs.Family?.Name ?? "";
        return GetSystemFamilyName(typeElement) ?? "System Type";
    }

    /// <summary>
    /// Parses the documents file.
    /// </summary>
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
                if ((DateTime.Now - doc.LastHeartbeat).TotalSeconds <= 120 &&
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

    /// <summary>
    /// Finds documents across the network that have matching types.
    /// Queries remote sessions in PARALLEL for performance.
    /// </summary>
    private List<NetworkDocumentInfo> FindNetworkDocumentsWithMatchingTypes(
        Autodesk.Revit.ApplicationServices.Application app,
        Document activeDoc,
        List<DocumentInfo> documents,
        string currentSessionId,
        List<TypeParameterData> sourceTypeData,
        string authToken,
        CommandDiagnostics.DiagnosticSession diagnostics)
    {
        var result = new ConcurrentBag<NetworkDocumentInfo>();

        // Build type keys for matching
        var sourceTypeKeys = new HashSet<string>(
            sourceTypeData.Select(t => $"{t.CategoryName}|{t.FamilyName}|{t.TypeName}"));

        // Separate local and remote documents
        var localDocs = documents.Where(d => d.SessionId == currentSessionId).ToList();
        var remoteDocs = documents.Where(d => d.SessionId != currentSessionId).ToList();

        // Check local documents (fast, no network)
        foreach (Document doc in app.Documents)
        {
            if (doc.IsLinked) continue;
            if (doc.IsFamilyDocument) continue;
            if (doc.Title == activeDoc.Title && doc.PathName == activeDoc.PathName) continue;

            // Find matching entry in localDocs to get network info
            var docEntry = localDocs.FirstOrDefault(d =>
                d.DocumentTitle == doc.Title &&
                (string.IsNullOrEmpty(d.DocumentPath) || d.DocumentPath == doc.PathName));

            if (docEntry == null) continue;

            int matchCount = CountMatchingTypesInLocalDocument(doc, sourceTypeKeys);
            if (matchCount > 0)
            {
                result.Add(new NetworkDocumentInfo
                {
                    SessionId = docEntry.SessionId,
                    DocumentTitle = doc.Title,
                    DocumentPath = doc.PathName ?? "",
                    Hostname = docEntry.Hostname,
                    Port = docEntry.Port,
                    MatchingTypeCount = matchCount,
                    IsLocal = true
                });
                diagnostics.Log($"Local document '{doc.Title}' has {matchCount} matching type(s)");
            }
        }

        // Run remote queries in parallel
        if (remoteDocs.Count > 0)
        {
            diagnostics.Log($"Querying {remoteDocs.Count} remote document(s) in parallel...");

            Parallel.ForEach(remoteDocs, new ParallelOptions { MaxDegreeOfParallelism = 8 }, docInfo =>
            {
                try
                {
                    // Skip active document
                    if (docInfo.DocumentTitle == activeDoc.Title) return;

                    int matchCount = QueryRemoteDocumentForMatchingTypes(docInfo, sourceTypeKeys, authToken, diagnostics);
                    if (matchCount > 0)
                    {
                        result.Add(new NetworkDocumentInfo
                        {
                            SessionId = docInfo.SessionId,
                            DocumentTitle = docInfo.DocumentTitle,
                            DocumentPath = docInfo.DocumentPath,
                            Hostname = docInfo.Hostname,
                            Port = docInfo.Port,
                            MatchingTypeCount = matchCount,
                            IsLocal = false
                        });
                        diagnostics.Log($"Remote document '{docInfo.DocumentTitle}' in session {docInfo.SessionId} has {matchCount} matching type(s)");
                    }
                }
                catch (Exception ex)
                {
                    diagnostics.Log($"Failed to query '{docInfo.DocumentTitle}' in session {docInfo.SessionId}: {ex.Message}");
                }
            });

            diagnostics.Log($"Parallel queries complete. Found {result.Count} document(s) with matching types.");
        }

        return result.ToList();
    }

    private int CountMatchingTypesInLocalDocument(Document doc, HashSet<string> sourceTypeKeys)
    {
        int count = 0;

        var elementTypes = new FilteredElementCollector(doc)
            .OfClass(typeof(ElementType))
            .Cast<ElementType>()
            .Where(t => t.Category != null)
            .ToList();

        foreach (var typeElem in elementTypes)
        {
            string categoryName = typeElem.Category?.Name ?? "";
            string familyName = GetFamilyNameFromType(typeElem);
            string typeName = typeElem.Name;

            string key = $"{categoryName}|{familyName}|{typeName}";
            if (sourceTypeKeys.Contains(key))
                count++;
        }

        return count;
    }

    private int QueryRemoteDocumentForMatchingTypes(
        DocumentInfo docInfo,
        HashSet<string> sourceTypeKeys,
        string authToken,
        CommandDiagnostics.DiagnosticSession diagnostics)
    {
        try
        {
            // Build type keys array for the query
            var typeKeysArray = sourceTypeKeys.Select(k => $"\"{EscapeForCSharp(k)}\"").ToList();
            var typeKeysArrayStr = string.Join(", ", typeKeysArray);

            // Build Roslyn query to count matching types
            var query = $@"
var sourceTypeKeys = new HashSet<string>(new[] {{ {typeKeysArrayStr} }});
int matchCount = 0;

var elementTypes = new FilteredElementCollector(Doc)
    .OfClass(typeof(ElementType))
    .Cast<ElementType>()
    .Where(t => t.Category != null)
    .ToList();

foreach (var typeElem in elementTypes)
{{
    string categoryName = typeElem.Category?.Name ?? """";
    string familyName = """";

    if (typeElem is FamilySymbol fs)
        familyName = fs.Family?.Name ?? """";
    else
    {{
        var familyParam = typeElem.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
        familyName = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
            ? familyParam.AsString()
            : ""System Type"";
    }}

    string typeName = typeElem.Name;
    string key = categoryName + ""|"" + familyName + ""|"" + typeName;

    if (sourceTypeKeys.Contains(key))
        matchCount++;
}}

Console.WriteLine(""MATCH_COUNT|"" + matchCount);
";

            var response = SendRoslynQuery(docInfo.Port, authToken, query);
            if (response.Success && !string.IsNullOrEmpty(response.Output))
            {
                // Parse match count from output
                foreach (var line in response.Output.Split('\n'))
                {
                    if (line.StartsWith("MATCH_COUNT|"))
                    {
                        if (int.TryParse(line.Substring("MATCH_COUNT|".Length).Trim(), out int count))
                            return count;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            diagnostics.Log($"Failed to query remote session {docInfo.SessionId}: {ex.Message}");
        }

        return 0;
    }

    /// <summary>
    /// Shows document selection UI.
    /// </summary>
    private List<NetworkDocumentInfo> ShowDocumentSelectionUI(
        List<NetworkDocumentInfo> documents,
        CommandDiagnostics.DiagnosticSession diagnostics)
    {
        var gridData = new List<Dictionary<string, object>>();
        var columns = new List<string> { "Document", "Matching Types", "Session", "Hostname" };

        foreach (var docInfo in documents)
        {
            var row = new Dictionary<string, object>
            {
                ["Document"] = docInfo.DocumentTitle,
                ["Matching Types"] = docInfo.MatchingTypeCount,
                ["Session"] = docInfo.IsLocal ? "(Current)" : docInfo.SessionId,
                ["Hostname"] = docInfo.Hostname,
                ["_DocInfo"] = docInfo
            };
            gridData.Add(row);
        }

        gridData = gridData.OrderBy(r => r["Document"].ToString()).ToList();

        var selectedRows = CustomGUIs.DataGrid(gridData, columns, false);

        if (selectedRows == null || selectedRows.Count == 0)
            return null;

        var result = new List<NetworkDocumentInfo>();
        foreach (var row in selectedRows)
        {
            if (row.TryGetValue("_DocInfo", out var docInfoObj) && docInfoObj is NetworkDocumentInfo docInfo)
                result.Add(docInfo);
        }

        return result;
    }

    /// <summary>
    /// Copies parameters to matching types in a local document.
    /// </summary>
    private CopyResult CopyParametersToLocalDocument(
        Autodesk.Revit.ApplicationServices.Application app,
        string documentTitle,
        List<TypeParameterData> sourceTypeData,
        CommandDiagnostics.DiagnosticSession diagnostics)
    {
        var result = new CopyResult();

        // Find the document
        Document targetDoc = null;
        foreach (Document doc in app.Documents)
        {
            if (!doc.IsLinked && doc.Title == documentTitle)
            {
                targetDoc = doc;
                break;
            }
        }

        if (targetDoc == null)
        {
            result.FailCount++;
            result.Errors.Add($"Document '{documentTitle}' not found");
            return result;
        }

        // Build lookup
        var sourceTypeLookup = sourceTypeData.ToDictionary(
            t => $"{t.CategoryName}|{t.FamilyName}|{t.TypeName}",
            t => t);

        // Find matching types
        var elementTypes = new FilteredElementCollector(targetDoc)
            .OfClass(typeof(ElementType))
            .Cast<ElementType>()
            .Where(t => t.Category != null)
            .ToList();

        var typesToModify = new List<(Element TargetType, TypeParameterData SourceData)>();

        foreach (var typeElem in elementTypes)
        {
            string categoryName = typeElem.Category?.Name ?? "";
            string familyName = GetFamilyNameFromType(typeElem);
            string typeName = typeElem.Name;

            string key = $"{categoryName}|{familyName}|{typeName}";
            if (sourceTypeLookup.TryGetValue(key, out var sourceData))
                typesToModify.Add((typeElem, sourceData));
        }

        if (typesToModify.Count == 0)
            return result;

        using (var trans = new Transaction(targetDoc, "Copy Type Parameters"))
        {
            trans.Start();

            foreach (var (targetType, sourceData) in typesToModify)
            {
                try
                {
                    int paramsSet = 0;

                    foreach (var paramData in sourceData.Parameters)
                    {
                        Parameter targetParam = targetType.LookupParameter(paramData.Name);
                        if (targetParam == null || targetParam.IsReadOnly)
                            continue;

                        if (targetParam.StorageType != paramData.StorageType)
                            continue;

                        try
                        {
                            switch (paramData.StorageType)
                            {
                                case StorageType.String:
                                    if (paramData.StringValue != null)
                                        targetParam.Set(paramData.StringValue);
                                    break;
                                case StorageType.Integer:
                                    targetParam.Set(paramData.IntegerValue);
                                    break;
                                case StorageType.Double:
                                    targetParam.Set(paramData.DoubleValue);
                                    break;
                            }
                            paramsSet++;
                        }
                        catch { }
                    }

                    if (paramsSet > 0)
                    {
                        result.SuccessCount++;
                        diagnostics.Log($"Set {paramsSet} parameter(s) on {sourceData.FamilyName}:{sourceData.TypeName}");
                    }
                }
                catch (Exception ex)
                {
                    result.FailCount++;
                    result.Errors.Add($"{sourceData.FamilyName}:{sourceData.TypeName} - {ex.Message}");
                }
            }

            trans.Commit();
        }

        return result;
    }

    /// <summary>
    /// Copies parameters to matching types in a remote document via Roslyn.
    /// </summary>
    private CopyResult CopyParametersToRemoteDocument(
        NetworkDocumentInfo docInfo,
        List<TypeParameterData> sourceTypeData,
        string authToken,
        CommandDiagnostics.DiagnosticSession diagnostics)
    {
        var result = new CopyResult();

        try
        {
            // Build a safer, more efficient parameter update script
            // Uses dictionary lookup instead of nested loops (O(n) instead of O(n*m))
            // Proper transaction handling with try-finally
            var sb = new StringBuilder();

            // Build the type lookup dictionary as C# code
            sb.AppendLine("int successCount = 0;");
            sb.AppendLine("int failCount = 0;");
            sb.AppendLine();
            sb.AppendLine("// Type key lookup: category|family|type -> parameter updates");
            sb.AppendLine("var typeUpdates = new Dictionary<string, Action<Element>>();");
            sb.AppendLine();

            // Generate an Action for each source type
            int typeIndex = 0;
            foreach (var typeData in sourceTypeData)
            {
                string categoryName = EscapeForCSharp(typeData.CategoryName);
                string familyName = EscapeForCSharp(typeData.FamilyName);
                string typeName = EscapeForCSharp(typeData.TypeName);
                string typeKey = $"{categoryName}|{familyName}|{typeName}";

                sb.AppendLine($"// Type {typeIndex}: {familyName}:{typeName}");
                sb.AppendLine($"typeUpdates[\"{typeKey}\"] = (typeElem) => {{");
                sb.AppendLine("    int paramsSet = 0;");
                sb.AppendLine("    Parameter p = null;");

                foreach (var paramData in typeData.Parameters)
                {
                    string paramName = EscapeForCSharp(paramData.Name);

                    sb.AppendLine($"    p = typeElem.LookupParameter(\"{paramName}\");");
                    sb.AppendLine("    if (p != null && !p.IsReadOnly) {");
                    sb.AppendLine("        try {");

                    switch (paramData.StorageType)
                    {
                        case StorageType.String:
                            string strValue = EscapeForCSharp(paramData.StringValue ?? "");
                            sb.AppendLine($"            p.Set(\"{strValue}\");");
                            break;
                        case StorageType.Integer:
                            sb.AppendLine($"            p.Set({paramData.IntegerValue});");
                            break;
                        case StorageType.Double:
                            sb.AppendLine($"            p.Set({paramData.DoubleValue.ToString(System.Globalization.CultureInfo.InvariantCulture)});");
                            break;
                    }

                    sb.AppendLine("            paramsSet++;");
                    sb.AppendLine("        } catch { }");
                    sb.AppendLine("    }");
                }

                sb.AppendLine("    if (paramsSet > 0) successCount++;");
                sb.AppendLine("};");
                sb.AppendLine();
                typeIndex++;
            }

            // Generate the main processing loop - single pass through types
            sb.AppendLine("// Process types in a single pass with proper transaction handling");
            sb.AppendLine("var trans = new Transaction(Doc, \"Copy Type Parameters\");");
            sb.AppendLine("try");
            sb.AppendLine("{");
            sb.AppendLine("    trans.Start();");
            sb.AppendLine();
            sb.AppendLine("    var collector = new FilteredElementCollector(Doc).OfClass(typeof(ElementType));");
            sb.AppendLine("    foreach (Element elem in collector)");
            sb.AppendLine("    {");
            sb.AppendLine("        try");
            sb.AppendLine("        {");
            sb.AppendLine("            var typeElem = elem as ElementType;");
            sb.AppendLine("            if (typeElem == null || typeElem.Category == null) continue;");
            sb.AppendLine();
            sb.AppendLine("            string catName = typeElem.Category.Name ?? \"\";");
            sb.AppendLine("            string famName = \"\";");
            sb.AppendLine("            if (typeElem is FamilySymbol fs)");
            sb.AppendLine("                famName = fs.Family?.Name ?? \"\";");
            sb.AppendLine("            else");
            sb.AppendLine("            {");
            sb.AppendLine("                var familyParam = typeElem.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);");
            sb.AppendLine("                famName = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))");
            sb.AppendLine("                    ? familyParam.AsString() : \"System Type\";");
            sb.AppendLine("            }");
            sb.AppendLine();
            sb.AppendLine("            string key = catName + \"|\" + famName + \"|\" + typeElem.Name;");
            sb.AppendLine("            if (typeUpdates.TryGetValue(key, out var updateAction))");
            sb.AppendLine("            {");
            sb.AppendLine("                updateAction(typeElem);");
            sb.AppendLine("            }");
            sb.AppendLine("        }");
            sb.AppendLine("        catch { failCount++; }");
            sb.AppendLine("    }");
            sb.AppendLine();
            sb.AppendLine("    trans.Commit();");
            sb.AppendLine("}");
            sb.AppendLine("catch");
            sb.AppendLine("{");
            sb.AppendLine("    if (trans.HasStarted()) trans.RollBack();");
            sb.AppendLine("    throw;");
            sb.AppendLine("}");
            sb.AppendLine("finally");
            sb.AppendLine("{");
            sb.AppendLine("    trans.Dispose();");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("Console.WriteLine(\"RESULT|\" + successCount + \"|\" + failCount);");

            string script = sb.ToString();

            // Write script to diagnostic file for debugging
            try
            {
                string scriptPath = Path.Combine(
                    PathHelper.RuntimeDirectory,
                    "diagnostics",
                    $"CopyTypeParams-Script-{DateTime.Now:yyyyMMdd-HHmmss-fff}.cs");
                Directory.CreateDirectory(Path.GetDirectoryName(scriptPath));
                File.WriteAllText(scriptPath, script);
                diagnostics.Log($"Script saved to: {scriptPath}");
            }
            catch { }

            var response = SendRoslynQuery(docInfo.Port, authToken, script);

            if (response.Success && !string.IsNullOrEmpty(response.Output))
            {
                foreach (var line in response.Output.Split('\n'))
                {
                    if (line.StartsWith("RESULT|"))
                    {
                        var parts = line.Split('|');
                        if (parts.Length >= 3)
                        {
                            if (int.TryParse(parts[1].Trim(), out int success))
                                result.SuccessCount = success;
                            if (int.TryParse(parts[2].Trim(), out int fail))
                                result.FailCount = fail;
                        }
                    }
                }
                diagnostics.Log($"Remote update complete: {result.SuccessCount} succeeded, {result.FailCount} failed");
            }
            else
            {
                result.FailCount++;
                string errorDetail = response.Error ?? "Unknown error";
                if (response.Diagnostics != null && response.Diagnostics.Length > 0)
                {
                    errorDetail += " | Diagnostics: " + string.Join("; ", response.Diagnostics);
                }
                result.Errors.Add($"Remote session error: {errorDetail}");
                diagnostics.LogError($"Remote update failed: {errorDetail}");
            }
        }
        catch (Exception ex)
        {
            result.FailCount++;
            result.Errors.Add($"Failed to update remote document: {ex.Message}");
            diagnostics.LogError($"Remote update failed: {ex.Message}");
        }

        return result;
    }

    private string EscapeForCSharp(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        return value
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\n", "\\n")
            .Replace("\r", "\\r")
            .Replace("\t", "\\t");
    }

    private RoslynResponse SendRoslynQuery(string port, string authToken, string query)
    {
        string url = $"https://127.0.0.1:{port}/roslyn";

#if NET8_0_OR_GREATER
        using (var handler = new HttpClientHandler())
        {
            handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;

            using (var client = new HttpClient(handler))
            {
                client.Timeout = TimeSpan.FromSeconds(120);
                client.DefaultRequestHeaders.Add("X-Auth-Token", authToken);

                var content = new StringContent(query, Encoding.UTF8, "text/plain");
                var response = client.PostAsync(url, content).Result;
                var responseText = response.Content.ReadAsStringAsync().Result;

                return Newtonsoft.Json.JsonConvert.DeserializeObject<RoslynResponse>(responseText);
            }
        }
#else
        var request = (HttpWebRequest)WebRequest.Create(url);
        request.Method = "POST";
        request.ContentType = "text/plain";
        request.Headers.Add("X-Auth-Token", authToken);
        request.Timeout = 120000;
        request.ServerCertificateValidationCallback = (sender, cert, chain, errors) => true;

        byte[] bodyBytes = Encoding.UTF8.GetBytes(query);
        request.ContentLength = bodyBytes.Length;

        using (var requestStream = request.GetRequestStream())
        {
            requestStream.Write(bodyBytes, 0, bodyBytes.Length);
        }

        using (var response = request.GetResponse())
        using (var reader = new StreamReader(response.GetResponseStream()))
        {
            var responseText = reader.ReadToEnd();
            return Newtonsoft.Json.JsonConvert.DeserializeObject<RoslynResponse>(responseText);
        }
#endif
    }

    #region Helper Classes

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

    private class TypeParameterData
    {
        public string CategoryName { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public List<ParameterValueData> Parameters { get; set; }
    }

    private class ParameterValueData
    {
        public string Name { get; set; }
        public StorageType StorageType { get; set; }
        public string StringValue { get; set; }
        public int IntegerValue { get; set; }
        public double DoubleValue { get; set; }
        public bool IsShared { get; set; }
    }

    private class NetworkDocumentInfo
    {
        public string SessionId { get; set; }
        public string DocumentTitle { get; set; }
        public string DocumentPath { get; set; }
        public string Hostname { get; set; }
        public string Port { get; set; }
        public int MatchingTypeCount { get; set; }
        public bool IsLocal { get; set; }
    }

    private class CopyResult
    {
        public int SuccessCount { get; set; }
        public int FailCount { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }

    private class RoslynResponse
    {
        public bool Success { get; set; }
        public string Output { get; set; }
        public string Error { get; set; }
        public string[] Diagnostics { get; set; }
    }

    #endregion
}
