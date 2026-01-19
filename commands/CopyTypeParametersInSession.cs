using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

/// <summary>
/// Copies type parameter values from selected types to matching types in other
/// open documents within the current Revit session.
/// </summary>
[Transaction(TransactionMode.Manual)]
public class CopyTypeParametersInSession : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document activeDoc = uidoc.Document;

        using (var executionLog = CommandExecutionLogger.Start("CopyTypeParametersInSession", commandData))
        using (var diagnostics = CommandDiagnostics.StartCommand("CopyTypeParametersInSession", uiapp))
        {
            try
            {
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

                // Step 3: Find other documents with matching types
                var targetDocuments = FindDocumentsWithMatchingTypes(app, activeDoc, sourceTypeData, diagnostics);
                if (targetDocuments.Count == 0)
                {
                    TaskDialog.Show("Info", "No other open documents have matching types.");
                    executionLog.SetResult(Result.Cancelled);
                    return Result.Cancelled;
                }

                // Step 4: Show document selection UI
                var selectedDocuments = ShowDocumentSelectionUI(targetDocuments, diagnostics);
                if (selectedDocuments == null || selectedDocuments.Count == 0)
                {
                    diagnostics.Log("User cancelled document selection");
                    executionLog.SetResult(Result.Cancelled);
                    return Result.Cancelled;
                }

                diagnostics.Log($"Selected {selectedDocuments.Count} target document(s)");

                // Step 5: Copy parameters to matching types in selected documents
                int successCount = 0;
                int failCount = 0;
                var errors = new List<string>();

                foreach (var targetDocInfo in selectedDocuments)
                {
                    Document targetDoc = targetDocInfo.Document;
                    diagnostics.Log($"Processing document: {targetDoc.Title}");

                    var result = CopyParametersToDocument(targetDoc, sourceTypeData, diagnostics);
                    successCount += result.SuccessCount;
                    failCount += result.FailCount;
                    errors.AddRange(result.Errors);
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

        // No types in selection - show type selection UI (like SelectFamilyTypesInViews)
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
                // Skip DWG import symbols
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

        // Sort entries
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

        // Show DataGrid
        CustomGUIs.SetCurrentUIDocument(uidoc);
        var propertyNames = new List<string> { "Category", "Family", "Type Name", "Count" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, false);

        if (selectedEntries == null || selectedEntries.Count == 0)
        {
            return null;
        }

        // Convert selected entries back to Element list
        var selectedTypes = new List<Element>();
        foreach (var entry in selectedEntries)
        {
            if (entry.TryGetValue("ElementIdObject", out var idObj) && idObj is ElementId id)
            {
                Element elem = doc.GetElement(id);
                if (elem != null)
                {
                    selectedTypes.Add(elem);
                }
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

            // Collect all editable parameters
            foreach (Parameter param in typeElem.Parameters)
            {
                // Skip read-only and certain built-in parameters
                if (param.IsReadOnly) continue;
                if (!param.HasValue) continue;

                // Get parameter value based on storage type
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
                        // Skip ElementId parameters as they are document-specific
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
        {
            return fs.Family?.Name ?? "";
        }

        return GetSystemFamilyName(typeElement) ?? "System Type";
    }

    /// <summary>
    /// Finds documents that have types matching the source types.
    /// </summary>
    private List<DocumentMatchInfo> FindDocumentsWithMatchingTypes(
        Autodesk.Revit.ApplicationServices.Application app,
        Document activeDoc,
        List<TypeParameterData> sourceTypeData,
        CommandDiagnostics.DiagnosticSession diagnostics)
    {
        var result = new List<DocumentMatchInfo>();

        // Build lookup for source type keys
        var sourceTypeKeys = new HashSet<string>(
            sourceTypeData.Select(t => $"{t.CategoryName}|{t.FamilyName}|{t.TypeName}"));

        foreach (Document doc in app.Documents)
        {
            if (doc.IsLinked) continue;
            if (doc.IsFamilyDocument) continue;
            if (doc.Title == activeDoc.Title && doc.PathName == activeDoc.PathName) continue;

            // Find matching types in this document
            var matchingTypes = new List<Element>();

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
                {
                    matchingTypes.Add(typeElem);
                }
            }

            if (matchingTypes.Count > 0)
            {
                result.Add(new DocumentMatchInfo
                {
                    Document = doc,
                    MatchingTypeCount = matchingTypes.Count
                });
                diagnostics.Log($"Document '{doc.Title}' has {matchingTypes.Count} matching type(s)");
            }
        }

        return result;
    }

    /// <summary>
    /// Shows document selection UI.
    /// </summary>
    private List<DocumentMatchInfo> ShowDocumentSelectionUI(
        List<DocumentMatchInfo> documents,
        CommandDiagnostics.DiagnosticSession diagnostics)
    {
        var gridData = new List<Dictionary<string, object>>();
        var columns = new List<string> { "Document", "Matching Types", "Path" };

        foreach (var docInfo in documents)
        {
            var row = new Dictionary<string, object>
            {
                ["Document"] = docInfo.Document.Title,
                ["Matching Types"] = docInfo.MatchingTypeCount,
                ["Path"] = string.IsNullOrEmpty(docInfo.Document.PathName) ? "(Not saved)" : docInfo.Document.PathName,
                ["_Document"] = docInfo.Document
            };
            gridData.Add(row);
        }

        gridData = gridData.OrderBy(r => r["Document"].ToString()).ToList();

        var selectedRows = CustomGUIs.DataGrid(gridData, columns, false);

        if (selectedRows == null || selectedRows.Count == 0)
            return null;

        var result = new List<DocumentMatchInfo>();
        foreach (var row in selectedRows)
        {
            if (row.TryGetValue("_Document", out var docObj) && docObj is Document doc)
            {
                result.Add(new DocumentMatchInfo { Document = doc });
            }
        }

        return result;
    }

    /// <summary>
    /// Copies parameters to matching types in a target document.
    /// </summary>
    private CopyResult CopyParametersToDocument(
        Document targetDoc,
        List<TypeParameterData> sourceTypeData,
        CommandDiagnostics.DiagnosticSession diagnostics)
    {
        var result = new CopyResult();

        // Build lookup for source types
        var sourceTypeLookup = sourceTypeData.ToDictionary(
            t => $"{t.CategoryName}|{t.FamilyName}|{t.TypeName}",
            t => t);

        // Find and process matching types
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
            {
                typesToModify.Add((typeElem, sourceData));
            }
        }

        if (typesToModify.Count == 0)
        {
            return result;
        }

        // Apply parameter changes in a transaction
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
                        {
                            continue;
                        }

                        // Verify storage type matches
                        if (targetParam.StorageType != paramData.StorageType)
                        {
                            continue;
                        }

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
                        catch
                        {
                            // Skip parameters that fail to set
                        }
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

    #region Helper Classes

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

    private class DocumentMatchInfo
    {
        public Document Document { get; set; }
        public int MatchingTypeCount { get; set; }
    }

    private class CopyResult
    {
        public int SuccessCount { get; set; }
        public int FailCount { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }

    #endregion
}
