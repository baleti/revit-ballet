using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public static class ElementDataHelper
{
    public static List<Dictionary<string, object>> GetElementData(UIDocument uiDoc, bool selectedOnly = false, bool includeParameters = false)
    {
        Document doc = uiDoc.Document;
        var elementData = new List<Dictionary<string, object>>();

        if (selectedOnly)
        {
            // Handle both regular elements and linked elements (via References)
            var selectedIds = uiDoc.Selection.GetElementIds();
            var selectedRefs = uiDoc.Selection.GetReferences();

            if (!selectedIds.Any() && !selectedRefs.Any())
                throw new InvalidOperationException("No elements are selected.");

            // Keep track of processed elements to avoid duplicates
            var processedElements = new HashSet<ElementId>();

            // Process regular elements from current document
            foreach (var id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element != null)
                {
                    processedElements.Add(id);
                    var data = GetElementDataDictionary(element, doc, null, null, null, includeParameters);
                    elementData.Add(data);
                }
            }

            // Process linked elements via References
            foreach (var reference in selectedRefs)
            {
                try
                {
                    // Get the linked instance
                    var linkedInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkedInstance != null)
                    {
                        Document linkedDoc = linkedInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            // Get the actual element in the linked document
                            var linkedElementId = reference.LinkedElementId;
                            if (linkedElementId != ElementId.InvalidElementId)
                            {
                                Element linkedElement = linkedDoc.GetElement(linkedElementId);
                                if (linkedElement != null)
                                {
                                    // Store the linked instance and element ID for later reference creation
                                    var data = GetElementDataDictionary(linkedElement, linkedDoc, linkedInstance.Name, linkedInstance, linkedElement.Id, includeParameters);
                                    elementData.Add(data);
                                }
                            }
                        }
                    }
                    else
                    {
                        // This might be a regular element selected via reference
                        // Only process if we haven't already processed it via selectedIds
                        if (!processedElements.Contains(reference.ElementId))
                        {
                            Element element = doc.GetElement(reference);
                            if (element != null)
                            {
                                processedElements.Add(reference.ElementId);
                                var data = GetElementDataDictionary(element, doc, null, null, null, includeParameters);
                                elementData.Add(data);
                            }
                        }
                    }
                }
                catch { /* Skip problematic references */ }
            }
        }
        else
        {
            // Get elements from active view (current document only)
            var elementIds = new FilteredElementCollector(doc, doc.ActiveView.Id).ToElementIds();
            foreach (var id in elementIds)
            {
                Element element = doc.GetElement(id);
                if (element != null)
                {
                    var data = GetElementDataDictionary(element, doc, null, null, null, includeParameters);
                    elementData.Add(data);
                }
            }
        }

        return elementData;
    }

    private static Dictionary<string, object> GetElementDataDictionary(Element element, Document elementDoc, string linkName, RevitLinkInstance linkInstance, ElementId linkedElementId, bool includeParameters = false)
    {
        string groupName = string.Empty;
        if (element.GroupId != null && element.GroupId != ElementId.InvalidElementId && element.GroupId.IntegerValue != -1)
        {
            if (elementDoc.GetElement(element.GroupId) is Group g)
                groupName = g.Name;
        }

        string ownerViewName = string.Empty;
        if (element.OwnerViewId != null && element.OwnerViewId != ElementId.InvalidElementId)
        {
            if (elementDoc.GetElement(element.OwnerViewId) is View v)
                ownerViewName = v.Name;
        }

        var data = new Dictionary<string, object>
        {
            ["Name"] = element.Name,
            ["Category"] = element.Category?.Name ?? string.Empty,
            ["Group"] = groupName,
            ["OwnerView"] = ownerViewName,
            ["Id"] = element.Id.IntegerValue,
            ["IsLinked"] = !string.IsNullOrEmpty(linkName),
            ["LinkName"] = linkName ?? string.Empty,
            ["ElementIdObject"] = element.Id, // Store full ElementId for selection
            ["LinkInstanceObject"] = linkInstance, // Store link instance for reference creation
            ["LinkedElementIdObject"] = linkedElementId // Store linked element ID for reference creation
        };

        data["DisplayName"] = element.Name;

        // Include parameters if requested
        if (includeParameters)
        {
            foreach (Parameter p in element.Parameters)
            {
                try
                {
                    string pName = p.Definition.Name;
                    string pValue = p.AsValueString() ?? p.AsString() ?? "None";

                    // Avoid conflicts with existing keys
                    if (!data.ContainsKey(pName))
                    {
                        data[pName] = pValue;
                    }
                    else
                    {
                        // Add suffix to avoid key collision
                        data[$"{pName}_param"] = pValue;
                    }
                }
                catch { /* Skip problematic parameters */ }
            }
        }

        return data;
    }
}

/// <summary>
/// Base class for commands that display Revit elements in a custom data‑grid for filtering and re‑selection.
/// Now supports elements from linked models.
/// </summary>
public abstract class FilterElementsBase : IExternalCommand
{
    public abstract bool SpanAllScreens { get; }
    public abstract bool UseSelectedElements { get; }
    public abstract bool IncludeParameters { get; }

    public Result Execute(ExternalCommandData cData, ref string message, ElementSet elements)
    {
        try
        {
            var uiDoc = cData.Application.ActiveUIDocument;
            var elementData = ElementDataHelper.GetElementData(uiDoc, UseSelectedElements, IncludeParameters);

            if (!elementData.Any())
            {
                TaskDialog.Show("Info", "No elements found.");
                return Result.Cancelled;
            }

            // Create a unique key for each element to map back to full data
            var elementDataMap = new Dictionary<string, Dictionary<string, object>>();
            var displayData = new List<Dictionary<string, object>>();
            
            for (int i = 0; i < elementData.Count; i++)
            {
                var data = elementData[i];
                // Create a unique key combining multiple properties
                string uniqueKey = $"{data["Id"]}_{data["Name"]}_{data["Category"]}_{data["LinkName"]}_{i}";
                
                // Store the full data in our map
                elementDataMap[uniqueKey] = data;
                
                // Create display data with the unique key
                var display = new Dictionary<string, object>(data);
                display["UniqueKey"] = uniqueKey;
                displayData.Add(display);
            }

            // Get property names, including UniqueKey but excluding internal object fields
            var propertyNames = displayData.First().Keys
                .Where(k => !k.EndsWith("Object"))
                .ToList();

            // Reorder to put most useful columns first
            var orderedProps = new List<string> { "DisplayName", "Category", "LinkName", "Group", "OwnerView", "Id" };
            var remainingProps = propertyNames.Except(orderedProps).OrderBy(p => p);
            propertyNames = orderedProps.Where(p => propertyNames.Contains(p))
                .Concat(remainingProps)
                .ToList();

            var chosenRows = CustomGUIs.DataGrid(displayData, propertyNames, SpanAllScreens);
            if (chosenRows.Count == 0)
                return Result.Cancelled;

            // Separate regular elements and linked elements
            var regularIds = new List<ElementId>();
            var linkedReferences = new List<Reference>();

            foreach (var row in chosenRows)
            {
                // Get the unique key to look up full data
                if (!row.TryGetValue("UniqueKey", out var keyObj) || !(keyObj is string uniqueKey))
                    continue;
                
                // Get the full data from our map
                if (!elementDataMap.TryGetValue(uniqueKey, out var fullData))
                    continue;
                
                // Check if this is a linked element
                if (fullData.TryGetValue("LinkInstanceObject", out var linkObj) && linkObj is RevitLinkInstance linkInstance &&
                    fullData.TryGetValue("LinkedElementIdObject", out var linkedIdObj) && linkedIdObj is ElementId linkedElementId)
                {
                    // This is a linked element - create reference the same way SelectCategories does
                    try
                    {
                        Document linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            Element linkedElement = linkedDoc.GetElement(linkedElementId);
                            if (linkedElement != null)
                            {
                                // Create reference for the element
                                Reference elemRef = new Reference(linkedElement);
                                Reference linkedRef = elemRef.CreateLinkReference(linkInstance);
                                
                                if (linkedRef != null)
                                {
                                    linkedReferences.Add(linkedRef);
                                }
                            }
                        }
                    }
                    catch { }
                }
                else if (fullData.TryGetValue("ElementIdObject", out var idObj) && idObj is ElementId elemId)
                {
                    // This is a regular element (not linked)
                    regularIds.Add(elemId);
                }
                // Handle backward compatibility - if someone stored just the integer ID
                else if (fullData.TryGetValue("Id", out var intId) && intId is int id)
                {
                    regularIds.Add(new ElementId(id));
                }
            }

            // Set selection based on what we have
            if (linkedReferences.Any() && !regularIds.Any())
            {
                // Only linked elements - use SetReferences
                uiDoc.Selection.SetReferences(linkedReferences);
            }
            else if (!linkedReferences.Any() && regularIds.Any())
            {
                // Only regular elements - use SetElementIds
                uiDoc.Selection.SetElementIds(regularIds);
            }
            else if (linkedReferences.Any() && regularIds.Any())
            {
                // Mixed selection - Revit API doesn't support mixing SetElementIds with SetReferences
                // We need to choose one approach. Let's prioritize what we have more of.
                if (linkedReferences.Count >= regularIds.Count)
                {
                    // More linked elements - convert regular elements to references if possible
                    var allReferences = new List<Reference>(linkedReferences);
                    
                    // Try to add regular elements as references
                    foreach (var elemId in regularIds)
                    {
                        try
                        {
                            Element elem = cData.Application.ActiveUIDocument.Document.GetElement(elemId);
                            if (elem != null)
                            {
                                Reference elemRef = new Reference(elem);
                                if (elemRef != null)
                                {
                                    allReferences.Add(elemRef);
                                }
                            }
                        }
                        catch { }
                    }
                    
                    uiDoc.Selection.SetReferences(allReferences);
                }
                else
                {
                    // More regular elements - just select those
                    uiDoc.Selection.SetElementIds(regularIds);
                    
                    TaskDialog.Show("Mixed Selection",
                        $"Selected {regularIds.Count} regular elements.\n" +
                        $"Note: {linkedReferences.Count} linked elements could not be included in the selection " +
                        "because Revit doesn't support mixing regular and linked selections.");
                }
            }

            return Result.Succeeded;
        }
        catch (InvalidOperationException ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
        catch (Exception ex)
        {
            message = $"Unexpected error: {ex.Message}";
            return Result.Failed;
        }
    }
}

#region Concrete commands

[Transaction(TransactionMode.Manual)]
public class FilterSelectedElements : FilterElementsBase
{
    public override bool SpanAllScreens      => false;
    public override bool UseSelectedElements => true;
    public override bool IncludeParameters   => true;
}

[Transaction(TransactionMode.Manual)]
public class FilterSelectedElementsSpanAllScreens : FilterElementsBase
{
    public override bool SpanAllScreens      => true;
    public override bool UseSelectedElements => true;
    public override bool IncludeParameters   => true;
}

#endregion
