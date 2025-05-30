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
            
            // Process regular elements from current document
            foreach (var id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element != null)
                {
                    var data = GetElementDataDictionary(element, doc, null, null, includeParameters);
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
                            var linkedRef = reference.LinkedElementId;
                            if (linkedRef != ElementId.InvalidElementId)
                            {
                                Element linkedElement = linkedDoc.GetElement(linkedRef);
                                if (linkedElement != null)
                                {
                                    var data = GetElementDataDictionary(linkedElement, linkedDoc, linkedInstance.Name, reference, includeParameters);
                                    elementData.Add(data);
                                }
                            }
                        }
                    }
                    else
                    {
                        // This might be a regular element selected via reference
                        Element element = doc.GetElement(reference);
                        if (element != null)
                        {
                            var data = GetElementDataDictionary(element, doc, null, null, includeParameters);
                            elementData.Add(data);
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
                    var data = GetElementDataDictionary(element, doc, null, null, includeParameters);
                    elementData.Add(data);
                }
            }
        }
        
        return elementData;
    }
    
    private static Dictionary<string, object> GetElementDataDictionary(Element element, Document elementDoc, string linkName, Reference linkReference, bool includeParameters = false)
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
            ["_ElementId"] = element.Id, // Store full ElementId for selection
            ["_LinkReference"] = linkReference // Store reference for linked elements
        };
        
        // Use just the element name for DisplayName, regardless of whether it's linked
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
            
            // Get property names, but exclude internal fields starting with _
            var propertyNames = elementData.First().Keys
                .Where(k => !k.StartsWith("_"))
                .ToList();
            
            // Reorder to put most useful columns first
            var orderedProps = new List<string> { "DisplayName", "Category", "LinkName", "Group", "OwnerView", "Id" };
            var remainingProps = propertyNames.Except(orderedProps).OrderBy(p => p);
            propertyNames = orderedProps.Where(p => propertyNames.Contains(p))
                .Concat(remainingProps)
                .ToList();
            
            var chosenRows = CustomGUIs.DataGrid(elementData, propertyNames, SpanAllScreens);
            if (chosenRows.Count == 0)
                return Result.Cancelled;
            
            // Separate regular elements and linked elements
            var regularIds = new List<ElementId>();
            var linkedReferences = new List<Reference>();
            
            foreach (var row in chosenRows)
            {
                if (row.TryGetValue("_LinkReference", out var refObj) && refObj is Reference linkRef)
                {
                    // This is a linked element
                    linkedReferences.Add(linkRef);
                }
                else if (row.TryGetValue("_ElementId", out var idObj) && idObj is ElementId elemId)
                {
                    // This is a regular element
                    regularIds.Add(elemId);
                }
            }
            
            // Clear current selection
            uiDoc.Selection.SetElementIds(new List<ElementId>());
            
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
                // Mixed selection - need to handle this carefully
                // First set regular elements
                uiDoc.Selection.SetElementIds(regularIds);
                
                // Then try to add linked references
                // Note: Revit may not support mixed selection well
                TaskDialog.Show("Mixed Selection", 
                    $"Selected {regularIds.Count} regular elements and {linkedReferences.Count} linked elements.\n" +
                    "Note: Revit may only show one type in the selection.");
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
