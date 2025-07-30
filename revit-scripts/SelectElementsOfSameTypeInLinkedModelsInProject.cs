using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectElementsOfSameTypeInLinkedModelsInProject : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Document doc = uiapp.ActiveUIDocument.Document;
        UIDocument uiDoc = uiapp.ActiveUIDocument;
        
        // Get currently selected elements using SelectionModeManager methods
        var selectedIds = uiDoc.GetSelectionIds();
        var selectedRefs = uiDoc.GetReferences();
        
        if (!selectedIds.Any() && !selectedRefs.Any())
        {
            TaskDialog.Show("Info", "No elements are selected.");
            return Result.Cancelled;
        }
        
        // Collect type information from selected elements
        var selectedTypes = new HashSet<string>(); // Use type signatures for comparison
        var typeDescriptions = new List<string>(); // For user feedback
        var selectedLinkIds = new HashSet<ElementId>(); // Track link instance IDs that contain selected elements
        
        // Process regular elements from host document
        foreach (var id in selectedIds)
        {
            Element element = doc.GetElement(id);
            if (element != null && !(element is RevitLinkInstance))
            {
                string typeSignature = GetElementTypeSignature(element);
                if (!string.IsNullOrEmpty(typeSignature))
                {
                    selectedTypes.Add(typeSignature);
                    typeDescriptions.Add(GetElementTypeDescription(element));
                }
            }
        }
        
        // Process linked elements from references
        foreach (var reference in selectedRefs)
        {
            try
            {
                if (reference.LinkedElementId != ElementId.InvalidElementId)
                {
                    // This is a linked element reference
                    var linkedInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkedInstance != null)
                    {
                        // Track this link instance
                        selectedLinkIds.Add(reference.ElementId);
                        
                        Document linkedDoc = linkedInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            Element linkedElement = linkedDoc.GetElement(reference.LinkedElementId);
                            if (linkedElement != null)
                            {
                                string typeSignature = GetElementTypeSignature(linkedElement);
                                if (!string.IsNullOrEmpty(typeSignature))
                                {
                                    selectedTypes.Add(typeSignature);
                                    typeDescriptions.Add(GetElementTypeDescription(linkedElement));
                                }
                            }
                        }
                    }
                }
                else if (reference.ElementId != ElementId.InvalidElementId)
                {
                    // Regular element reference
                    Element element = doc.GetElement(reference.ElementId);
                    if (element != null && !(element is RevitLinkInstance))
                    {
                        string typeSignature = GetElementTypeSignature(element);
                        if (!string.IsNullOrEmpty(typeSignature))
                        {
                            selectedTypes.Add(typeSignature);
                            typeDescriptions.Add(GetElementTypeDescription(element));
                        }
                    }
                }
            }
            catch { /* Skip problematic references */ }
        }
        
        if (!selectedTypes.Any())
        {
            TaskDialog.Show("Info", "No valid element types found in selection.");
            return Result.Cancelled;
        }
        
        // Get all loaded linked models
        var allLinkInstances = new FilteredElementCollector(doc)
            .OfClass(typeof(RevitLinkInstance))
            .Cast<RevitLinkInstance>()
            .Where(link => link.GetLinkDocument() != null)
            .ToList();
        
        if (!allLinkInstances.Any())
        {
            TaskDialog.Show("Info", "No loaded linked models found in the project.");
            return Result.Cancelled;
        }
        
        // Create dictionary entries for the DataGrid
        var linkEntries = new List<Dictionary<string, object>>();
        var linkToIndexMap = new Dictionary<ElementId, int>(); // Map link ElementId to index
        
        // Sort link instances by name first
        allLinkInstances = allLinkInstances.OrderBy(link => link.Name).ToList();
        
        for (int i = 0; i < allLinkInstances.Count; i++)
        {
            var link = allLinkInstances[i];
            var entry = new Dictionary<string, object>
            {
                { "Name", link.Name },
                { "LinkType", link.GetTypeId() != ElementId.InvalidElementId ? 
                    doc.GetElement(link.GetTypeId())?.Name ?? "Unknown" : "Unknown" }
            };
            linkEntries.Add(entry);
            linkToIndexMap[link.Id] = i;
        }
        
        // Determine which models should be pre-selected based on selectedLinkIds
        var preSelectedIndices = new List<int>();
        foreach (var linkId in selectedLinkIds)
        {
            if (linkToIndexMap.ContainsKey(linkId))
            {
                preSelectedIndices.Add(linkToIndexMap[linkId]);
            }
        }
        
        // Define properties to display
        var linkPropertyNames = new List<string> { "Name", "LinkType" };
        
        // Show the DataGrid to let the user select linked models to search
        var selectedEntries = CustomGUIs.DataGrid(
            linkEntries,
            linkPropertyNames,
            false, // spanAllScreens
            preSelectedIndices.Count > 0 ? preSelectedIndices : null
        );
        
        if (selectedEntries == null || selectedEntries.Count == 0)
            return Result.Cancelled;
        
        // Extract the selected link instances based on selected entries
        var selectedLinkInstances = new List<RevitLinkInstance>();
        foreach (var entry in selectedEntries)
        {
            var linkName = entry["Name"].ToString();
            var matchingLink = allLinkInstances.FirstOrDefault(link => link.Name == linkName);
            if (matchingLink != null)
            {
                selectedLinkInstances.Add(matchingLink);
            }
        }
        
        List<Reference> matchingReferences = new List<Reference>();
        int linksSearched = 0;
        int totalElementsFound = 0;
        
        // Search only the user-selected linked models for matching element types
        foreach (var linkInstance in selectedLinkInstances)
        {
            Document linkedDoc = linkInstance.GetLinkDocument();
            if (linkedDoc == null) continue;
            
            linksSearched++;
            int elementsInThisLink = 0;
            
            FilteredElementCollector collector = new FilteredElementCollector(linkedDoc);
            collector.WhereElementIsNotElementType();
            
            foreach (Element elem in collector)
            {
                try
                {
                    string elemTypeSignature = GetElementTypeSignature(elem);
                    if (selectedTypes.Contains(elemTypeSignature))
                    {
                        Reference elemRef = new Reference(elem);
                        Reference linkedRef = elemRef.CreateLinkReference(linkInstance);
                        
                        if (linkedRef != null)
                        {
                            matchingReferences.Add(linkedRef);
                            elementsInThisLink++;
                            totalElementsFound++;
                        }
                    }
                }
                catch
                {
                    // Some elements might not support reference creation
                    continue;
                }
            }
        }
        
        if (matchingReferences.Count > 0)
        {
            // Select all matching elements using SelectionModeManager
            uiDoc.SetReferences(matchingReferences);
        }
        
        return Result.Succeeded;
    }
    
    /// <summary>
    /// Creates a unique signature for an element type that can be used for cross-document comparison
    /// </summary>
    private string GetElementTypeSignature(Element element)
    {
        try
        {
            // Get the element's type
            ElementId typeId = element.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                ElementType elementType = element.Document.GetElement(typeId) as ElementType;
                if (elementType != null)
                {
                    // For loadable families, use family name + type name
                    if (!string.IsNullOrEmpty(elementType.FamilyName))
                    {
                        return $"Family:{elementType.FamilyName}:{elementType.Name}";
                    }
                    else
                    {
                        // For system families, use category + type name
                        string categoryName = element.Category?.Name ?? "Unknown";
                        return $"System:{categoryName}:{elementType.Name}";
                    }
                }
            }
            
            // Fallback for elements without types (like some annotation elements)
            if (element.Category != null)
            {
                // Use category + some identifying parameters
                string categoryName = element.Category.Name;
                
                // Try to get some distinguishing parameters
                string additionalInfo = "";
                
                // For text elements, include text style
                if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_TextNotes)
                {
                    var styleParam = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                    if (styleParam != null)
                    {
                        var textStyle = element.Document.GetElement(styleParam.AsElementId()) as TextNoteType;
                        additionalInfo = textStyle?.Name ?? "";
                    }
                }
                // For dimensions, include dimension style
                else if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Dimensions)
                {
                    var styleParam = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                    if (styleParam != null)
                    {
                        var dimStyle = element.Document.GetElement(styleParam.AsElementId()) as DimensionType;
                        additionalInfo = dimStyle?.Name ?? "";
                    }
                }
                
                return $"NoType:{categoryName}:{additionalInfo}:{element.GetType().Name}";
            }
            
            // Last resort fallback
            return $"Unknown:{element.GetType().Name}";
        }
        catch
        {
            return string.Empty;
        }
    }
    
    /// <summary>
    /// Gets a human-readable description of the element type for user feedback
    /// </summary>
    private string GetElementTypeDescription(Element element)
    {
        try
        {
            ElementId typeId = element.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                ElementType elementType = element.Document.GetElement(typeId) as ElementType;
                if (elementType != null)
                {
                    if (!string.IsNullOrEmpty(elementType.FamilyName))
                    {
                        return $"{elementType.FamilyName}: {elementType.Name}";
                    }
                    else
                    {
                        return $"{element.Category?.Name ?? "Unknown"}: {elementType.Name}";
                    }
                }
            }
            
            // Fallback
            return element.Category?.Name ?? element.GetType().Name;
        }
        catch
        {
            return "Unknown Type";
        }
    }
}
