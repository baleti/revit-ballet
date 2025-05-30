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
        
        // Get currently selected elements (both regular and linked)
        var selectedIds = uiDoc.Selection.GetElementIds();
        var selectedRefs = uiDoc.Selection.GetReferences();
        
        if (!selectedIds.Any() && !selectedRefs.Any())
        {
            TaskDialog.Show("Info", "No elements are selected.");
            return Result.Cancelled;
        }
        
        // Collect type information from selected elements
        var selectedTypes = new HashSet<string>(); // Use type signatures for comparison
        var typeDescriptions = new List<string>(); // For user feedback
        
        // Process regular elements from host document
        foreach (var id in selectedIds)
        {
            Element element = doc.GetElement(id);
            if (element != null)
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
                var linkedInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                if (linkedInstance != null)
                {
                    Document linkedDoc = linkedInstance.GetLinkDocument();
                    if (linkedDoc != null)
                    {
                        var linkedRef = reference.LinkedElementId;
                        if (linkedRef != ElementId.InvalidElementId)
                        {
                            Element linkedElement = linkedDoc.GetElement(linkedRef);
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
                else
                {
                    // Handle regular element selected via reference
                    Element element = doc.GetElement(reference);
                    if (element != null)
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
        
        // Create wrappers for the DataGrid
        var linkWrappers = allLinkInstances
            .Select(link => new LinkedModelWrapper(link))
            .OrderBy(w => w.Name)
            .ToList();
        
        // Determine which models should be pre-selected (models that contain selected elements)
        var preSelectedIndices = new List<int>();
        var selectedLinkNames = new HashSet<string>();
        
        // Collect link names from currently selected linked elements
        foreach (var reference in selectedRefs)
        {
            try
            {
                var linkedInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                if (linkedInstance != null)
                {
                    selectedLinkNames.Add(linkedInstance.Name);
                }
            }
            catch { /* Skip problematic references */ }
        }
        
        // Find indices of models that should be pre-selected
        for (int i = 0; i < linkWrappers.Count; i++)
        {
            if (selectedLinkNames.Contains(linkWrappers[i].Name))
            {
                preSelectedIndices.Add(i);
            }
        }
        
        // Define properties to display
        var linkPropertyNames = new List<string> { "Name", "LinkType" };
        
        // Show the DataGrid to let the user select linked models to search
        var selectedLinkWrappers = CustomGUIs.DataGrid<LinkedModelWrapper>(
            linkWrappers, 
            linkPropertyNames, 
            preSelectedIndices.Count > 0 ? preSelectedIndices : null,
            "Select Linked Models to Search"
        );
        
        if (selectedLinkWrappers.Count == 0)
            return Result.Cancelled;
        
        // Extract the selected link instances
        var selectedLinkInstances = selectedLinkWrappers.Select(w => w.LinkInstance).ToList();
        
        List<Reference> matchingReferences = new List<Reference>();
        int linksSearched = 0;
        
        // Search only the user-selected linked models for matching element types
        foreach (var linkInstance in selectedLinkInstances)
        {
            Document linkedDoc = linkInstance.GetLinkDocument();
            if (linkedDoc == null) continue;
            
            linksSearched++;
            
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
            // Select all matching elements
            uiDoc.Selection.SetReferences(matchingReferences);
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
