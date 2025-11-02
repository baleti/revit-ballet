using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

[Transaction(TransactionMode.Manual)]
public class AddTextNoteToSelectedElements : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;
        
        // Get the selected elements
        var selectedIds = uiDoc.GetSelectionIds();
        var selectedRefs = uiDoc.GetReferences();
        
        if (!selectedIds.Any() && !selectedRefs.Any())
        {
            TaskDialog.Show("Error", "Please select elements to add text notes.");
            return Result.Failed;
        }
        
        try
        {
            // Get all unique parameters from selected elements
            var allParameters = new HashSet<string>();
            var elementsList = new List<Element>();
            
            // Process regular elements
            foreach (var id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element != null)
                {
                    elementsList.Add(element);
                    foreach (Parameter param in element.Parameters)
                    {
                        if (param.Definition != null && !string.IsNullOrEmpty(param.Definition.Name))
                        {
                            allParameters.Add(param.Definition.Name);
                        }
                    }
                }
            }
            
            // Process linked elements
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
                            var linkedElementId = reference.LinkedElementId;
                            if (linkedElementId != ElementId.InvalidElementId)
                            {
                                Element linkedElement = linkedDoc.GetElement(linkedElementId);
                                if (linkedElement != null)
                                {
                                    elementsList.Add(linkedElement);
                                    foreach (Parameter param in linkedElement.Parameters)
                                    {
                                        if (param.Definition != null && !string.IsNullOrEmpty(param.Definition.Name))
                                        {
                                            allParameters.Add(param.Definition.Name);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Element element = doc.GetElement(reference);
                        if (element != null && !elementsList.Any(e => e.Id == element.Id))
                        {
                            elementsList.Add(element);
                            foreach (Parameter param in element.Parameters)
                            {
                                if (param.Definition != null && !string.IsNullOrEmpty(param.Definition.Name))
                                {
                                    allParameters.Add(param.Definition.Name);
                                }
                            }
                        }
                    }
                }
                catch { }
            }
            
            if (!elementsList.Any())
            {
                TaskDialog.Show("Error", "No valid elements found.");
                return Result.Failed;
            }
            
            // Prepare parameter data for grid selection
            var parameterData = allParameters.OrderBy(p => p)
                .Select(p => new Dictionary<string, object> 
                { 
                    ["Parameter Name"] = p,
                    ["Selected"] = false
                })
                .ToList();
            
            if (!parameterData.Any())
            {
                TaskDialog.Show("Error", "No parameters found in selected elements.");
                return Result.Failed;
            }
            
            // Show parameter selection dialog
            var chosenParameters = CustomGUIs.DataGrid(
                parameterData, 
                new List<string> { "Parameter Name" }, 
                false);
            
            if (!chosenParameters.Any())
            {
                return Result.Cancelled;
            }
            
            // Extract selected parameter names
            var selectedParameterNames = chosenParameters
                .Select(dict => dict["Parameter Name"] as string)
                .Where(name => !string.IsNullOrEmpty(name))
                .ToList();
            
            // Get the last used text note type
            ElementId textNoteTypeId = GetLastUsedTextNoteType(doc);
            if (textNoteTypeId == null || textNoteTypeId == ElementId.InvalidElementId)
            {
                // If no last used, get the first available text note type
                var textNoteType = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .FirstElement() as TextNoteType;
                
                if (textNoteType == null)
                {
                    TaskDialog.Show("Error", "No text note types found in the document.");
                    return Result.Failed;
                }
                textNoteTypeId = textNoteType.Id;
            }
            
            // Generate UUID for this batch of text notes
            string uuid = Guid.NewGuid().ToString();
            string commentPrefix = $"temporary note {uuid}";
            
            using (Transaction trans = new Transaction(doc, "Add Text Notes to Elements"))
            {
                trans.Start();
                
                try
                {
                    View activeView = doc.ActiveView;
                    int viewScale = activeView.Scale;
                    
                    foreach (Element element in elementsList)
                    {
                        // Build text content from selected parameters
                        var textContent = BuildTextContent(element, selectedParameterNames);
                        if (string.IsNullOrEmpty(textContent))
                            continue;
                        
                        // Get element position
                        XYZ position = GetElementPosition(element);
                        if (position == null)
                            continue;
                        
                        // Calculate text note position with offset
                        BoundingBoxXYZ boundingBox = element.get_BoundingBox(null);
                        if (boundingBox != null)
                        {
                            double offsetY = 0.01 * viewScale; // Offset based on view scale
                            XYZ textPosition = new XYZ(position.X, boundingBox.Min.Y - offsetY, position.Z);
                            
                            // Create text note
                            TextNote textNote = TextNote.Create(
                                doc,
                                activeView.Id,
                                textPosition,
                                textContent,
                                textNoteTypeId);
                            
                            if (textNote != null)
                            {
                                // Set comment parameter
                                Parameter commentParam = textNote.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                                if (commentParam != null && !commentParam.IsReadOnly)
                                {
                                    commentParam.Set(commentPrefix);
                                }
                            }
                        }
                    }
                    
                    trans.Commit();
                    
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    trans.RollBack();
                    return Result.Failed;
                }
            }
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Unexpected error: {ex.Message}";
            return Result.Failed;
        }
    }
    
    private ElementId GetLastUsedTextNoteType(Document doc)
    {
        // Try to get from document settings or user preferences
        // This is a simplified approach - you might want to store this in extensible storage
        // For now, return InvalidElementId to use default behavior
        return ElementId.InvalidElementId;
    }
    
    private string BuildTextContent(Element element, List<string> parameterNames)
    {
        var sb = new StringBuilder();
        
        foreach (string paramName in parameterNames)
        {
            try
            {
                Parameter param = element.LookupParameter(paramName);
                if (param != null)
                {
                    string value = GetParameterValue(param);
                    if (!string.IsNullOrEmpty(value))
                    {
                        sb.AppendLine($"{paramName}: {value}");
                    }
                }
            }
            catch { }
        }
        
        return sb.ToString().TrimEnd();
    }
    
    private string GetParameterValue(Parameter param)
    {
        if (param == null)
            return string.Empty;
        
        try
        {
            // Try AsValueString first (works for most parameter types with units)
            string valueString = param.AsValueString();
            if (!string.IsNullOrEmpty(valueString))
                return valueString;
            
            // Try AsString for text parameters
            string stringValue = param.AsString();
            if (!string.IsNullOrEmpty(stringValue))
                return stringValue;
            
            // Handle specific storage types
            switch (param.StorageType)
            {
                case StorageType.Double:
                    double doubleValue = param.AsDouble();
                    return doubleValue.ToString("F2");
                    
                case StorageType.Integer:
                    int intValue = param.AsInteger();
                    return intValue.ToString();
                    
                case StorageType.ElementId:
                    ElementId idValue = param.AsElementId();
                    if (idValue != null && idValue != ElementId.InvalidElementId)
                    {
                        Element elem = param.Element.Document.GetElement(idValue);
                        return elem?.Name ?? idValue.IntegerValue.ToString();
                    }
                    break;
            }
            
            return "None";
        }
        catch
        {
            return "Error";
        }
    }
    
    private XYZ GetElementPosition(Element element)
    {
        // Try location point first
        LocationPoint locPoint = element.Location as LocationPoint;
        if (locPoint != null)
            return locPoint.Point;
        
        // Try location curve
        LocationCurve locCurve = element.Location as LocationCurve;
        if (locCurve != null && locCurve.Curve != null)
        {
            // Get midpoint of curve
            double param = (locCurve.Curve.GetEndParameter(0) + locCurve.Curve.GetEndParameter(1)) / 2;
            return locCurve.Curve.Evaluate(param, false);
        }
        
        // Try bounding box center
        BoundingBoxXYZ bb = element.get_BoundingBox(null);
        if (bb != null)
        {
            return (bb.Min + bb.Max) / 2;
        }
        
        return null;
    }
}
