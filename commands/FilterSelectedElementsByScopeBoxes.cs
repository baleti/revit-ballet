using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// Command to filter selected elements showing scope boxes from both project and linked models.
/// Shows only DisplayName, ElementId, Level, and scope box columns.
/// Only displays scope box columns that have at least one intersection.
/// Uses linked model Type names for column headers.
/// </summary>
[Transaction(TransactionMode.ReadOnly)]
[Regeneration(RegenerationOption.Manual)]
public class FilterSelectedElementsByScopeBoxes : IExternalCommand
{
    // Tolerance for bounding box intersections (0.5 feet = 6 inches)
    private const double BoundingBoxTolerance = 0.5;
    
    public Result Execute(ExternalCommandData cData, ref string message, ElementSet elements)
    {
        try
        {
            var uiDoc = cData.Application.ActiveUIDocument;
            var doc = uiDoc.Document;
            
            // Get selected elements
            var selectedIds = uiDoc.GetSelectionIds();
            var selectedRefs = uiDoc.GetReferences();

            if (!selectedIds.Any() && !selectedRefs.Any())
            {
                TaskDialog.Show("Info", "No elements are selected.");
                return Result.Cancelled;
            }
            
            // Get scope boxes from main project
            var projectScopeBoxes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .Where(e => e.Name != null)
                .ToList();
            
            // Get all linked documents and their scope boxes
            // Store: Link instance -> (Type name, scope boxes)
            var linkedScopeBoxes = new Dictionary<RevitLinkInstance, (string TypeName, List<Element> ScopeBoxes)>();
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();
            
            foreach (var linkInstance in linkInstances)
            {
                Document linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc != null)
                {
                    var linkScopeBoxes = new FilteredElementCollector(linkedDoc)
                        .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                        .WhereElementIsNotElementType()
                        .Cast<Element>()
                        .Where(e => e.Name != null)
                        .ToList();
                    
                    if (linkScopeBoxes.Any())
                    {
                        // Get the Type name for better readability
                        string typeName = linkInstance.Name;
                        var linkType = doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
                        if (linkType != null && !string.IsNullOrEmpty(linkType.Name))
                        {
                            typeName = linkType.Name;
                            // Remove .rvt extension if present for cleaner display
                            if (typeName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                            {
                                typeName = typeName.Substring(0, typeName.Length - 4);
                            }
                        }
                        
                        linkedScopeBoxes[linkInstance] = (typeName, linkScopeBoxes);
                    }
                }
            }
            
            // Process selected elements
            var elementData = new List<Dictionary<string, object>>();
            var processedElements = new HashSet<ElementId>();
            
            // Track which scope box columns have any intersections
            bool hasProjectScopeBoxIntersections = false;
            var linkedModelsWithIntersections = new HashSet<string>();
            
            // Process regular elements from current document
            foreach (var id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element != null)
                {
                    processedElements.Add(id);
                    var data = GetElementDataWithScopeBoxes(element, doc, null, null, null, 
                        projectScopeBoxes, linkedScopeBoxes, linkInstances);
                    elementData.Add(data);
                    
                    // Check if this element has any scope box intersections
                    UpdateIntersectionTracking(data, ref hasProjectScopeBoxIntersections, linkedModelsWithIntersections);
                }
            }
            
            // Process linked elements via References
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
                                    var data = GetElementDataWithScopeBoxes(linkedElement, linkedDoc, 
                                        linkedInstance.Name, linkedInstance, linkedElement.Id,
                                        projectScopeBoxes, linkedScopeBoxes, linkInstances);
                                    elementData.Add(data);
                                    
                                    // Check if this element has any scope box intersections
                                    UpdateIntersectionTracking(data, ref hasProjectScopeBoxIntersections, linkedModelsWithIntersections);
                                }
                            }
                        }
                    }
                    else if (!processedElements.Contains(reference.ElementId))
                    {
                        Element element = doc.GetElement(reference);
                        if (element != null)
                        {
                            processedElements.Add(reference.ElementId);
                            var data = GetElementDataWithScopeBoxes(element, doc, null, null, null,
                                projectScopeBoxes, linkedScopeBoxes, linkInstances);
                            elementData.Add(data);
                            
                            // Check if this element has any scope box intersections
                            UpdateIntersectionTracking(data, ref hasProjectScopeBoxIntersections, linkedModelsWithIntersections);
                        }
                    }
                }
                catch { /* Skip problematic references */ }
            }
            
            if (!elementData.Any())
            {
                TaskDialog.Show("Info", "No valid elements found.");
                return Result.Cancelled;
            }
            
            // Create unique keys and prepare display data
            var elementDataMap = new Dictionary<string, Dictionary<string, object>>();
            var displayData = new List<Dictionary<string, object>>();
            
            for (int i = 0; i < elementData.Count; i++)
            {
                var data = elementData[i];
                string uniqueKey = $"{data["Id"]}_{data["Type"]}_{data["LinkName"]}_{i}";
                
                elementDataMap[uniqueKey] = data;
                
                var display = new Dictionary<string, object>(data);
                display["UniqueKey"] = uniqueKey;
                displayData.Add(display);
            }
            
            // Define columns to show - Type, Id, Level, and only scope box columns with intersections
            var propertyNames = new List<string> { "Type", "Id", "Level" };
            
            // Add project scope box column only if there are intersections
            if (hasProjectScopeBoxIntersections)
            {
                propertyNames.Add("Project");
            }
            
            // Add columns for each linked model that has scope box intersections
            // Sort by type name for consistent ordering
            var sortedLinkedModels = linkedModelsWithIntersections.OrderBy(k => k).ToList();
            foreach (var typeName in sortedLinkedModels)
            {
                propertyNames.Add(typeName);
            }
            
            // Filter to only include existing columns
            var existingProps = displayData.First().Keys;
            propertyNames = propertyNames.Where(p => existingProps.Contains(p)).ToList();
            
            // Show data grid
            var chosenRows = CustomGUIs.DataGrid(displayData, propertyNames, false);
            if (chosenRows.Count == 0)
                return Result.Cancelled;
            
            // Handle selection (same as before)
            var regularIds = new List<ElementId>();
            var linkedReferences = new List<Reference>();
            
            foreach (var row in chosenRows)
            {
                if (!row.TryGetValue("UniqueKey", out var keyObj) || !(keyObj is string uniqueKey))
                    continue;
                
                if (!elementDataMap.TryGetValue(uniqueKey, out var fullData))
                    continue;
                
                if (fullData.TryGetValue("LinkInstanceObject", out var linkObj) && linkObj is RevitLinkInstance linkInstance &&
                    fullData.TryGetValue("LinkedElementIdObject", out var linkedIdObj) && linkedIdObj is ElementId linkedElementId)
                {
                    try
                    {
                        Document linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            Element linkedElement = linkedDoc.GetElement(linkedElementId);
                            if (linkedElement != null)
                            {
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
                    regularIds.Add(elemId);
                }
                else if (fullData.TryGetValue("Id", out var intId) && intId is int id)
                {
                    regularIds.Add(new ElementId(id));
                }
            }
            
            // Set selection
            if (linkedReferences.Any() && !regularIds.Any())
            {
                uiDoc.SetReferences(linkedReferences);
            }
            else if (!linkedReferences.Any() && regularIds.Any())
            {
                uiDoc.SetSelectionIds(regularIds);
            }
            else if (linkedReferences.Any() && regularIds.Any())
            {
                if (linkedReferences.Count >= regularIds.Count)
                {
                    var allReferences = new List<Reference>(linkedReferences);
                    foreach (var elemId in regularIds)
                    {
                        try
                        {
                            Element elem = doc.GetElement(elemId);
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
                    uiDoc.SetReferences(allReferences);
                }
                else
                {
                    uiDoc.SetSelectionIds(regularIds);
                    TaskDialog.Show("Mixed Selection",
                        $"Selected {regularIds.Count} regular elements.\n" +
                        $"Note: {linkedReferences.Count} linked elements could not be included in the selection.");
                }
            }
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Error: {ex.Message}";
            return Result.Failed;
        }
    }
    
    private void UpdateIntersectionTracking(Dictionary<string, object> data, 
        ref bool hasProjectScopeBoxIntersections, 
        HashSet<string> linkedModelsWithIntersections)
    {
        // Check project scope boxes
        if (data.ContainsKey("Project") && 
            data["Project"] is string projectSB && 
            !string.IsNullOrEmpty(projectSB))
        {
            hasProjectScopeBoxIntersections = true;
        }
        
        // Check linked model scope boxes
        foreach (var kvp in data)
        {
            if (kvp.Key != "Project" &&
                kvp.Key != "Type" &&
                kvp.Key != "Id" &&
                kvp.Key != "Level" &&
                kvp.Key != "LinkName" &&
                kvp.Key != "ElementIdObject" &&
                kvp.Key != "LinkInstanceObject" &&
                kvp.Key != "LinkedElementIdObject" &&
                kvp.Value is string linkedSB && 
                !string.IsNullOrEmpty(linkedSB))
            {
                linkedModelsWithIntersections.Add(kvp.Key);
            }
        }
    }
    
    private Dictionary<string, object> GetElementDataWithScopeBoxes(
        Element element, 
        Document elementDoc, 
        string linkName,
        RevitLinkInstance linkInstance, 
        ElementId linkedElementId,
        List<Element> projectScopeBoxes,
        Dictionary<RevitLinkInstance, (string TypeName, List<Element> ScopeBoxes)> linkedScopeBoxes,
        List<RevitLinkInstance> linkInstances)
    {
        var data = new Dictionary<string, object>
        {
            ["Type"] = element.Name ?? string.Empty,
            ["Id"] = element.Id.IntegerValue,
            ["LinkName"] = linkName ?? string.Empty,
            ["ElementIdObject"] = element.Id,
            ["LinkInstanceObject"] = linkInstance,
            ["LinkedElementIdObject"] = linkedElementId
        };
        
        // Get Level
        string levelName = string.Empty;
        if (element.LevelId != null && element.LevelId != ElementId.InvalidElementId)
        {
            if (elementDoc.GetElement(element.LevelId) is Level level)
                levelName = level.Name;
        }
        data["Level"] = levelName;
        
        // Get element's bounding box
        BoundingBoxXYZ elementBB = null;
        try
        {
            elementBB = element.get_BoundingBox(null);
            if (elementBB == null)
            {
                var options = new Options();
                var geom = element.get_Geometry(options);
                if (geom != null)
                {
                    elementBB = geom.GetBoundingBox();
                }
            }
            
            if (elementBB != null && linkInstance != null)
            {
                // Transform for linked element
                Transform transform = linkInstance.GetTotalTransform();
                elementBB = TransformBoundingBox(elementBB, transform);
            }
        }
        catch { }
        
        if (elementBB != null)
        {
            // Check project scope boxes
            var projectSBList = new List<string>();
            foreach (var scopeBox in projectScopeBoxes)
            {
                BoundingBoxXYZ scopeBB = scopeBox.get_BoundingBox(null);
                if (scopeBB != null && DoesBoundingBoxIntersect(elementBB, scopeBB))
                {
                    projectSBList.Add(scopeBox.Name);
                }
            }
            data["Project"] = string.Join(", ", projectSBList.OrderBy(s => s));
            
            // Check linked model scope boxes
            foreach (var kvp in linkedScopeBoxes)
            {
                var thisLinkInstance = kvp.Key;
                var typeName = kvp.Value.TypeName;
                var linkScopeBoxes = kvp.Value.ScopeBoxes;
                var linkedSBList = new List<string>();
                
                Transform linkTransform = thisLinkInstance.GetTotalTransform();
                
                foreach (var scopeBox in linkScopeBoxes)
                {
                    BoundingBoxXYZ scopeBB = scopeBox.get_BoundingBox(null);
                    if (scopeBB != null)
                    {
                        // Transform the linked scope box to project coordinates
                        scopeBB = TransformBoundingBox(scopeBB, linkTransform);
                        
                        if (DoesBoundingBoxIntersect(elementBB, scopeBB))
                        {
                            linkedSBList.Add(scopeBox.Name);
                        }
                    }
                }
                
                data[typeName] = string.Join(", ", linkedSBList.OrderBy(s => s));
            }
        }
        else
        {
            // No bounding box - set empty values
            data["Project"] = string.Empty;
            foreach (var kvp in linkedScopeBoxes)
            {
                var typeName = kvp.Value.TypeName;
                data[typeName] = string.Empty;
            }
        }
        
        return data;
    }
    
    private BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bb, Transform transform)
    {
        XYZ min = bb.Min;
        XYZ max = bb.Max;
        
        var corners = new[]
        {
            new XYZ(min.X, min.Y, min.Z),
            new XYZ(max.X, min.Y, min.Z),
            new XYZ(min.X, max.Y, min.Z),
            new XYZ(max.X, max.Y, min.Z),
            new XYZ(min.X, min.Y, max.Z),
            new XYZ(max.X, min.Y, max.Z),
            new XYZ(min.X, max.Y, max.Z),
            new XYZ(max.X, max.Y, max.Z)
        };
        
        var transformedCorners = corners.Select(c => transform.OfPoint(c)).ToArray();
        
        double minX = transformedCorners.Min(p => p.X);
        double minY = transformedCorners.Min(p => p.Y);
        double minZ = transformedCorners.Min(p => p.Z);
        double maxX = transformedCorners.Max(p => p.X);
        double maxY = transformedCorners.Max(p => p.Y);
        double maxZ = transformedCorners.Max(p => p.Z);
        
        var result = new BoundingBoxXYZ();
        result.Min = new XYZ(minX, minY, minZ);
        result.Max = new XYZ(maxX, maxY, maxZ);
        
        return result;
    }
    
    private bool DoesBoundingBoxIntersect(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        double tolerance = BoundingBoxTolerance;
        
        if (bb1.Max.X + tolerance < bb2.Min.X || bb1.Min.X - tolerance > bb2.Max.X) return false;
        if (bb1.Max.Y + tolerance < bb2.Min.Y || bb1.Min.Y - tolerance > bb2.Max.Y) return false;
        if (bb1.Max.Z + tolerance < bb2.Min.Z || bb1.Min.Z - tolerance > bb2.Max.Z) return false;
        
        return true;
    }
}
