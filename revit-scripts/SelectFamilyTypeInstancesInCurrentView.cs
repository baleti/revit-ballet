using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Architecture;

[Transaction(TransactionMode.Manual)]
public class SelectFamilyTypeInstancesInCurrentView : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
        Dictionary<string, Element> typeElementMap = new Dictionary<string, Element>();

        ElementId currentViewId = doc.ActiveView.Id;

        // Collect ALL elements in the current view without category restrictions
        var elementsInView = new FilteredElementCollector(doc, currentViewId)
            .WhereElementIsNotElementType()
            .Where(e => e.Category != null && IsElementVisibleInView(e, doc.ActiveView))
            .ToList();

        // Process different types of elements
        foreach (var element in elementsInView)
        {
            Element typeElement = null;
            string typeName = "";
            string familyName = "";
            
            if (element is FamilyInstance familyInstance)
            {
                typeElement = familyInstance.Symbol;
                typeName = familyInstance.Symbol.Name;
                familyName = familyInstance.Symbol.FamilyName;
            }
            else if (element is MEPCurve mepCurve) // Handles pipes, ducts, conduits, cable trays, etc.
            {
                typeElement = doc.GetElement(mepCurve.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    // Get more specific family name based on the MEPCurve type
                    if (element is Pipe)
                        familyName = "Pipe";
                    else if (element is Duct)
                        familyName = "Duct";
                    else if (element is Conduit)
                        familyName = "Conduit";
                    else if (element is CableTray)
                        familyName = "Cable Tray";
                    else
                        familyName = "MEP Curve";
                }
            }
            else if (element is Wall wall)
            {
                typeElement = doc.GetElement(wall.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    familyName = "Wall";
                }
            }
            else if (element is Floor floor)
            {
                typeElement = doc.GetElement(floor.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    familyName = "Floor";
                }
            }
            else if (element is RoofBase roof)
            {
                typeElement = doc.GetElement(roof.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    familyName = "Roof";
                }
            }
            else if (element is Ceiling ceiling)
            {
                typeElement = doc.GetElement(ceiling.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    familyName = "Ceiling";
                }
            }
            else if (element is Grid grid)
            {
                typeElement = doc.GetElement(grid.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    familyName = "Grid";
                }
            }
            else if (element is Level level)
            {
                // Levels don't have types in the same way
                typeName = level.Name;
                familyName = "Level";
                typeElement = element; // Use the element itself as reference
            }
            else if (element is ReferencePlane refPlane)
            {
                typeName = refPlane.Name;
                familyName = "Reference Plane";
                typeElement = element;
            }
            else if (element is Room room)
            {
                typeName = room.Name;
                familyName = "Room";
                typeElement = element;
            }
            else if (element is Area area)
            {
                typeName = area.Name;
                familyName = "Area";
                typeElement = element;
            }
            else if (element is Space space)
            {
                typeName = space.Name;
                familyName = "Space";
                typeElement = element;
            }
            else
            {
                // For any other element type, try to get its type
                ElementId typeId = element.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    typeElement = doc.GetElement(typeId);
                    if (typeElement != null)
                    {
                        typeName = typeElement.Name;
                        // Try to get a meaningful family name
                        if (typeElement is FamilySymbol fs)
                        {
                            familyName = fs.FamilyName;
                        }
                        else
                        {
                            familyName = element.GetType().Name;
                        }
                    }
                }
                else
                {
                    // For elements without types, use the element itself
                    typeName = element.Name;
                    familyName = element.GetType().Name;
                    typeElement = element;
                }
            }
            
            if (typeElement != null && !string.IsNullOrEmpty(typeName))
            {
                string uniqueKey = $"{familyName}:{typeName}:{element.Category.Name}";
                if (!typeElementMap.ContainsKey(uniqueKey))
                {
                    var entry = new Dictionary<string, object>
                    {
                        { "Type Name", typeName },
                        { "Family", familyName },
                        { "Category", element.Category.Name },
                        { "Count", 0 } // Will update this later
                    };

                    typeElementMap[uniqueKey] = typeElement;
                    typeEntries.Add(entry);
                }
            }
        }

        // Update counts for each type
        foreach (var element in elementsInView)
        {
            string uniqueKey = GetUniqueKeyForElement(element, doc);
            if (!string.IsNullOrEmpty(uniqueKey))
            {
                var entry = typeEntries.FirstOrDefault(e => 
                    $"{e["Family"]}:{e["Type Name"]}:{e["Category"]}" == uniqueKey);
                if (entry != null)
                {
                    entry["Count"] = (int)entry["Count"] + 1;
                }
            }
        }

        // Sort by Category, then Family, then Type Name
        typeEntries = typeEntries
            .OrderBy(e => e["Category"].ToString())
            .ThenBy(e => e["Family"].ToString())
            .ThenBy(e => e["Type Name"].ToString())
            .ToList();

        // Display the list of unique types with counts
        var propertyNames = new List<string> { "Category", "Family", "Type Name", "Count" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, spanAllScreens: false);

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled;
        }

        // Find all instances of selected types
        var selectedInstances = new List<ElementId>();
        
        foreach (var entry in selectedEntries)
        {
            string uniqueKey = $"{entry["Family"]}:{entry["Type Name"]}:{entry["Category"]}";
            
            var matchingElements = elementsInView
                .Where(e => GetUniqueKeyForElement(e, doc) == uniqueKey)
                .Select(e => e.Id)
                .ToList();
                
            selectedInstances.AddRange(matchingElements);
        }

        // Combine with existing selection
        ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();
        List<ElementId> combinedSelection = new List<ElementId>(currentSelection);

        foreach (var instanceId in selectedInstances)
        {
            if (!combinedSelection.Contains(instanceId))
            {
                combinedSelection.Add(instanceId);
            }
        }

        if (combinedSelection.Any())
        {
            uidoc.SetSelectionIds(combinedSelection);
            TaskDialog.Show("Selection", $"Selected {selectedInstances.Count} instance(s) of {selectedEntries.Count} type(s).");
        }
        else
        {
            TaskDialog.Show("Selection", "No visible instances of the selected types were found in the current view.");
        }

        return Result.Succeeded;
    }

    private string GetUniqueKeyForElement(Element element, Document doc)
    {
        string typeName = "";
        string familyName = "";
        string categoryName = element.Category?.Name ?? "Unknown";

        if (element is FamilyInstance familyInstance)
        {
            typeName = familyInstance.Symbol.Name;
            familyName = familyInstance.Symbol.FamilyName;
        }
        else if (element is MEPCurve mepCurve)
        {
            var typeElement = doc.GetElement(mepCurve.GetTypeId());
            if (typeElement != null)
            {
                typeName = typeElement.Name;
                if (element is Pipe)
                    familyName = "Pipe";
                else if (element is Duct)
                    familyName = "Duct";
                else if (element is Conduit)
                    familyName = "Conduit";
                else if (element is CableTray)
                    familyName = "Cable Tray";
                else
                    familyName = "MEP Curve";
            }
        }
        else if (element is Wall wall)
        {
            var typeElement = doc.GetElement(wall.GetTypeId());
            if (typeElement != null)
            {
                typeName = typeElement.Name;
                familyName = "Wall";
            }
        }
        else if (element is Floor floor)
        {
            var typeElement = doc.GetElement(floor.GetTypeId());
            if (typeElement != null)
            {
                typeName = typeElement.Name;
                familyName = "Floor";
            }
        }
        else if (element is RoofBase roof)
        {
            var typeElement = doc.GetElement(roof.GetTypeId());
            if (typeElement != null)
            {
                typeName = typeElement.Name;
                familyName = "Roof";
            }
        }
        else if (element is Ceiling ceiling)
        {
            var typeElement = doc.GetElement(ceiling.GetTypeId());
            if (typeElement != null)
            {
                typeName = typeElement.Name;
                familyName = "Ceiling";
            }
        }
        else if (element is Grid grid)
        {
            var typeElement = doc.GetElement(grid.GetTypeId());
            if (typeElement != null)
            {
                typeName = typeElement.Name;
                familyName = "Grid";
            }
        }
        else if (element is Level level)
        {
            typeName = level.Name;
            familyName = "Level";
        }
        else if (element is ReferencePlane refPlane)
        {
            typeName = refPlane.Name;
            familyName = "Reference Plane";
        }
        else if (element is Room room)
        {
            typeName = room.Name;
            familyName = "Room";
        }
        else if (element is Area area)
        {
            typeName = area.Name;
            familyName = "Area";
        }
        else if (element is Space space)
        {
            typeName = space.Name;
            familyName = "Space";
        }
        else
        {
            ElementId typeId = element.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                var typeElement = doc.GetElement(typeId);
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    if (typeElement is FamilySymbol fs)
                    {
                        familyName = fs.FamilyName;
                    }
                    else
                    {
                        familyName = element.GetType().Name;
                    }
                }
            }
            else
            {
                typeName = element.Name;
                familyName = element.GetType().Name;
            }
        }

        if (!string.IsNullOrEmpty(typeName))
        {
            return $"{familyName}:{typeName}:{categoryName}";
        }

        return "";
    }

    private bool IsElementVisibleInView(Element element, View view)
    {
        // Check if element has a bounding box in the view
        BoundingBoxXYZ bbox = element.get_BoundingBox(view);
        
        // Additional check for elements that might not have a bounding box
        // but are still visible (like some annotation elements)
        if (bbox == null && element.IsHidden(view))
        {
            return false;
        }
        
        return bbox != null || !element.IsHidden(view);
    }
}
