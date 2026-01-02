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
public class SelectByFamilyTypesInViews : IExternalCommand
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

        // Check if any views are currently selected
        ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();
        List<View> targetViews = new List<View>();

        bool hasSelectedViews = currentSelection.Any(id =>
        {
            Element elem = doc.GetElement(id);
            return elem is View;
        });

        if (hasSelectedViews)
        {
            // Use selected views
            targetViews = currentSelection
                .Select(id => doc.GetElement(id))
                .OfType<View>()
                .ToList();
        }
        else
        {
            // Use current view only
            targetViews.Add(doc.ActiveView);
        }

        // Track which views each family type appears in
        Dictionary<string, HashSet<string>> typeToViewsMap = new Dictionary<string, HashSet<string>>();
        Dictionary<string, List<ElementId>> typeToElementsMap = new Dictionary<string, List<ElementId>>();

        // Collect ALL elements in the target views without category restrictions
        foreach (View view in targetViews)
        {
            var elementsInView = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && IsElementVisibleInView(e, view))
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
                        familyName = GetSystemFamilyName(typeElement) ?? "Wall";
                    }
                }
                else if (element is Floor floor)
                {
                    typeElement = doc.GetElement(floor.GetTypeId());
                    if (typeElement != null)
                    {
                        typeName = typeElement.Name;
                        familyName = GetSystemFamilyName(typeElement) ?? "Floor";
                    }
                }
                else if (element is RoofBase roof)
                {
                    typeElement = doc.GetElement(roof.GetTypeId());
                    if (typeElement != null)
                    {
                        typeName = typeElement.Name;
                        familyName = GetSystemFamilyName(typeElement) ?? "Roof";
                    }
                }
                else if (element is Ceiling ceiling)
                {
                    typeElement = doc.GetElement(ceiling.GetTypeId());
                    if (typeElement != null)
                    {
                        typeName = typeElement.Name;
                        familyName = GetSystemFamilyName(typeElement) ?? "Ceiling";
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
                                familyName = GetSystemFamilyName(typeElement) ?? element.GetType().Name;
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
                    // Filter out DWG import symbols
                    if (familyName.Contains("Import Symbol") ||
                        (element.Category != null && element.Category.Name.Contains("Import Symbol")))
                        continue;

                    string uniqueKey = $"{familyName}:{typeName}:{element.Category.Name}";

                    // Track which views this type appears in
                    if (!typeToViewsMap.ContainsKey(uniqueKey))
                    {
                        typeToViewsMap[uniqueKey] = new HashSet<string>();
                        typeToElementsMap[uniqueKey] = new List<ElementId>();
                    }
                    typeToViewsMap[uniqueKey].Add(view.Name);

                    // Store element ID for later selection
                    if (!typeToElementsMap[uniqueKey].Contains(element.Id))
                    {
                        typeToElementsMap[uniqueKey].Add(element.Id);
                    }

                    if (!typeElementMap.ContainsKey(uniqueKey))
                    {
                        typeElementMap[uniqueKey] = typeElement;
                    }
                }
            }
        }

        // Build the type entries with view information
        foreach (var kvp in typeElementMap)
        {
            string uniqueKey = kvp.Key;
            string[] parts = uniqueKey.Split(':');
            string familyName = parts[0];
            string typeName = parts[1];
            string categoryName = parts[2];

            int count = typeToElementsMap[uniqueKey].Count;
            string viewNames = string.Join(", ", typeToViewsMap[uniqueKey].OrderBy(v => v));

            var entry = new Dictionary<string, object>
            {
                { "Type Name", typeName },
                { "Family", familyName },
                { "Category", categoryName },
                { "Count", count },
                { "_UniqueKey", uniqueKey },  // Store the original key for lookup after edits
                { "ElementIdObject", kvp.Value.Id }  // Store ElementId for edit mode support
            };

            // Only add View column if multiple views are being queried
            if (hasSelectedViews)
            {
                entry.Add("View", viewNames);
            }

            typeEntries.Add(entry);
        }

        // Sort by Category, then Family, then Type Name
        typeEntries = typeEntries
            .OrderBy(e => e["Category"].ToString())
            .ThenBy(e => e["Family"].ToString())
            .ThenBy(e => e["Type Name"].ToString())
            .ToList();

        // Display the list of unique types with counts
        // Set UIDocument for edit mode support
        CustomGUIs.SetCurrentUIDocument(uidoc);

        var propertyNames = hasSelectedViews
            ? new List<string> { "Category", "Family", "Type Name", "Count", "View" }
            : new List<string> { "Category", "Family", "Type Name", "Count" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, false);

        // Apply any pending edits (family/type renames)
        if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
        {
            CustomGUIs.ApplyCellEditsToEntities();
        }

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled;
        }

        // Find all instances of selected types
        var selectedInstances = new List<ElementId>();

        foreach (var entry in selectedEntries)
        {
            // Use the stored _UniqueKey to look up instances
            if (entry.ContainsKey("_UniqueKey"))
            {
                string uniqueKey = entry["_UniqueKey"].ToString();

                if (typeToElementsMap.ContainsKey(uniqueKey))
                {
                    selectedInstances.AddRange(typeToElementsMap[uniqueKey]);
                }
            }
        }

        // Combine with existing selection
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
        }
        else
        {
            TaskDialog.Show("Selection", "No visible instances of the selected types were found in the target view(s).");
        }

        return Result.Succeeded;
    }

    private string GetSystemFamilyName(Element typeElement)
    {
        Parameter familyParam = typeElement.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
        return (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
            ? familyParam.AsString()
            : null;
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
