using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectFamilyTypeInstancesInCurrentView : IExternalCommand
{
    // List of built-in categories to include
    private static readonly BuiltInCategory[] TargetCategories = new BuiltInCategory[]
    {
        BuiltInCategory.OST_PipeFitting,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_PipeAccessory,
        BuiltInCategory.OST_PlumbingFixtures,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_Roofs
    };

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

        // Create a multi-category filter for our target categories
        ElementMulticategoryFilter categoryFilter = new ElementMulticategoryFilter(TargetCategories.ToList());

        // Collect all relevant elements in the current view
        var elementsInView = new FilteredElementCollector(doc, currentViewId)
            .WherePasses(categoryFilter)
            .WhereElementIsNotElementType()
            .Where(e => IsElementVisibleInView(e, doc.ActiveView))
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
            else if (element is MEPCurve mepCurve) // Handles pipes
            {
                typeElement = doc.GetElement(mepCurve.GetTypeId());
                typeName = typeElement.Name;
                familyName = "Pipe";
            }
            else if (element is Wall wall) // Handles walls
            {
                typeElement = doc.GetElement(wall.GetTypeId());
                typeName = typeElement.Name;
                familyName = "Wall";
            }
            else if (element is Floor floor) // Handles floors
            {
                typeElement = doc.GetElement(floor.GetTypeId());
                typeName = typeElement.Name;
                familyName = "Floor";
            }
            else if (element is RoofBase roof) // Handles roofs
            {
                typeElement = doc.GetElement(roof.GetTypeId());
                typeName = typeElement.Name;
                familyName = "Roof";
            }
            
            if (typeElement != null)
            {
                string uniqueKey = $"{familyName}:{typeName}";
                if (!typeElementMap.ContainsKey(uniqueKey))
                {
                    var entry = new Dictionary<string, object>
                    {
                        { "Type Name", typeName },
                        { "Family", familyName },
                        { "Category", element.Category.Name }
                    };

                    typeElementMap[uniqueKey] = typeElement;
                    typeEntries.Add(entry);
                }
            }
        }

        // Display the list of unique types
        var propertyNames = new List<string> { "Type Name", "Family", "Category" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, spanAllScreens: false);

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled;
        }

        // Collect ElementIds of selected types
        List<ElementId> selectedTypeIds = selectedEntries
            .Select(entry =>
            {
                string uniqueKey = $"{entry["Family"]}:{entry["Type Name"]}";
                return typeElementMap[uniqueKey].Id;
            })
            .ToList();

        // Find all instances of selected types
        var selectedInstances = elementsInView
            .Where(e =>
            {
                if (e is FamilyInstance fi)
                    return selectedTypeIds.Contains(fi.Symbol.Id);
                if (e is MEPCurve mc)
                    return selectedTypeIds.Contains(mc.GetTypeId());
                if (e is Wall wall)
                    return selectedTypeIds.Contains(wall.GetTypeId());
                if (e is Floor floor)
                    return selectedTypeIds.Contains(floor.GetTypeId());
                if (e is RoofBase roof)
                    return selectedTypeIds.Contains(roof.GetTypeId());
                return false;
            })
            .Select(e => e.Id)
            .ToList();

        // Combine with existing selection
        ICollection<ElementId> currentSelection = uidoc.Selection.GetElementIds();
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
            uidoc.Selection.SetElementIds(combinedSelection);
        }
        else
        {
            TaskDialog.Show("Selection", "No visible instances of the selected types were found in the current view.");
        }

        return Result.Succeeded;
    }

    private bool IsElementVisibleInView(Element element, View view)
    {
        BoundingBoxXYZ bbox = element.get_BoundingBox(view);
        return bbox != null;
    }
}
