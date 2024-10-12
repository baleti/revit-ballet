using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

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

        // Step 1: Prepare a list of unique types across all categories in the current view
        List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
        Dictionary<string, FamilySymbol> typeElementMap = new Dictionary<string, FamilySymbol>(); // Map unique keys to FamilySymbols

        // Get the current view's Id
        ElementId currentViewId = doc.ActiveView.Id;

        // Collect all visible FamilyInstance elements in the current view
        var familyInstancesInView = new FilteredElementCollector(doc, currentViewId)
            .OfClass(typeof(FamilyInstance))
            .WhereElementIsNotElementType()
            .Cast<FamilyInstance>()
            .Where(fi => IsElementVisibleInView(fi, doc.ActiveView))
            .ToList();

        // Get the unique FamilySymbols used by these instances
        var familySymbolsInView = familyInstancesInView
            .Select(fi => fi.Symbol)
            .Distinct(new FamilySymbolComparer())
            .ToList();

        // Iterate through the unique FamilySymbols and collect their info
        foreach (var familySymbol in familySymbolsInView)
        {
            // Get the family for the current symbol
            Family family = familySymbol.Family;

            var entry = new Dictionary<string, object>
            {
                { "Type Name", familySymbol.Name },
                { "Family", family.Name },
                { "Category", familySymbol.Category.Name }
            };

            // Store the FamilySymbol with a unique key (Family:Type)
            string uniqueKey = $"{family.Name}:{familySymbol.Name}";
            typeElementMap[uniqueKey] = familySymbol;

            typeEntries.Add(entry);
        }

        // Step 2: Display the list of unique types using CustomGUIs.DataGrid
        var propertyNames = new List<string> { "Type Name", "Family", "Category" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, spanAllScreens: false);

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled; // No selection made
        }

        // Step 3: Collect ElementIds of the selected types using the typeElementMap
        List<ElementId> selectedTypeIds = selectedEntries
            .Select(entry =>
            {
                string uniqueKey = $"{entry["Family"]}:{entry["Type Name"]}";
                return typeElementMap[uniqueKey].Id; // Retrieve the FamilySymbol from the map and get its Id
            })
            .ToList();

        // Step 4: Collect all visible instances of the selected types in the current view
        var selectedInstances = familyInstancesInView
            .Where(fi => selectedTypeIds.Contains(fi.Symbol.Id))
            .Select(fi => fi.Id)
            .ToList();

        // Step 5: Add the new selection to the existing selection
        ICollection<ElementId> currentSelection = uidoc.Selection.GetElementIds();
        List<ElementId> combinedSelection = new List<ElementId>(currentSelection);

        // Add new instances to the combined selection without duplicates
        foreach (var instanceId in selectedInstances)
        {
            if (!combinedSelection.Contains(instanceId))
            {
                combinedSelection.Add(instanceId);
            }
        }

        // Update the selection with both previous and newly selected elements
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

    // Custom comparer for FamilySymbol to handle duplicates
    private class FamilySymbolComparer : IEqualityComparer<FamilySymbol>
    {
        public bool Equals(FamilySymbol x, FamilySymbol y)
        {
            return x.Id.IntegerValue == y.Id.IntegerValue;
        }

        public int GetHashCode(FamilySymbol obj)
        {
            return obj.Id.IntegerValue.GetHashCode();
        }
    }

    // Helper method to check if an element is visible in a given view
    private bool IsElementVisibleInView(Element element, View view)
    {
        // Check if the element is visible in the view by testing its bounding box
        BoundingBoxXYZ bbox = element.get_BoundingBox(view);
        return bbox != null;
    }
}
