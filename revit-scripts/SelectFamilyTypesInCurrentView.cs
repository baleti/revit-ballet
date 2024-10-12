using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectFamilyTypesInCurrentView : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Step 1: Prepare a list of unique family types across all categories in the current view
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
            Family family = familySymbol.Family;

            var entry = new Dictionary<string, object>
            {
                { "Type Name", familySymbol.Name },
                { "Family", family.Name },
                { "Category", familySymbol.Category.Name }
            };

            string uniqueKey = $"{family.Name}:{familySymbol.Name}";
            typeElementMap[uniqueKey] = familySymbol;

            typeEntries.Add(entry);
        }

        // Step 2: Display the list of unique family types using CustomGUIs.DataGrid
        var propertyNames = new List<string> { "Type Name", "Family", "Category" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, spanAllScreens: false);

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled; // No selection made
        }

        // Step 3: Collect ElementIds of the selected family types using the typeElementMap
        List<ElementId> selectedTypeIds = selectedEntries
            .Select(entry =>
            {
                string uniqueKey = $"{entry["Family"]}:{entry["Type Name"]}";
                return typeElementMap[uniqueKey].Id;
            })
            .ToList();

        // Step 4: Retrieve and display all available parameters for the selected family types
        List<Dictionary<string, object>> parameterEntries = new List<Dictionary<string, object>>();

        foreach (ElementId typeId in selectedTypeIds)
        {
            FamilySymbol familySymbol = doc.GetElement(typeId) as FamilySymbol;

            if (familySymbol != null)
            {
                var entry = new Dictionary<string, object>
                {
                    { "Type Name", familySymbol.Name },
                    { "Family", familySymbol.Family.Name },
                    { "Category", familySymbol.Category.Name }
                };

                // Add all available parameters
                foreach (Parameter param in familySymbol.Parameters)
                {
                    string paramName = param.Definition.Name;
                    string paramValue = param.AsValueString() ?? param.AsString() ?? "None";
                    entry[paramName] = paramValue;
                }

                parameterEntries.Add(entry);
            }
        }

        // Step 5: Display the parameters of the selected family types in a second DataGrid
        var paramPropertyNames = parameterEntries.FirstOrDefault()?.Keys.ToList();
        var finalSelection = CustomGUIs.DataGrid(parameterEntries, paramPropertyNames, spanAllScreens: false);

        if (finalSelection.Count == 0)
        {
            return Result.Cancelled; // No selection made
        }

        // Step 6: Select the family types in the project
        List<ElementId> finalSelectedTypeIds = finalSelection
            .Select(entry =>
            {
                string uniqueKey = $"{entry["Family"]}:{entry["Type Name"]}";
                return typeElementMap[uniqueKey].Id;
            })
            .ToList();

        // Add the selected family types to the current selection
        uidoc.Selection.SetElementIds(finalSelectedTypeIds);

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
        BoundingBoxXYZ bbox = element.get_BoundingBox(view);
        return bbox != null;
    }
}
