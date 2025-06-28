using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectFamilyTypeInstancesInProject : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Step 1: Prepare a list of types across all categories in the current project
        List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
        Dictionary<string, FamilySymbol> typeElementMap = new Dictionary<string, FamilySymbol>(); // Map unique keys to FamilySymbols

        // Collect all FamilySymbol elements in the project
        var familySymbols = new FilteredElementCollector(doc)
            .OfClass(typeof(FamilySymbol))
            .Cast<FamilySymbol>()
            .ToList();

        // Iterate through the FamilySymbols, and collect their info
        foreach (var familySymbol in familySymbols)
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

        // Step 2: Display the list of types using CustomGUIs.DataGrid
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

        // Step 4: Collect all instances of the selected types in the model
        var selectedInstances = new FilteredElementCollector(doc)
            .WherePasses(new ElementMulticlassFilter(new List<System.Type> { typeof(FamilyInstance), typeof(ElementType) }))
            .Where(x => selectedTypeIds.Contains(x.GetTypeId())) // Filter elements by type
            .Select(x => x.Id)
            .ToList();

        // Step 5: Add the new selection to the existing selection
        ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();
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
        uidoc.SetSelectionIds(combinedSelection);

        return Result.Succeeded;
    }
}
