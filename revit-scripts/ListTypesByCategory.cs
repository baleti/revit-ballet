using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ListTypesByCategory : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Step 1: Collect all unique categories from the families
        var allCategories = new FilteredElementCollector(doc)
            .OfClass(typeof(Family))
            .Cast<Family>()
            .Select(family => family.FamilyCategory)
            .Where(category => category != null)
            .GroupBy(category => category.Id) // Group by the unique Category Id
            .Select(group => group.First())   // Take the first unique category in each group
            .ToList();

        // Step 2: Prepare a dictionary to map category names to their ElementIds
        Dictionary<string, ElementId> categoryNameToIdMap = allCategories
            .ToDictionary(category => category.Name, category => category.Id);

        // Step 3: Prepare a list of categories for user selection
        List<Dictionary<string, object>> categoryEntries = new List<Dictionary<string, object>>();

        foreach (var category in allCategories)
        {
            var entry = new Dictionary<string, object>
            {
                { "Category Name", category.Name }
            };
            categoryEntries.Add(entry);
        }

        // Step 4: Display the list of categories using CustomGUIs.DataGrid
        var categoryPropertyNames = new List<string> { "Category Name" };
        var selectedCategoryEntry = CustomGUIs.DataGrid(categoryEntries, categoryPropertyNames, spanAllScreens: false).FirstOrDefault();

        if (selectedCategoryEntry == null)
        {
            return Result.Cancelled; // No category selection made
        }

        // Step 5: Get the selected category name and map it to its corresponding ElementId
        string selectedCategoryName = selectedCategoryEntry["Category Name"] as string;
        ElementId selectedCategoryId = categoryNameToIdMap[selectedCategoryName];

        // Step 6: Collect all families in the selected category
        var allFamiliesInCategory = new FilteredElementCollector(doc)
            .OfClass(typeof(Family))
            .Cast<Family>()
            .Where(family => family.FamilyCategory.Id == selectedCategoryId)
            .ToList();

        // Step 7: Prepare a list of types within the selected category in one go
        List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
        Dictionary<string, FamilySymbol> typeElementMap = new Dictionary<string, FamilySymbol>(); // Map unique keys to FamilySymbols

        // Collect all FamilySymbol elements in the selected category
        var familySymbolsInCategory = new FilteredElementCollector(doc)
            .OfClass(typeof(FamilySymbol))
            .Where(symbol => symbol.Category.Id == selectedCategoryId) // Filter by the selected category
            .Cast<FamilySymbol>()
            .ToList();

        // Iterate through the FamilySymbols, instead of iterating through families
        foreach (var familySymbol in familySymbolsInCategory)
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

        // Step 8: Display the list of types using CustomGUIs.DataGrid
        var propertyNames = new List<string> { "Type Name", "Family", "Category" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, spanAllScreens: false);

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled; // No selection made
        }

        // Step 9: Collect ElementIds of the selected types using the typeElementMap
        List<ElementId> selectedTypeIds = selectedEntries
            .Select(entry => 
            {
                string uniqueKey = $"{entry["Family"]}:{entry["Type Name"]}";
                return typeElementMap[uniqueKey].Id; // Retrieve the FamilySymbol from the map and get its Id
            })
            .ToList();

        // Step 10: Collect all instances of the selected types in the model
        var selectedInstances = new FilteredElementCollector(doc)
            .WherePasses(new ElementMulticlassFilter(new List<System.Type> { typeof(FamilyInstance), typeof(ElementType) }))
            .Where(x => selectedTypeIds.Contains(x.GetTypeId())) // Filter elements by type
            .Select(x => x.Id)
            .ToList();

        // Step 11: Select the instances in the model
        uidoc.SetSelectionIds(selectedInstances);

        return Result.Succeeded;
    }
}
