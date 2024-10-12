using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectFamilyTypesInProject : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Step 1: Prepare a list of family types (FamilySymbols) across all categories in the current project
        List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
        Dictionary<string, FamilySymbol> typeElementMap = new Dictionary<string, FamilySymbol>(); // Map unique keys to FamilySymbols

        // Collect all FamilySymbol elements in the project
        var familySymbols = new FilteredElementCollector(doc)
            .OfClass(typeof(FamilySymbol))
            .Cast<FamilySymbol>()
            .ToList();

        // Iterate through the FamilySymbols, and collect their basic info
        foreach (var familySymbol in familySymbols)
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

        // Step 2: Display the list of family types using CustomGUIs.DataGrid
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

        // Step 4: Retrieve and display all available parameters for the selected types
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

        // Step 5: Display the parameters of the selected types in a second DataGrid
        var paramPropertyNames = parameterEntries.FirstOrDefault()?.Keys.ToList();
        var finalSelection = CustomGUIs.DataGrid(parameterEntries, paramPropertyNames, spanAllScreens: false);

        if (finalSelection.Count == 0)
        {
            return Result.Cancelled; // No selection made
        }

        return Result.Succeeded;
    }
}
