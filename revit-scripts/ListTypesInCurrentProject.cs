using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ListAllTypesInProject : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Step 1: Prepare a list of types across all categories in the current project, excluding null categories and imported DWG types
        List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
        Dictionary<string, ElementType> typeElementMap = new Dictionary<string, ElementType>(); // Map unique keys to ElementTypes

        // Collect all ElementType elements in the project (including FamilySymbol, WallType, FloorType, RoofType, etc.)
        var allTypes = new FilteredElementCollector(doc)
            .OfClass(typeof(ElementType))
            .Cast<ElementType>()
            .Where(elementType => elementType.Category != null) // Exclude types with null categories
            .Where(elementType => !(elementType is ImportInstance)) // Exclude imported DWG types
            .ToList();

        // Iterate through all ElementType elements
        foreach (var elementType in allTypes)
        {
            var entry = new Dictionary<string, object>
            {
                { "Type Name", elementType.Name },
                { "Family", elementType is FamilySymbol familySymbol ? familySymbol.Family.Name : "N/A" },
                { "Category", elementType.Category.Name }
            };

            // Store the ElementType with a unique key (Family:Type)
            string uniqueKey = $"{entry["Family"]}:{elementType.Name}";
            typeElementMap[uniqueKey] = elementType;

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
                return typeElementMap[uniqueKey].Id; // Retrieve the ElementType from the map and get its Id
            })
            .ToList();

        // Step 4: Collect all instances of the selected types in the model
        var selectedInstances = new FilteredElementCollector(doc)
            .WherePasses(new ElementMulticlassFilter(new List<System.Type> { typeof(FamilyInstance), typeof(ElementType) }))
            .Where(x => selectedTypeIds.Contains(x.GetTypeId())) // Filter elements by type
            .Select(x => x.Id)
            .ToList();

        // Step 5: Select the instances in the model
        uidoc.Selection.SetElementIds(selectedInstances);

        return Result.Succeeded;
    }
}
