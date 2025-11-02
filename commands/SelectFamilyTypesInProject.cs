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

        // Step 1: Collect all element types (both loaded family types and system types) in the current project.
        List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
        Dictionary<string, ElementType> typeElementMap = new Dictionary<string, ElementType>();

        var elementTypes = new FilteredElementCollector(doc)
            .OfClass(typeof(ElementType))
            .Cast<ElementType>()
            .ToList();

        foreach (var elementType in elementTypes)
        {
            string typeName = elementType.Name;
            string familyName = "";
            string categoryName = "";

            // If the element type is a FamilySymbol (loaded family), gather its family name and category.
            if (elementType is FamilySymbol fs)
            {
                familyName = fs.Family.Name;
                categoryName = fs.Category != null ? fs.Category.Name : "N/A";
            }
            else
            {
                // For system types (like WallType), try to extract the family name via the built-in parameter.
                Parameter familyParam = elementType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                familyName = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
                    ? familyParam.AsString()
                    : "System Type";
                categoryName = elementType.Category != null ? elementType.Category.Name : "N/A";
            }

            var entry = new Dictionary<string, object>
            {
                { "Type Name", typeName },
                { "Family", familyName },
                { "Category", categoryName }
            };

            // Build a unique key based on family and type name.
            string uniqueKey = $"{familyName}:{typeName}";
            // If duplicate keys exist, append an index.
            if (typeElementMap.ContainsKey(uniqueKey))
            {
                int duplicateIndex = 1;
                string newKey = uniqueKey + "_" + duplicateIndex;
                while (typeElementMap.ContainsKey(newKey))
                {
                    duplicateIndex++;
                    newKey = uniqueKey + "_" + duplicateIndex;
                }
                uniqueKey = newKey;
            }

            typeElementMap[uniqueKey] = elementType;
            typeEntries.Add(entry);
        }

        // Step 2: Display a DataGrid for the user to select types.
        // 'CustomGUIs.DataGrid' is assumed to be a method that displays the data in a grid and returns the selected rows.
        var propertyNames = new List<string> { "Type Name", "Family", "Category" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, spanAllScreens: false);

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled; // No selection made
        }

        // Step 3: Retrieve the ElementIds from the selected entries using the typeElementMap.
        List<ElementId> selectedTypeIds = selectedEntries
            .Select(entry =>
            {
                string uniqueKey = $"{entry["Family"]}:{entry["Type Name"]}";
                if (!typeElementMap.ContainsKey(uniqueKey))
                {
                    uniqueKey = typeElementMap.Keys.FirstOrDefault(k => k.StartsWith($"{entry["Family"]}:{entry["Type Name"]}"));
                }
                return typeElementMap.ContainsKey(uniqueKey) ? typeElementMap[uniqueKey].Id : null;
            })
            .Where(id => id != null)
            .ToList();

        // Step 4: Set the selected ElementIds in the UIDocument, which selects them in the Revit UI.
        uidoc.SetSelectionIds(selectedTypeIds);

        return Result.Succeeded;
    }
}
