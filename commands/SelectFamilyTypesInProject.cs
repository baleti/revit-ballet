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

        // Step 1: Collect all element types (both loaded family types and system types) in the current project
        List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();

        var elementTypes = new FilteredElementCollector(doc)
            .OfClass(typeof(ElementType))
            .Cast<ElementType>()
            .ToList();

        // Create a map to count instances of each type
        Dictionary<ElementId, int> typeInstanceCounts = new Dictionary<ElementId, int>();

        // Collect all instances and count them by type
        var allInstances = new FilteredElementCollector(doc)
            .WhereElementIsNotElementType()
            .Where(x => x.GetTypeId() != null && x.GetTypeId() != ElementId.InvalidElementId)
            .ToList();

        foreach (var instance in allInstances)
        {
            ElementId typeId = instance.GetTypeId();
            if (!typeInstanceCounts.ContainsKey(typeId))
            {
                typeInstanceCounts[typeId] = 0;
            }
            typeInstanceCounts[typeId]++;
        }

        foreach (var elementType in elementTypes)
        {
            string typeName = elementType.Name;
            string familyName = "";
            string categoryName = "";

            // If the element type is a FamilySymbol (loaded family), gather its family name and category
            if (elementType is FamilySymbol fs)
            {
                familyName = fs.Family.Name;
                categoryName = fs.Category != null ? fs.Category.Name : "N/A";

                // Filter out DWG import symbols
                if (categoryName.Contains("Import Symbol") || familyName.Contains("Import Symbol"))
                    continue;
            }
            else
            {
                // For system types (like WallType, FloorType), try to extract the family name
                familyName = GetSystemFamilyName(elementType) ?? "System Type";
                categoryName = elementType.Category != null ? elementType.Category.Name : "N/A";

                // Filter out DWG import symbols for system types too
                if (categoryName.Contains("Import Symbol") || familyName.Contains("Import Symbol"))
                    continue;
            }

            // Get the instance count for this type
            int count = typeInstanceCounts.ContainsKey(elementType.Id) ? typeInstanceCounts[elementType.Id] : 0;

            var entry = new Dictionary<string, object>
            {
                { "Type Name", typeName },
                { "Family", familyName },
                { "Category", categoryName },
                { "Count", count },
                { "ElementIdObject", elementType.Id }  // Store ElementId for reliable lookup after edits
            };

            typeEntries.Add(entry);
        }

        // Sort by Category, then Family, then Type Name
        typeEntries = typeEntries
            .OrderBy(e => e["Category"].ToString())
            .ThenBy(e => e["Family"].ToString())
            .ThenBy(e => e["Type Name"].ToString())
            .ToList();

        // Step 2: Display a DataGrid for the user to select types
        // Set UIDocument for edit mode support
        CustomGUIs.SetCurrentUIDocument(uidoc);

        var propertyNames = new List<string> { "Category", "Family", "Type Name", "Count" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, false);

        // Apply any pending edits (family/type renames)
        if (CustomGUIs.HasPendingEdits())
        {
            CustomGUIs.ApplyCellEditsToEntities();
        }

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled;
        }

        // Step 3: Retrieve the ElementIds from the selected entries using the stored ElementIdObject
        List<ElementId> selectedTypeIds = selectedEntries
            .Where(entry => entry.ContainsKey("ElementIdObject") && entry["ElementIdObject"] is ElementId)
            .Select(entry => (ElementId)entry["ElementIdObject"])
            .ToList();

        // Step 4: Set the selected ElementIds in the UIDocument, which selects them in the Revit UI
        uidoc.SetSelectionIds(selectedTypeIds);

        return Result.Succeeded;
    }

    private string GetSystemFamilyName(Element typeElement)
    {
        Parameter familyParam = typeElement.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
        return (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
            ? familyParam.AsString()
            : null;
    }
}
