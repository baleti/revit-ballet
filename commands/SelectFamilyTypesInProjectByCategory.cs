// SelectFamilyTypesInProjectByCategory.cs
// C# 7.3 – Revit API 2024

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace YourCompany.YourAddIn
{
    [Transaction(TransactionMode.Manual)]
    public class SelectFamilyTypesInProjectByCategory : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string          message,
            ElementSet          elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uidoc.Document;

            /* ------------------------------------------------------------------
             * 1. Collect every ElementType in the project and build
             *    one Dictionary-row per distinct category.
             * -----------------------------------------------------------------*/

            List<ElementType> allElementTypes = new FilteredElementCollector(doc)
                                                .OfClass(typeof(ElementType))
                                                .Cast<ElementType>()
                                                .ToList();

            var categoryRows = new List<Dictionary<string, object>>();
            var seenIds      = new HashSet<int>();

            foreach (ElementType et in allElementTypes)
            {
                Category cat = et.Category;
                if (cat == null) continue;

                if (seenIds.Add(cat.Id.IntegerValue))          // only once per category
                {
                    categoryRows.Add(new Dictionary<string, object>
                    {
                        { "Id",   cat.Id },                      // keep the ElementId!
                        { "Name", cat.Name }
                    });
                }
            }

            // Show the category DataGrid.  --- NOTE: keep "Id" so it returns!
            var categoryColumns = new List<string> { "Name", "Id" };

            List<Dictionary<string, object>> selectedCatRows =
                CustomGUIs.DataGrid(
                    categoryRows,
                    categoryColumns,
                    spanAllScreens: false);

            if (selectedCatRows.Count == 0)
                return Result.Cancelled;

            /* ------------------------------------------------------------------
             * 2. Filter the ElementTypes so we keep only those whose category
             *    matches the user’s selection.
             * -----------------------------------------------------------------*/

            HashSet<ElementId> allowedCategoryIds =
                selectedCatRows
                    .Select(r => (ElementId)r["Id"])
                    .ToHashSet();

            List<ElementType> filteredTypes =
                allElementTypes
                    .Where(t => t.Category != null &&
                                allowedCategoryIds.Contains(t.Category.Id))
                    .ToList();

            if (filteredTypes.Count == 0)
            {
                TaskDialog.Show("Select Types",
                                "No family or system types exist for the selected categories.");
                return Result.Cancelled;
            }

            /* ------------------------------------------------------------------
             * 3. Build the rows for the second DataGrid (Type | Family | Category)
             *    and keep a map so we can retrieve the ElementId later.
             * -----------------------------------------------------------------*/

            var typeRows      = new List<Dictionary<string, object>>();
            var keyToElemType = new Dictionary<string, ElementType>();

            foreach (ElementType et in filteredTypes)
            {
                string familyName;
                string categoryName = et.Category?.Name ?? "N/A";

                if (et is FamilySymbol fs)
                {
                    familyName = fs.Family.Name;               // loaded family type
                }
                else
                {
                    Parameter p = et.get_Parameter(
                        BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                    familyName = (!string.IsNullOrEmpty(p?.AsString()))
                                 ? p.AsString()
                                 : "System Type";
                }

                var row = new Dictionary<string, object>
                {
                    { "Type Name", et.Name },
                    { "Family",    familyName },
                    { "Category",  categoryName }
                };
                typeRows.Add(row);

                // Unique key = Family:Type  (add suffix if duplicate)
                string key = $"{familyName}:{et.Name}";
                int    i   = 1;
                while (keyToElemType.ContainsKey(key))
                    key = $"{familyName}:{et.Name}_{i++}";

                keyToElemType[key] = et;
            }

            /* ------------------------------------------------------------------
             * 4. Show the family-/system-type grid.
             * -----------------------------------------------------------------*/

            var typeColumns = new List<string> { "Type Name", "Family", "Category" };

            var selectedTypeRows = CustomGUIs.DataGrid(
                                       typeRows,
                                       typeColumns,
                                       spanAllScreens: false);

            if (selectedTypeRows.Count == 0)
                return Result.Cancelled;

            /* ------------------------------------------------------------------
             * 5. Convert the selected rows into ElementIds and select them.
             * -----------------------------------------------------------------*/
            List<ElementId> idsToSelect = selectedTypeRows
                .Select(r =>
                {
                    string key = $"{r["Family"]}:{r["Type Name"]}";
                    if (!keyToElemType.ContainsKey(key))
                        key = keyToElemType.Keys.FirstOrDefault(
                                  k => k.StartsWith($"{r["Family"]}:{r["Type Name"]}"));
                    return key != null ? keyToElemType[key].Id : null;
                })
                .Where(id => id != null)
                .ToList();

            uidoc.SetSelectionIds(idsToSelect);

            return Result.Succeeded;
        }
    }
}
