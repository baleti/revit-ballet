// SelectFamilyTypesInProjectByCategory.cs
// C# 7.3 – Revit API 2024

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace YourCompany.YourAddIn
{
    /// <summary>
    /// Lightweight wrapper so the DataGrid can show a list of categories.
    /// </summary>
    public class CategoryWrapper
    {
        public ElementId Id  { get; }
        public string    Name { get; }

        public CategoryWrapper(ElementId id, string name)
        {
            Id   = id;
            Name = name;
        }
    }

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
             * 1.  Show a DataGrid with every category that has at least
             *     one ElementType in the project.
             * -----------------------------------------------------------------*/

            // Collect every ElementType first (we need them later anyway).
            List<ElementType> allElementTypes = new FilteredElementCollector(doc)
                                                .OfClass(typeof(ElementType))
                                                .Cast<ElementType>()
                                                .ToList();

            // Build a unique list of those categories.
            List<CategoryWrapper> categoryWrappers =
                allElementTypes
                    .Where(t => t.Category != null)
                    .Select(t => new CategoryWrapper(t.Category.Id, t.Category.Name))
                    .GroupBy(cw => cw.Id.IntegerValue)      // ensure uniqueness
                    .Select(g => g.First())
                    .OrderBy(cw => cw.Name)                 // nice to sort alphabetically
                    .ToList();

            // Let the user pick the categories.
            List<CategoryWrapper> selectedCategories =
                CustomGUIs.DataGrid<CategoryWrapper>(
                    categoryWrappers,
                    new List<string> { "Name" }            // which property columns to show
                    );

            if (selectedCategories.Count == 0)
                return Result.Cancelled;                    // user hit Cancel / closed box

            /* ------------------------------------------------------------------
             * 2.  Filter the previously-collected ElementTypes so we keep only
             *     those whose Category.Id is in the user’s selection.
             * -----------------------------------------------------------------*/

            HashSet<ElementId> allowedCategoryIds =
                selectedCategories
                    .Select(cw => cw.Id)
                    .ToHashSet();

            List<ElementType> filteredTypes =
                allElementTypes
                    .Where(t => t.Category != null && allowedCategoryIds.Contains(t.Category.Id))
                    .ToList();

            if (filteredTypes.Count == 0)
            {
                TaskDialog.Show("Select Types", "No family or system types exist for the selected categories.");
                return Result.Cancelled;
            }

            /* ------------------------------------------------------------------
             * 3.  Build the rows for the second DataGrid (Type  | Family | Category)
             *     and keep a map so we can retrieve the ElementId later.
             * -----------------------------------------------------------------*/

            var typeRows       = new List<Dictionary<string, object>>();
            var keyToElemType  = new Dictionary<string, ElementType>();

            foreach (ElementType et in filteredTypes)
            {
                string familyName;
                string categoryName = et.Category?.Name ?? "N/A";

                if (et is FamilySymbol fs)                  // loaded family type
                {
                    familyName = fs.Family.Name;
                }
                else                                        // system type (wall, floor, etc.)
                {
                    Parameter p = et.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                    familyName  = (p != null && !string.IsNullOrEmpty(p.AsString())) ? p.AsString() : "System Type";
                }

                var row = new Dictionary<string, object>
                {
                    { "Type Name",  et.Name },
                    { "Family",     familyName },
                    { "Category",   categoryName }
                };
                typeRows.Add(row);

                // key = Family:Type; add suffix if duplicate
                string key = $"{familyName}:{et.Name}";
                int    i   = 1;
                while (keyToElemType.ContainsKey(key))
                    key = $"{familyName}:{et.Name}_{i++}";

                keyToElemType[key] = et;
            }

            /* ------------------------------------------------------------------
             * 4.  Show the second grid so the user can pick the types.
             * -----------------------------------------------------------------*/
            var propertyNames = new List<string> { "Type Name", "Family", "Category" };

            var selectedRows = CustomGUIs.DataGrid(
                                   typeRows,
                                   propertyNames,
                                   spanAllScreens: false);

            if (selectedRows.Count == 0)
                return Result.Cancelled;

            /* ------------------------------------------------------------------
             * 5.  Turn the selected rows into ElementIds and select them in Revit.
             * -----------------------------------------------------------------*/
            List<ElementId> idsToSelect = selectedRows
                .Select(r =>
                {
                    string key = $"{r["Family"]}:{r["Type Name"]}";
                    if (!keyToElemType.ContainsKey(key))
                    {
                        // handle “_1” / “_2” duplicates
                        key = keyToElemType.Keys.FirstOrDefault(k => k.StartsWith($"{r["Family"]}:{r["Type Name"]}"));
                    }
                    return key != null ? keyToElemType[key].Id : null;
                })
                .Where(id => id != null)
                .ToList();

            uidoc.Selection.SetElementIds(idsToSelect);

            return Result.Succeeded;
        }
    }
}
