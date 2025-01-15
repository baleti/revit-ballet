using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class FilterSelectedCategories : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get currently selected elements
        ICollection<ElementId> selectedElementIds = uidoc.Selection.GetElementIds();
        if (!selectedElementIds.Any())
        {
            TaskDialog.Show("Filter Categories", "No elements are selected.");
            return Result.Cancelled;
        }

        // Get unique categories from selected elements
        var selectedElements = selectedElementIds.Select(id => doc.GetElement(id)).ToList();
        var categories = selectedElements
            .Select(el => el.Category)
            .Where(cat => cat != null) // Exclude null categories
            .Distinct()
            .ToList();

        if (!categories.Any())
        {
            TaskDialog.Show("Filter Categories", "No valid categories found in the selection.");
            return Result.Cancelled;
        }

        // Prepare data for the DataGrid
        List<Dictionary<string, object>> dataGridEntries = categories.Select(cat => new Dictionary<string, object>
        {
            { "Category Name", cat.Name }
        }).ToList();

        List<string> propertyNames = new List<string> { "Category Name" };

        // Show DataGrid to let user select categories
        var selectedEntries = CustomGUIs.DataGrid(dataGridEntries, propertyNames, false);

        if (!selectedEntries.Any())
        {
            TaskDialog.Show("Filter Categories", "No categories were selected.");
            return Result.Cancelled;
        }

        // Get selected category names
        var selectedCategoryNames = selectedEntries
            .Select(entry => entry["Category Name"].ToString())
            .ToHashSet();

        // Filter elements by selected categories
        var filteredElements = selectedElements
            .Where(el => selectedCategoryNames.Contains(el.Category.Name))
            .Select(el => el.Id)
            .ToList();

        // Update selection in Revit
        uidoc.Selection.SetElementIds(filteredElements);

        return Result.Succeeded;
    }
}
