using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class FilterSelectedByCategory : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get currently selected elements
        ICollection<ElementId> selectedElementIds = uidoc.GetSelectionIds();
        if (!selectedElementIds.Any())
        {
            TaskDialog.Show("Filter Categories", "No elements are selected.");
            return Result.Cancelled;
        }

        // Get unique categories from selected elements
        var selectedElements = selectedElementIds.Select(id => doc.GetElement(id)).ToList();
        var categoryGroups = selectedElements
            .Where(el => el.Category != null) // Exclude elements with no category
            .GroupBy(el => el.Category.Id)
            .Select(group => new
            {
                Category = group.First().Category,
                Elements = group.ToList()
            })
            .ToList();

        if (!categoryGroups.Any())
        {
            TaskDialog.Show("Filter Categories", "No valid categories found in the selection.");
            return Result.Cancelled;
        }

        // Prepare data for the DataGrid
        List<Dictionary<string, object>> dataGridEntries = categoryGroups.Select(catGroup => new Dictionary<string, object>
        {
            { "Category Name", catGroup.Category.Name },
            { "Element Count", catGroup.Elements.Count }
        }).ToList();

        List<string> propertyNames = new List<string> { "Category Name", "Element Count" };

        // Show DataGrid to let user select categories
        var selectedEntries = CustomGUIs.DataGrid(dataGridEntries, propertyNames, false);

        if (!selectedEntries.Any())
        {
            return Result.Cancelled;
        }

        // Get selected category names
        var selectedCategoryNames = selectedEntries
            .Select(entry => entry["Category Name"].ToString())
            .ToHashSet();

        // Collect all elements belonging to the selected categories
        var filteredElementIds = categoryGroups
            .Where(catGroup => selectedCategoryNames.Contains(catGroup.Category.Name))
            .SelectMany(catGroup => catGroup.Elements)
            .Select(el => el.Id)
            .ToList();

        // Update selection in Revit
        uidoc.SetSelectionIds(filteredElementIds);

        return Result.Succeeded;
    }
}
