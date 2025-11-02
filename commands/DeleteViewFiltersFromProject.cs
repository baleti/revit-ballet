using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// A row for the DataGrid representing a single Parameter Filter,
/// showing its name and how many views use it.
/// </summary>
public class DeleteFilterRow
{
    public string FilterName { get; set; }
    public int Views { get; set; }
}

/// <summary>
/// This command displays a list of Parameter Filters (including
/// how many views each filter is on). Once the user selects rows
/// and presses Enter/double-clicks, the selected filters are 
/// deleted from the project.
/// </summary>
[Transaction(TransactionMode.Manual)]
public class DeleteViewFiltersFromProject : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // -------------------------------------------------------------
        // 1. Collect all ParameterFilterElements
        // -------------------------------------------------------------
        List<ParameterFilterElement> allFilters = new FilteredElementCollector(doc)
            .OfClass(typeof(ParameterFilterElement))
            .Cast<ParameterFilterElement>()
            .OrderBy(f => f.Name)
            .ToList();

        if (!allFilters.Any())
        {
            TaskDialog.Show("Delete Filters", "No Parameter Filters found in this project.");
            return Result.Succeeded;
        }

        // -------------------------------------------------------------
        // 2. Compute how many views each filter is used on
        //    (Similar logic to previous commands)
        // -------------------------------------------------------------
        // Create a dictionary from filterId => 0 usage initially
        Dictionary<ElementId, int> usageCountByFilter = allFilters
            .ToDictionary(f => f.Id, f => 0);

        // Collect all Views to check usage
        List<View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .ToList();

        foreach (View v in allViews)
        {
            // Attempt to get filters
            ICollection<ElementId> directFilters;
            try
            {
                directFilters = v.GetFilters();
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                // Some view types (schedules, sheets, legends) don't support filters
                continue;
            }

            // Check if the view uses a template
            ICollection<ElementId> templateFilters = new List<ElementId>();
            if (v.ViewTemplateId != ElementId.InvalidElementId)
            {
                try
                {
                    var templateView = doc.GetElement(v.ViewTemplateId) as View;
                    if (templateView != null)
                    {
                        templateFilters = templateView.GetFilters();
                    }
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    // skip
                }
            }

            // Increment usage for each filter that appears in direct or template filters
            foreach (ParameterFilterElement filterElem in allFilters)
            {
                ElementId fid = filterElem.Id;
                bool usedDirectly = directFilters.Contains(fid);
                bool usedFromTemplate = templateFilters.Contains(fid);

                if (usedDirectly || usedFromTemplate)
                {
                    usageCountByFilter[fid]++;
                }
            }
        }

        // -------------------------------------------------------------
        // 3. Build a dictionary FilterName => FilterId 
        //    for deletion later
        // -------------------------------------------------------------
        Dictionary<string, ElementId> filterNameToId = new Dictionary<string, ElementId>();
        foreach (var f in allFilters)
        {
            filterNameToId[f.Name] = f.Id;
        }

        // -------------------------------------------------------------
        // 4. Build the single DataGrid data (FilterName, Views)
        // -------------------------------------------------------------
        var gridData = new List<DeleteFilterRow>();
        foreach (var filterElem in allFilters)
        {
            gridData.Add(new DeleteFilterRow
            {
                FilterName = filterElem.Name,
                Views      = usageCountByFilter[filterElem.Id]
            });
        }

        // Columns for the DataGrid
        // We'll display: FilterName, Views
        List<string> columns = new List<string> { "FilterName", "Views" };

        // -------------------------------------------------------------
        // 5. Show the DataGrid so user can pick which filters to delete
        // -------------------------------------------------------------
        List<DeleteFilterRow> selectedRows = CustomGUIs.DataGrid(
            entries: gridData,
            propertyNames: columns,
            initialSelectionIndices: null,
            Title: "Select Filters to Delete"
        );

        // If user made no selection, just end quietly
        if (!selectedRows.Any())
        {
            return Result.Succeeded;
        }

        // -------------------------------------------------------------
        // 6. Delete the selected filters
        // -------------------------------------------------------------
        using (Transaction t = new Transaction(doc, "Delete Filters"))
        {
            t.Start();

            foreach (DeleteFilterRow row in selectedRows)
            {
                string filterName = row.FilterName;
                if (filterNameToId.TryGetValue(filterName, out ElementId filterId))
                {
                    doc.Delete(filterId);
                }
            }

            t.Commit();
        }

        return Result.Succeeded;
    }
}
