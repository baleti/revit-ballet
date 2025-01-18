using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// A row in the FIRST grid (filter list). Displays only FilterName and number of Views.
/// </summary>
public class FilterGridRow
{
    public string FilterName { get; set; }
    public int Views { get; set; }
}

/// <summary>
/// A row in the SECOND grid (usage details).
/// Includes enough info to either open the view itself or the sheet.
/// </summary>
public class UsageGridRow
{
    public string FilterName { get; set; }
    public string ViewName   { get; set; }
    public string ViewType   { get; set; }
    public string SheetNumber { get; set; } // e.g. "A101"
    public string SheetName   { get; set; } // e.g. "First Floor Plan"

    // For opening the view:
    public int ViewId { get; set; }

    // For opening the sheet (if any):
    public int SheetId { get; set; } = -1;  // -1 means not placed on a sheet
}

/// <summary>
/// Abstract base class that implements IExternalCommand and handles all logic
/// for listing filters, collecting usage, and showing two DataGrids.
/// The final step of "opening" something is deferred to an abstract method
/// so that derived classes can choose to open views or sheets.
/// </summary>
public abstract class BaseFilterCommand : IExternalCommand
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
            TaskDialog.Show("Filters", "No Parameter Filters found in this project.");
            return Result.Succeeded;
        }

        // Map FilterName => FilterId for later
        Dictionary<string, ElementId> filterNameToId = new Dictionary<string, ElementId>();
        foreach (var f in allFilters)
        {
            filterNameToId[f.Name] = f.Id;
        }

        // -------------------------------------------------------------
        // 2. Collect all Views in the document (including templates).
        // -------------------------------------------------------------
        List<View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .ToList();

        // -------------------------------------------------------------
        // 3. Count usage for each filter
        // -------------------------------------------------------------
        Dictionary<ElementId, int> usageCountByFilter = allFilters
            .ToDictionary(f => f.Id, f => 0);

        foreach (View v in allViews)
        {
            ICollection<ElementId> directFilters;
            try
            {
                directFilters = v.GetFilters();
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                // e.g. Schedules, Sheets, Legends do not support V/G overrides
                continue;
            }

            // If the view uses a template, gather that template's filters
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
                    // skip if that template also doesn't support filters
                }
            }

            // Increment usage count
            foreach (ParameterFilterElement f in allFilters)
            {
                bool usedDirectly = directFilters.Contains(f.Id);
                bool usedViaTemplate = templateFilters.Contains(f.Id);

                if (usedDirectly || usedViaTemplate)
                {
                    usageCountByFilter[f.Id]++;
                }
            }
        }

        // -------------------------------------------------------------
        // 4. Build the FIRST grid data
        // -------------------------------------------------------------
        var firstGridData = new List<FilterGridRow>();
        foreach (var f in allFilters)
        {
            firstGridData.Add(new FilterGridRow
            {
                FilterName = f.Name,
                Views      = usageCountByFilter[f.Id]
            });
        }

        // Columns for the first DataGrid
        List<string> firstGridColumns = new List<string> { "FilterName", "Views" };

        // -------------------------------------------------------------
        // 5. Show the first DataGrid (pick filter(s))
        // -------------------------------------------------------------
        List<FilterGridRow> chosenFilters = CustomGUIs.DataGrid(
            entries: firstGridData,
            propertyNames: firstGridColumns,
            Title: "All Filters"
        );

        if (!chosenFilters.Any())
        {
            return Result.Succeeded;
        }

        // Convert chosen filter names -> ElementIds
        List<ElementId> chosenFilterIds = new List<ElementId>();
        foreach (var row in chosenFilters)
        {
            if (filterNameToId.ContainsKey(row.FilterName))
            {
                chosenFilterIds.Add(filterNameToId[row.FilterName]);
            }
        }

        // -------------------------------------------------------------
        // 6. Precompute sheet info for each view (if placed on a sheet)
        //    We'll store sheetId, sheetNumber, sheetName.
        // -------------------------------------------------------------
        Dictionary<ElementId, (int SheetId, string SheetNumber, string SheetName)> viewIdToSheetInfo
            = new Dictionary<ElementId, (int, string, string)>();

        var allViewports = new FilteredElementCollector(doc)
            .OfClass(typeof(Viewport))
            .Cast<Viewport>();

        foreach (Viewport vp in allViewports)
        {
            ElementId vId = vp.ViewId;
            ViewSheet sheet = doc.GetElement(vp.SheetId) as ViewSheet;
            if (sheet != null)
            {
                int sheetIdInt = sheet.Id.IntegerValue;
                string sheetNum = sheet.SheetNumber;
                string sheetNm  = sheet.Name;

                if (!viewIdToSheetInfo.ContainsKey(vId))
                {
                    viewIdToSheetInfo[vId] = (sheetIdInt, sheetNum, sheetNm);
                }
                // If a view can appear on multiple sheets (rare), you might combine them here
            }
        }

        // -------------------------------------------------------------
        // 7. Build the SECOND grid data 
        //    (FilterName, ViewName, ViewType, SheetNumber, SheetName, etc.)
        // -------------------------------------------------------------
        var secondGridData = new List<UsageGridRow>();

        // For quick lookup
        Dictionary<ElementId, string> idToFilterName = allFilters
            .ToDictionary(f => f.Id, f => f.Name);

        foreach (View v in allViews)
        {
            ICollection<ElementId> directFilters;
            try
            {
                directFilters = v.GetFilters();
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                continue;
            }

            ICollection<ElementId> templateFilters = new List<ElementId>();
            if (v.ViewTemplateId != ElementId.InvalidElementId)
            {
                try
                {
                    var tv = doc.GetElement(v.ViewTemplateId) as View;
                    if (tv != null)
                    {
                        templateFilters = tv.GetFilters();
                    }
                }
                catch
                {
                    // skip
                }
            }

            foreach (ElementId fid in chosenFilterIds)
            {
                bool usedDirectly = directFilters.Contains(fid);
                bool usedFromTemplate = templateFilters.Contains(fid);

                if (usedDirectly || usedFromTemplate)
                {
                    var usageRow = new UsageGridRow
                    {
                        FilterName = idToFilterName[fid],
                        ViewName   = v.Name,
                        ViewType   = v.IsTemplate ? "Template" : v.ViewType.ToString(),
                        SheetNumber = "N/A",
                        SheetName   = "N/A",
                        ViewId     = v.Id.IntegerValue,
                        SheetId    = -1
                    };

                    ElementId vid = v.Id;
                    if (viewIdToSheetInfo.ContainsKey(vid))
                    {
                        var sheetInfo = viewIdToSheetInfo[vid];
                        usageRow.SheetId     = sheetInfo.SheetId;
                        usageRow.SheetNumber = sheetInfo.SheetNumber;
                        usageRow.SheetName   = sheetInfo.SheetName;
                    }

                    secondGridData.Add(usageRow);
                }
            }
        }

        if (!secondGridData.Any())
        {
            TaskDialog.Show("Filter Usage", "None of the chosen filters are used on any views.");
            return Result.Succeeded;
        }

        // -------------------------------------------------------------
        // 8. Show the SECOND DataGrid
        //    We'll show: FilterName, ViewName, ViewType, SheetNumber, SheetName
        // -------------------------------------------------------------
        List<string> secondGridColumns = new List<string>
        {
            "FilterName", 
            "ViewName", 
            "ViewType", 
            "SheetNumber", 
            "SheetName"
        };

        List<UsageGridRow> selectedUsage = CustomGUIs.DataGrid(
            entries: secondGridData,
            propertyNames: secondGridColumns,
            Title: "Filter Usage in Views"
        );

        // If user selected rows, call our abstract method to open them
        if (selectedUsage.Any())
        {
            OpenTarget(selectedUsage, doc, uidoc);
        }

        return Result.Succeeded;
    }

    /// <summary>
    /// Derived classes will implement how to open: 
    /// - either the views themselves 
    /// - or the sheets 
    /// </summary>
    protected abstract void OpenTarget(
        List<UsageGridRow> selectedUsage, 
        Document doc, 
        UIDocument uidoc
    );
}

/// <summary>
/// Opens the views themselves.
/// </summary>
[Transaction(TransactionMode.Manual)]
public class OpenViewsByFilters : BaseFilterCommand
{
    protected override void OpenTarget(
        List<UsageGridRow> selectedUsage, 
        Document doc, 
        UIDocument uidoc)
    {
        foreach (UsageGridRow row in selectedUsage)
        {
            ElementId viewId = new ElementId(row.ViewId);
            View theView = doc.GetElement(viewId) as View;
            if (theView != null)
            {
                uidoc.RequestViewChange(theView);
            }
        }
    }
}

/// <summary>
/// Opens the sheets on which the selected views are placed.
/// </summary>
[Transaction(TransactionMode.Manual)]
public class OpenSheetsByFilters : BaseFilterCommand
{
    protected override void OpenTarget(
        List<UsageGridRow> selectedUsage,
        Document doc,
        UIDocument uidoc)
    {
        foreach (UsageGridRow row in selectedUsage)
        {
            // If the view is not on a sheet, skip
            if (row.SheetId < 1)
                continue;

            ElementId sheetEid = new ElementId(row.SheetId);
            ViewSheet sheet = doc.GetElement(sheetEid) as ViewSheet;
            if (sheet != null)
            {
                uidoc.RequestViewChange(sheet);
            }
        }
    }
}
