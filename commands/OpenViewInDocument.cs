using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
[CommandMeta("View")]
public class OpenViewInDocument : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;

        // Selection-first: if views or viewports are selected, open them immediately
        var selectedIds = uidoc.GetSelectionIds();
        if (selectedIds.Count > 0)
        {
            var viewsFromSelection = new List<View>();
            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is View v &&
                    !(v is ViewSheet) &&
                    !v.IsTemplate &&
                    v.ViewType != ViewType.ProjectBrowser &&
                    v.ViewType != ViewType.SystemBrowser)
                {
                    viewsFromSelection.Add(v);
                }
                else if (elem is Viewport vp)
                {
                    var viewFromVp = doc.GetElement(vp.ViewId) as View;
                    if (viewFromVp != null && !(viewFromVp is ViewSheet))
                        viewsFromSelection.Add(viewFromVp);
                }
            }
            if (viewsFromSelection.Count > 0)
            {
                foreach (View view in viewsFromSelection)
                    uidoc.RequestViewChange(view);
                return Result.Succeeded;
            }
        }

        // No views selected — show picker
        List<Dictionary<string, object>> gridData;
        List<string> columns;

        if (ViewDataCache.TryGetDocumentCache(doc, "views", out gridData, out columns))
        {
            foreach (var row in gridData)
            {
                if (row.ContainsKey("ElementIdObject") && row["ElementIdObject"] is ElementId id)
                {
                    if (doc.GetElement(id) is View view)
                        row["__OriginalObject"] = view;
                }
            }
        }
        else
        {
            List<View> allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v =>
                    !(v is ViewSheet) &&
                    !v.IsTemplate &&
                    v.ViewType != ViewType.ProjectBrowser &&
                    v.ViewType != ViewType.SystemBrowser)
                .ToList();

            List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, allViews);

            var viewIdToViewport = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .GroupBy(vp => vp.ViewId)
                .ToDictionary(g => g.Key, g => g.First());

            gridData = new List<Dictionary<string, object>>();
            foreach (View v in allViews)
            {
                var dict = new Dictionary<string, object>();
                BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, v, doc, browserColumns);
                dict["Name"] = v.Name;

                Viewport viewport;
                if (viewIdToViewport.TryGetValue(v.Id, out viewport))
                {
                    ViewSheet containingSheet = doc.GetElement(viewport.SheetId) as ViewSheet;
                    dict["Sheet Number"] = containingSheet != null ? containingSheet.SheetNumber : "";
                    dict["Sheet Title"] = containingSheet != null ? containingSheet.Name : "";
                }
                else
                {
                    dict["Sheet Number"] = "";
                    dict["Sheet Title"] = "";
                }

                dict["ElementIdObject"] = v.Id;
                dict["__OriginalObject"] = v;
                gridData.Add(dict);
            }

            if (browserColumns != null && browserColumns.Count > 0)
                gridData = BrowserOrganizationHelper.SortByBrowserColumns(gridData, browserColumns);
            else
                gridData = gridData.OrderBy(row =>
                    (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View sv)
                        ? sv.Title : "").ToList();

            columns = new List<string>();
            if (browserColumns != null)
                columns.AddRange(browserColumns.Select(bc => bc.Name));
            columns.Add("Name");
            columns.Add("Sheet Number");
            columns.Add("Sheet Title");

            var cacheData = gridData.Select(row => {
                var r = new Dictionary<string, object>(row);
                r.Remove("__OriginalObject");
                return r;
            }).ToList();
            ViewDataCache.SaveDocumentCache(doc, "views", cacheData, columns);
        }

        // Pre-select the active view if it is not a sheet
        ElementId targetViewId = !(activeView is ViewSheet) ? activeView.Id : null;
        int selectedIndex = targetViewId != null
            ? gridData.FindIndex(row =>
                row.ContainsKey("__OriginalObject") &&
                row["__OriginalObject"] is View rv &&
                rv.Id.Equals(targetViewId))
            : -1;

        List<int> initialSelectionIndices = selectedIndex >= 0
            ? new List<int> { selectedIndex }
            : new List<int>();

        CustomGUIs.SetCurrentUIDocument(uidoc);
        List<Dictionary<string, object>> selectedRows =
            CustomGUIs.DataGrid(gridData, columns, false, initialSelectionIndices);

        bool editsWereApplied = false;
        if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
        {
            CustomGUIs.ApplyCellEditsToEntities();
            editsWereApplied = true;
            ViewDataCache.InvalidateDocument(doc, "views");
        }

        if (!editsWereApplied)
        {
            List<View> selectedViews = CustomGUIs.ExtractOriginalObjects<View>(selectedRows);
            if (selectedViews != null)
            {
                foreach (View view in selectedViews)
                    uidoc.RequestViewChange(view);
            }
        }

        return Result.Succeeded;
    }
}
