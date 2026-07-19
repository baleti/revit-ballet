using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
[CommandMeta("")]
public class OpenSheetInDocument : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;
        View       activeView = uidoc.ActiveView;

        List<Dictionary<string, object>> gridData;
        List<string> columns;

        if (ViewDataCache.TryGetDocumentCache(doc, "sheets", out gridData, out columns))
        {
            foreach (var row in gridData)
            {
                if (row.ContainsKey("ElementIdObject") && row["ElementIdObject"] is ElementId id)
                {
                    Element elem = doc.GetElement(id);
                    if (elem is ViewSheet sheet)
                        row["__OriginalObject"] = sheet;
                }
            }
        }
        else
        {
            List<ViewSheet> allSheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, allSheets.Cast<View>().ToList());

            gridData = new List<Dictionary<string, object>>();
            foreach (ViewSheet sheet in allSheets)
            {
                var dict = new Dictionary<string, object>();
                BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, sheet, doc, browserColumns);
                dict["SheetNumber"] = sheet.SheetNumber;
                dict["Name"] = sheet.Name;
                dict["ElementIdObject"] = sheet.Id;
                dict["__OriginalObject"] = sheet;
                gridData.Add(dict);
            }

            if (browserColumns != null && browserColumns.Count > 0)
                gridData = BrowserOrganizationHelper.SortByBrowserColumns(gridData, browserColumns, tiebreakerColumn: "SheetNumber");
            else
                gridData = gridData.OrderBy(row => row["SheetNumber"]?.ToString() ?? "").ToList();

            columns = new List<string>();
            columns.AddRange(browserColumns.Select(bc => bc.Name));
            columns.Add("SheetNumber");
            columns.Add("Name");

            var cacheData = gridData.Select(row => {
                var cacheRow = new Dictionary<string, object>(row);
                cacheRow.Remove("__OriginalObject");
                return cacheRow;
            }).ToList();
            ViewDataCache.SaveDocumentCache(doc, "sheets", cacheData, columns);
        }

        int selectedIndex = -1;
        ElementId targetSheetId = null;

        if (activeView is ViewSheet)
        {
            targetSheetId = activeView.Id;
        }
        else
        {
            Viewport vp = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .FirstOrDefault(vpt => vpt.ViewId.Equals(activeView.Id));
            if (vp != null)
            {
                ViewSheet containingSheet = doc.GetElement(vp.SheetId) as ViewSheet;
                if (containingSheet != null)
                    targetSheetId = containingSheet.Id;
            }
        }

        if (targetSheetId != null)
        {
            selectedIndex = gridData.FindIndex(row =>
                row.ContainsKey("__OriginalObject") &&
                row["__OriginalObject"] is ViewSheet s &&
                s.Id.Equals(targetSheetId));
        }

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
            ViewDataCache.InvalidateDocument(doc, "sheets");
        }

        if (!editsWereApplied)
        {
            List<ViewSheet> selectedSheets = CustomGUIs.ExtractOriginalObjects<ViewSheet>(selectedRows);
            if (selectedSheets != null && selectedSheets.Any())
            {
                foreach (ViewSheet sheet in selectedSheets)
                    uidoc.RequestViewChange(sheet);
            }
        }

        return Result.Succeeded;
    }
}
