using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenSheetsInDocument : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;
        View       activeView = uidoc.ActiveView;

        // ─────────────────────────────────────────────────────────────
        // 1. Try to get cached sheet data (if document state hasn't changed)
        // ─────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> gridData;
        List<string> columns;

        if (ViewDataCache.TryGetDocumentCache(doc, "sheets", out gridData, out columns))
        {
            // Cache hit! Skip expensive data collection
            // Note: gridData still contains __OriginalObject references from cache
        }
        else
        {
            // Cache miss - rebuild sheet data

            // Collect only ViewSheet views (sheets only)
            List<ViewSheet> allSheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            // Get browser organization columns (pass sheets as views for consistency)
            List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, allSheets.Cast<View>().ToList());

            // ─────────────────────────────────────────────────────────────
            // 2. Prepare data for the grid with SheetNumber and Name
            // ─────────────────────────────────────────────────────────────
            gridData = new List<Dictionary<string, object>>();

            foreach (ViewSheet sheet in allSheets)
            {
                var dict = new Dictionary<string, object>();

                // Add browser organization columns first
                BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, sheet, doc, browserColumns);

                // Then add standard columns
                dict["SheetNumber"] = sheet.SheetNumber;
                dict["Name"] = sheet.Name;

                dict["ElementIdObject"] = sheet.Id; // Required for edit functionality
                dict["__OriginalObject"] = sheet; // Store original object for extraction

                gridData.Add(dict);
            }

            // Sort by browser organization columns (if any), otherwise by SheetNumber
            if (browserColumns != null && browserColumns.Count > 0)
            {
                gridData = BrowserOrganizationHelper.SortByBrowserColumns(gridData, browserColumns, tiebreakerColumn: "SheetNumber");
            }
            else
            {
                gridData = gridData.OrderBy(row =>
                {
                    // Extract sheet number for sorting
                    if (row.ContainsKey("SheetNumber"))
                        return row["SheetNumber"]?.ToString() ?? "";
                    return "";
                }).ToList();
            }

            // Column headers (order determines column order) - browser columns first
            columns = new List<string>();
            columns.AddRange(browserColumns.Select(bc => bc.Name));
            columns.Add("SheetNumber");
            columns.Add("Name");

            // Save to cache for next time
            ViewDataCache.SaveDocumentCache(doc, "sheets", gridData, columns);
        }

        // ─────────────────────────────────────────────────────────────
        // 3. Figure out which row should be pre-selected (after sorting)
        // ─────────────────────────────────────────────────────────────
        int selectedIndex = -1;
        ElementId targetSheetId = null;

        if (activeView is ViewSheet)
        {
            targetSheetId = activeView.Id;
        }
        else
        {
            // If we're in a non-sheet view, find the sheet it's placed on
            Viewport vp = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .FirstOrDefault(vpt => vpt.ViewId == activeView.Id);

            if (vp != null)
            {
                ViewSheet containingSheet = doc.GetElement(vp.SheetId) as ViewSheet;
                if (containingSheet != null)
                    targetSheetId = containingSheet.Id;
            }
        }

        // Find the index in sorted gridData
        if (targetSheetId != null)
        {
            selectedIndex = gridData.FindIndex(row =>
            {
                if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is ViewSheet sheet)
                    return sheet.Id == targetSheetId;
                return false;
            });
        }

        List<int> initialSelectionIndices = selectedIndex >= 0
            ? new List<int> { selectedIndex }
            : new List<int>();

        // ─────────────────────────────────────────────────────────────
        // 4. Show the grid
        // ─────────────────────────────────────────────────────────────
        CustomGUIs.SetCurrentUIDocument(uidoc);
        List<Dictionary<string, object>> selectedRows =
            CustomGUIs.DataGrid(gridData, columns, false, initialSelectionIndices);

        // ─────────────────────────────────────────────────────────────
        // 5. Apply any pending edits to Revit elements
        // ─────────────────────────────────────────────────────────────
        bool editsWereApplied = false;
        if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
        {
            CustomGUIs.ApplyCellEditsToEntities();
            editsWereApplied = true;

            // Invalidate cache since edits may have changed sheet names
            ViewDataCache.InvalidateDocument(doc, "sheets");
        }

        // ─────────────────────────────────────────────────────────────
        // 6. Open every selected sheet
        //    BUT skip if user made edits (stay in current view instead)
        // ─────────────────────────────────────────────────────────────
        if (!editsWereApplied)
        {
            List<ViewSheet> selectedSheets = CustomGUIs.ExtractOriginalObjects<ViewSheet>(selectedRows);

            if (selectedSheets != null && selectedSheets.Any())
            {
                foreach (ViewSheet sheet in selectedSheets)
                {
                    uidoc.RequestViewChange(sheet);
                }
            }
        }

        return Result.Succeeded;
    }
}
