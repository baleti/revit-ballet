using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenSheetsInSession : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document activeDoc = uidoc?.Document;

        if (activeDoc == null)
        {
            TaskDialog.Show("Error", "No active document.");
            return Result.Failed;
        }

        View activeView = uidoc.ActiveView;

        // ─────────────────────────────────────────────────────────────
        // 1. Try to get cached session data (if no documents have changed)
        // ─────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> gridData;
        List<string> columns;

        if (ViewDataCache.TryGetSessionCache(uiApp.Application, "sheets", out gridData, out columns))
        {
            // Cache hit! Skip expensive data collection from all documents
            // Note: gridData still contains __OriginalObject and __Document references
        }
        else
        {
            // Cache miss - rebuild sheet data from all documents

            gridData = new List<Dictionary<string, object>>();
            Dictionary<Document, List<BrowserOrganizationHelper.BrowserColumn>> browserColumnsByDoc
                = new Dictionary<Document, List<BrowserOrganizationHelper.BrowserColumn>>();

            foreach (Document doc in uiApp.Application.Documents)
            {
                // Skip linked documents (read-only references)
                if (doc.IsLinked)
                    continue;

                // Skip family documents
                if (doc.IsFamilyDocument)
                    continue;

                // Collect only ViewSheet views (sheets only) in this document
                List<ViewSheet> sheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .ToList();

                // Get browser organization columns for this document's sheets
                List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                    BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, sheets.Cast<View>().ToList());
                browserColumnsByDoc[doc] = browserColumns;

                // Add each sheet to the grid data
                foreach (ViewSheet sheet in sheets)
                {
                    var dict = new Dictionary<string, object>();

                    // Add document name first
                    dict["Document"] = doc.Title;

                    // Add browser organization columns
                    BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, sheet, doc, browserColumns);

                    // Add standard columns
                    dict["SheetNumber"] = sheet.SheetNumber;
                    dict["Name"] = sheet.Name;

                    dict["ElementIdObject"] = sheet.Id; // Required for edit functionality
                    dict["__OriginalObject"] = sheet; // Store original object for extraction
                    dict["__Document"] = doc; // Store document reference for switching

                    gridData.Add(dict);
                }
            }

            if (gridData.Count == 0)
            {
                TaskDialog.Show("Info", "No sheets found in any open documents.");
                return Result.Failed;
            }

            // ─────────────────────────────────────────────────────────────
            // 2. Sort by Document first, then by browser columns or Title
            //    (same approach as OpenViewsInSession - using the helper method)
            // ─────────────────────────────────────────────────────────────
            // Group by document and sort each group separately
            var groupedByDocument = gridData.GroupBy(row => row["Document"]?.ToString() ?? "").ToList();
            gridData.Clear();

            foreach (var docGroup in groupedByDocument.OrderBy(g => g.Key))
            {
                var sheetsInDoc = docGroup.ToList();

                // Get the document from the first row to retrieve browser columns
                Document doc = sheetsInDoc.First()["__Document"] as Document;

                if (doc != null && browserColumnsByDoc.TryGetValue(doc, out var browserColumns) &&
                    browserColumns != null && browserColumns.Count > 0)
                {
                    // Sort by browser columns using the helper method with SheetNumber as tiebreaker
                    sheetsInDoc = BrowserOrganizationHelper.SortByBrowserColumns(sheetsInDoc, browserColumns, tiebreakerColumn: "SheetNumber");
                }
                else
                {
                    // Fallback: sort by sheet number
                    sheetsInDoc = sheetsInDoc.OrderBy(row =>
                    {
                        if (row.ContainsKey("SheetNumber"))
                            return row["SheetNumber"]?.ToString() ?? "";
                        return "";
                    }).ToList();
                }

                gridData.AddRange(sheetsInDoc);
            }

            // ─────────────────────────────────────────────────────────────
            // 3. Build column headers (Document first, then browser columns, then standard)
            // ─────────────────────────────────────────────────────────────
            columns = new List<string>();
            columns.Add("Document");

            // Add browser columns (union of all documents' browser columns)
            HashSet<string> allBrowserColumnNames = new HashSet<string>();
            foreach (var browserColumns in browserColumnsByDoc.Values)
            {
                if (browserColumns != null)
                {
                    foreach (var bc in browserColumns)
                    {
                        allBrowserColumnNames.Add(bc.Name);
                    }
                }
            }
            columns.AddRange(allBrowserColumnNames.OrderBy(n => n));

            columns.Add("SheetNumber");
            columns.Add("Name");

            // Save to session cache for next time
            ViewDataCache.SaveSessionCache(uiApp.Application, "sheets", gridData, columns);
        }

        // ─────────────────────────────────────────────────────────────
        // 4. Figure out which row should be pre-selected (active sheet in active doc)
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
            Viewport vp = new FilteredElementCollector(activeDoc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .FirstOrDefault(vpt => vpt.ViewId == activeView.Id);

            if (vp != null)
            {
                ViewSheet containingSheet = activeDoc.GetElement(vp.SheetId) as ViewSheet;
                if (containingSheet != null)
                    targetSheetId = containingSheet.Id;
            }
        }

        // Find the index in sorted gridData (matching both sheet and document)
        if (targetSheetId != null)
        {
            selectedIndex = gridData.FindIndex(row =>
            {
                if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is ViewSheet sheet &&
                    row.ContainsKey("__Document") && row["__Document"] is Document doc)
                {
                    return sheet.Id == targetSheetId && doc.Equals(activeDoc);
                }
                return false;
            });
        }

        List<int> initialSelectionIndices = selectedIndex >= 0
            ? new List<int> { selectedIndex }
            : new List<int>();

        // ─────────────────────────────────────────────────────────────
        // 5. Show the grid
        // ─────────────────────────────────────────────────────────────
        CustomGUIs.SetCurrentUIDocument(uidoc);
        List<Dictionary<string, object>> selectedRows =
            CustomGUIs.DataGrid(gridData, columns, false, initialSelectionIndices);

        // ─────────────────────────────────────────────────────────────
        // 6. Check if edits were applied (DataGrid auto-applies on close)
        // ─────────────────────────────────────────────────────────────
        bool editsWereApplied = CustomGUIs.WereEditsApplied();
        if (editsWereApplied)
        {
            // Invalidate session cache since edits may have changed sheet names
            ViewDataCache.InvalidateAll("sheets");
        }

        // ─────────────────────────────────────────────────────────────
        // 7. Open every selected sheet (switching documents if needed)
        //    BUT skip if user made edits (stay in current view instead)
        // ─────────────────────────────────────────────────────────────
        if (!editsWereApplied && selectedRows != null && selectedRows.Any())
        {
            Document currentDoc = activeDoc;
            UIDocument currentUidoc = uidoc;

            foreach (var row in selectedRows)
            {
                ViewSheet sheet = row["__OriginalObject"] as ViewSheet;
                Document sheetDoc = row["__Document"] as Document;

                if (sheet == null || sheetDoc == null)
                    continue;

                // Switch document if needed
                if (!sheetDoc.Equals(currentDoc))
                {
                    // Check if document is saved (required for programmatic switching)
                    if (string.IsNullOrEmpty(sheetDoc.PathName))
                    {
                        TaskDialog.Show("Unsaved Document",
                            $"Cannot open sheets in unsaved document '{sheetDoc.Title}'.\n\n" +
                            $"Please save the document first.");
                        continue;
                    }

                    try
                    {
                        currentUidoc = uiApp.OpenAndActivateDocument(sheetDoc.PathName);
                        currentDoc = sheetDoc;
                    }
                    catch (System.Exception ex)
                    {
                        TaskDialog.Show("Error",
                            $"Failed to switch to document '{sheetDoc.Title}':\n\n{ex.Message}");
                        continue;
                    }
                }

                // Open the sheet in the current document
                currentUidoc.RequestViewChange(sheet);
            }
        }

        return Result.Succeeded;
    }
}
