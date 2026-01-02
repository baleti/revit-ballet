using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenViewsInSession : IExternalCommand
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

        if (ViewDataCache.TryGetSessionCache(uiApp.Application, out gridData, out columns))
        {
            // Cache hit! Skip expensive data collection from all documents
            // Note: gridData still contains __OriginalObject and __Document references
        }
        else
        {
            // Cache miss - rebuild view data from all documents

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

                // Collect all non-template, non-browser views in this document
                List<View> views = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v =>
                           !v.IsTemplate &&
                           v.ViewType != ViewType.ProjectBrowser &&
                           v.ViewType != ViewType.SystemBrowser)
                    .ToList();

                // Get browser organization columns for this document's views
                List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                    BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, views);
                browserColumnsByDoc[doc] = browserColumns;

                // Add each view to the grid data
                foreach (View v in views)
                {
                    var dict = new Dictionary<string, object>();

                    // Add document name first
                    dict["Document"] = doc.Title;

                    // Add browser organization columns
                    BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, v, doc, browserColumns);

                    // Add standard columns
                    if (v is ViewSheet sheet)
                    {
                        dict["SheetNumber"] = sheet.SheetNumber;
                        dict["Name"] = sheet.Name;
                        dict["Sheet"] = ""; // Empty for sheets
                    }
                    else
                    {
                        dict["SheetNumber"] = ""; // Empty for non-sheet views
                        dict["Name"] = v.Name;

                        // Check if view is placed on a sheet
                        var viewport = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .FirstOrDefault(vp => vp.ViewId == v.Id);

                        if (viewport != null)
                        {
                            ViewSheet containingSheet = doc.GetElement(viewport.SheetId) as ViewSheet;
                            dict["Sheet"] = containingSheet != null ? containingSheet.Title : "";
                        }
                        else
                        {
                            dict["Sheet"] = ""; // Empty for views not on sheets
                        }
                    }

                    dict["ElementIdObject"] = v.Id; // Required for edit functionality
                    dict["__OriginalObject"] = v; // Store original object for extraction
                    dict["__Document"] = doc; // Store document reference for switching

                    gridData.Add(dict);
                }
            }

            if (gridData.Count == 0)
            {
                TaskDialog.Show("Info", "No views found in any open documents.");
                return Result.Failed;
            }

            // ─────────────────────────────────────────────────────────────
            // 2. Sort by Document first, then by browser columns or Title
            //    (same approach as OpenViewsInDocument - using the helper method)
            // ─────────────────────────────────────────────────────────────
            // Group by document and sort each group separately
            var groupedByDocument = gridData.GroupBy(row => row["Document"]?.ToString() ?? "").ToList();
            gridData.Clear();

            foreach (var docGroup in groupedByDocument.OrderBy(g => g.Key))
            {
                var viewsInDoc = docGroup.ToList();

                // Get the document from the first row to retrieve browser columns
                Document doc = viewsInDoc.First()["__Document"] as Document;

                if (doc != null && browserColumnsByDoc.TryGetValue(doc, out var browserColumns) &&
                    browserColumns != null && browserColumns.Count > 0)
                {
                    // Sort by browser columns using the helper method (same as OpenViewsInDocument)
                    viewsInDoc = BrowserOrganizationHelper.SortByBrowserColumns(viewsInDoc, browserColumns);
                }
                else
                {
                    // Fallback: sort by view Title
                    viewsInDoc = viewsInDoc.OrderBy(row =>
                    {
                        if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View v)
                            return v.Title;
                        return "";
                    }).ToList();
                }

                gridData.AddRange(viewsInDoc);
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
            columns.Add("Sheet");

            // Save to session cache for next time
            ViewDataCache.SaveSessionCache(uiApp.Application, gridData, columns);
        }

        // ─────────────────────────────────────────────────────────────
        // 4. Figure out which row should be pre-selected (active view in active doc)
        // ─────────────────────────────────────────────────────────────
        int selectedIndex = -1;
        ElementId targetViewId = null;

        if (activeView is ViewSheet)
        {
            targetViewId = activeView.Id;
        }
        else
        {
            Viewport vp = new FilteredElementCollector(activeDoc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .FirstOrDefault(vpt => vpt.ViewId == activeView.Id);

            if (vp != null)
            {
                ViewSheet containingSheet = activeDoc.GetElement(vp.SheetId) as ViewSheet;
                if (containingSheet != null)
                    targetViewId = containingSheet.Id;
            }

            if (targetViewId == null) // not on a sheet
                targetViewId = activeView.Id;
        }

        // Find the index in sorted gridData (matching both view and document)
        if (targetViewId != null)
        {
            selectedIndex = gridData.FindIndex(row =>
            {
                if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View view &&
                    row.ContainsKey("__Document") && row["__Document"] is Document doc)
                {
                    return view.Id == targetViewId && doc.Equals(activeDoc);
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
        // 6. Apply any pending edits to Revit elements
        // ─────────────────────────────────────────────────────────────
        bool editsWereApplied = false;
        if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
        {
            CustomGUIs.ApplyCellEditsToEntities();
            editsWereApplied = true;

            // Invalidate session cache since edits may have changed view names
            ViewDataCache.InvalidateAll();
        }

        // ─────────────────────────────────────────────────────────────
        // 7. Open every selected view (switching documents if needed)
        //    BUT skip if user made edits (stay in current view instead)
        // ─────────────────────────────────────────────────────────────
        if (!editsWereApplied && selectedRows != null && selectedRows.Any())
        {
            Document currentDoc = activeDoc;
            UIDocument currentUidoc = uidoc;

            foreach (var row in selectedRows)
            {
                View view = row["__OriginalObject"] as View;
                Document viewDoc = row["__Document"] as Document;

                if (view == null || viewDoc == null)
                    continue;

                // Switch document if needed
                if (!viewDoc.Equals(currentDoc))
                {
                    // Check if document is saved (required for programmatic switching)
                    if (string.IsNullOrEmpty(viewDoc.PathName))
                    {
                        TaskDialog.Show("Unsaved Document",
                            $"Cannot open views in unsaved document '{viewDoc.Title}'.\n\n" +
                            $"Please save the document first.");
                        continue;
                    }

                    try
                    {
                        currentUidoc = uiApp.OpenAndActivateDocument(viewDoc.PathName);
                        currentDoc = viewDoc;
                    }
                    catch (System.Exception ex)
                    {
                        TaskDialog.Show("Error",
                            $"Failed to switch to document '{viewDoc.Title}':\n\n{ex.Message}");
                        continue;
                    }
                }

                // Open the view in the current document
                currentUidoc.RequestViewChange(view);
            }
        }

        return Result.Succeeded;
    }
}
