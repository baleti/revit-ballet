using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenViewsInDocument : IExternalCommand
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
        // DIAGNOSTICS: Start performance monitoring
        // ─────────────────────────────────────────────────────────────
        var overallStopwatch = Stopwatch.StartNew();
        var diagnosticLines = new List<string>();
        diagnosticLines.Add($"=== OpenViewsInDocument Performance Analysis at {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ===");
        diagnosticLines.Add($"Document: {doc.Title}");
        diagnosticLines.Add($"Document Path: {doc.PathName}");
        diagnosticLines.Add("");

        // ─────────────────────────────────────────────────────────────
        // 1. Try to get cached view data (if document state hasn't changed)
        // ─────────────────────────────────────────────────────────────
        var cacheCheckStopwatch = Stopwatch.StartNew();
        List<Dictionary<string, object>> gridData;
        List<string> columns;

        if (ViewDataCache.TryGetDocumentCache(doc, out gridData, out columns))
        {
            // Cache hit! Reconstruct View objects from cached ElementIds
            // This is critical: sync operations invalidate View references but not ElementIds
            foreach (var row in gridData)
            {
                if (row.ContainsKey("ElementIdObject") && row["ElementIdObject"] is ElementId id)
                {
                    Element elem = doc.GetElement(id);
                    if (elem is View view)
                    {
                        row["__OriginalObject"] = view; // Reconstruct fresh View reference
                    }
                }
            }
            cacheCheckStopwatch.Stop();
            diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] Cache HIT - reconstructed {gridData.Count} views in {cacheCheckStopwatch.ElapsedMilliseconds} ms");
        }
        else
        {
            // Cache miss - rebuild view data
            cacheCheckStopwatch.Stop();
            diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] Cache MISS - rebuilding view data");

            // Collect every non-template, non-browser view (incl. sheets)
            var viewCollectionStopwatch = Stopwatch.StartNew();
            List<View> allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v =>
                       !v.IsTemplate &&
                       v.ViewType != ViewType.ProjectBrowser &&
                       v.ViewType != ViewType.SystemBrowser)
                .ToList();
            viewCollectionStopwatch.Stop();
            diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] Collected {allViews.Count} views in {viewCollectionStopwatch.ElapsedMilliseconds} ms");

            // Get browser organization columns (pass views so it can detect sheets vs views)
            var browserColumnsStopwatch = Stopwatch.StartNew();
            List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, allViews);
            browserColumnsStopwatch.Stop();
            diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] Retrieved {browserColumns?.Count ?? 0} browser columns in {browserColumnsStopwatch.ElapsedMilliseconds} ms");

            // ─────────────────────────────────────────────────────────────
            // 2. Collect all viewports ONCE and build lookup dictionary
            // ─────────────────────────────────────────────────────────────
            var viewportCollectionStopwatch = Stopwatch.StartNew();
            var allViewports = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .ToList();

            // Build ViewId → Viewport lookup dictionary for O(1) access
            var viewIdToViewport = new Dictionary<ElementId, Viewport>();
            foreach (var vp in allViewports)
            {
                if (!viewIdToViewport.ContainsKey(vp.ViewId))
                {
                    viewIdToViewport[vp.ViewId] = vp;
                }
            }
            viewportCollectionStopwatch.Stop();
            diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] Collected {allViewports.Count} viewports and built lookup dictionary in {viewportCollectionStopwatch.ElapsedMilliseconds} ms");

            // ─────────────────────────────────────────────────────────────
            // 3. Prepare data for the grid with SheetNumber, Name, ViewType
            // ─────────────────────────────────────────────────────────────
            var viewProcessingStopwatch = Stopwatch.StartNew();
            gridData = new List<Dictionary<string, object>>();

            long totalBrowserColumnsTime = 0;
            long totalViewportLookupTime = 0;
            int viewportLookupsCount = 0;

            foreach (View v in allViews)
            {
                var dict = new Dictionary<string, object>();

                // Add browser organization columns first
                var browserColStopwatch = Stopwatch.StartNew();
                BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, v, doc, browserColumns);
                browserColStopwatch.Stop();
                totalBrowserColumnsTime += browserColStopwatch.ElapsedMilliseconds;

                // Then add standard columns
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

                    // Check if view is placed on a sheet using prebuilt dictionary
                    var viewportLookupStopwatch = Stopwatch.StartNew();
                    Viewport viewport = null;
                    viewIdToViewport.TryGetValue(v.Id, out viewport);
                    viewportLookupStopwatch.Stop();
                    totalViewportLookupTime += viewportLookupStopwatch.ElapsedMilliseconds;
                    viewportLookupsCount++;

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

                // CRITICAL: Do NOT store View object in cache - it becomes invalid after sync
                // Store ElementId only; reconstruct View on cache hit (see above)
                dict["__OriginalObject"] = v; // Store for current use only

                gridData.Add(dict);
            }

            viewProcessingStopwatch.Stop();
            diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] Processed {allViews.Count} views in {viewProcessingStopwatch.ElapsedMilliseconds} ms");
            diagnosticLines.Add($"  - Browser columns: {totalBrowserColumnsTime} ms total");
            diagnosticLines.Add($"  - Viewport lookups: {totalViewportLookupTime} ms total ({viewportLookupsCount} lookups, avg {(viewportLookupsCount > 0 ? totalViewportLookupTime / viewportLookupsCount : 0)} ms each)");

            var sortStopwatch = Stopwatch.StartNew();
            if (browserColumns != null && browserColumns.Count > 0)
            {
                gridData = BrowserOrganizationHelper.SortByBrowserColumns(gridData, browserColumns);
            }
            else
            {
                gridData = gridData.OrderBy(row =>
                {
                    // Extract view to get Title for sorting
                    if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View v)
                        return v.Title;
                    return "";
                }).ToList();
            }
            sortStopwatch.Stop();
            diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] Sorted {gridData.Count} rows in {sortStopwatch.ElapsedMilliseconds} ms");

            // Column headers (order determines column order) - browser columns first
            columns = new List<string>();
            columns.AddRange(browserColumns.Select(bc => bc.Name));
            columns.Add("SheetNumber");
            columns.Add("Name");
            columns.Add("Sheet");

            // Save to cache for next time - remove __OriginalObject before caching
            var cacheSaveStopwatch = Stopwatch.StartNew();
            var cacheData = gridData.Select(row => {
                var cacheRow = new Dictionary<string, object>(row);
                cacheRow.Remove("__OriginalObject"); // Don't cache View objects!
                return cacheRow;
            }).ToList();
            ViewDataCache.SaveDocumentCache(doc, cacheData, columns);
            cacheSaveStopwatch.Stop();
            diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] Saved cache in {cacheSaveStopwatch.ElapsedMilliseconds} ms");
        }

        // ─────────────────────────────────────────────────────────────
        // 5. Figure out which row should be pre-selected (after sorting)
        // ─────────────────────────────────────────────────────────────
        var activeViewDetectionStopwatch = Stopwatch.StartNew();
        int selectedIndex = -1;
        ElementId targetViewId = null;

        if (activeView is ViewSheet)
        {
            targetViewId = activeView.Id;
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
                    targetViewId = containingSheet.Id;
            }

            if (targetViewId == null) // not on a sheet
                targetViewId = activeView.Id;
        }

        // Find the index in sorted gridData
        if (targetViewId != null)
        {
            selectedIndex = gridData.FindIndex(row =>
            {
                if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View view)
                    return view.Id.Equals(targetViewId);
                return false;
            });
        }

        List<int> initialSelectionIndices = selectedIndex >= 0
            ? new List<int> { selectedIndex }
            : new List<int>();

        activeViewDetectionStopwatch.Stop();
        diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] Active view detection completed in {activeViewDetectionStopwatch.ElapsedMilliseconds} ms (selected index: {selectedIndex})");

        // ─────────────────────────────────────────────────────────────
        // 6. Show the grid
        // ─────────────────────────────────────────────────────────────
        diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] Displaying DataGrid with {gridData.Count} rows...");
        var gridDisplayStopwatch = Stopwatch.StartNew();
        CustomGUIs.SetCurrentUIDocument(uidoc);
        List<Dictionary<string, object>> selectedRows =
            CustomGUIs.DataGrid(gridData, columns, false, initialSelectionIndices);
        gridDisplayStopwatch.Stop();
        diagnosticLines.Add($"[{overallStopwatch.ElapsedMilliseconds} ms] DataGrid closed after {gridDisplayStopwatch.ElapsedMilliseconds} ms (user interaction time)");

        // ─────────────────────────────────────────────────────────────
        // 7. Apply any pending edits to Revit elements
        // ─────────────────────────────────────────────────────────────
        bool editsWereApplied = false;
        if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
        {
            CustomGUIs.ApplyCellEditsToEntities();
            editsWereApplied = true;

            // Invalidate cache since edits may have changed view names
            ViewDataCache.InvalidateDocument(doc);
        }

        // ─────────────────────────────────────────────────────────────
        // 8. Open every selected view (sheet or model view)
        //    BUT skip if user made edits (stay in current view instead)
        // ─────────────────────────────────────────────────────────────
        if (!editsWereApplied)
        {
            List<View> selectedViews = CustomGUIs.ExtractOriginalObjects<View>(selectedRows);

            if (selectedViews != null && selectedViews.Any())
            {
                foreach (View view in selectedViews)
                {
                    uidoc.RequestViewChange(view);
                }
            }
        }

        // ─────────────────────────────────────────────────────────────
        // DIAGNOSTICS: Save performance report
        // ─────────────────────────────────────────────────────────────
        overallStopwatch.Stop();
        diagnosticLines.Add("");
        diagnosticLines.Add($"=== TOTAL EXECUTION TIME: {overallStopwatch.ElapsedMilliseconds} ms ===");

        // Diagnostic file writing disabled
        // string diagnosticPath = Path.Combine(
        //     PathHelper.RuntimeDirectory,
        //     "diagnostics",
        //     $"OpenViewsInDocument-{DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");
        //
        // try
        // {
        //     Directory.CreateDirectory(Path.GetDirectoryName(diagnosticPath));
        //     File.WriteAllLines(diagnosticPath, diagnosticLines);
        // }
        // catch
        // {
        //     // Silently ignore diagnostic save failures
        // }

        return Result.Succeeded;
    }
}
