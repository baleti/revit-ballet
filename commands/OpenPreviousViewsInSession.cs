using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenPreviousViewsInSession : IExternalCommand
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

        // Get SessionId for querying history
        string sessionId = RevitBallet.LogViewChanges.GetSessionId();

        // ─────────────────────────────────────────────────────────────
        // 1. Collect previously opened views from ALL open documents in the session
        // ─────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> gridData = new List<Dictionary<string, object>>();
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

            string documentTitle = doc.Title;

            // Get view history from database for this document (sorted by timestamp DESC)
            var history = LogViewChangesDatabase.GetViewHistoryForDocument(sessionId, documentTitle, limit: 1000);

            if (history.Count == 0)
                continue;

            // Build list of views from history
            var previousViews = new List<View>();
            var seenViewIds = new HashSet<ElementId>();
            var viewTimestamps = new Dictionary<ElementId, DateTime>();

            foreach (var entry in history)
            {
                ElementId viewId;
                try
                {
                    viewId = entry.ViewId.ToElementId();
                }
                catch (Exception)
                {
                    // Skip invalid ViewId entries (e.g., 0, -1, or corrupted data)
                    continue;
                }

                // Skip duplicates
                if (seenViewIds.Contains(viewId))
                    continue;

                Element viewElement = doc.GetElement(viewId);

                // If element not found by ID, try finding by title
                if (viewElement == null || !(viewElement is View))
                {
                    viewElement = new FilteredElementCollector(doc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .FirstOrDefault(v => v.Title == entry.ViewTitle);
                }

                if (viewElement != null && viewElement is View view &&
                    !view.IsTemplate &&
                    view.ViewType != ViewType.ProjectBrowser &&
                    view.ViewType != ViewType.SystemBrowser)
                {
                    previousViews.Add(view);
                    seenViewIds.Add(view.Id);
                    viewTimestamps[view.Id] = entry.Timestamp;
                }
            }

            if (previousViews.Count == 0)
                continue;

            // Get browser organization columns for this document's views
            List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, previousViews);
            browserColumnsByDoc[doc] = browserColumns;

            // Add each view to the grid data
            foreach (View v in previousViews)
            {
                var dict = new Dictionary<string, object>();

                // Add document name first
                dict["Document"] = documentTitle;

                // Add browser organization columns
                BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, v, doc, browserColumns);

                // Add timestamp column
                if (viewTimestamps.TryGetValue(v.Id, out DateTime timestamp))
                {
                    dict["Last Opened"] = timestamp.ToString("yyyy-MM-dd HH:mm:ss");
                    dict["__Timestamp"] = timestamp; // Store for sorting
                }
                else
                {
                    dict["Last Opened"] = "";
                    dict["__Timestamp"] = DateTime.MinValue;
                }

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
            TaskDialog.Show("Info", "No previously opened views found in any open documents.");
            return Result.Failed;
        }

        // ─────────────────────────────────────────────────────────────
        // 2. Sort by timestamp descending (most recently opened first)
        // ─────────────────────────────────────────────────────────────
        gridData = gridData.OrderByDescending(row =>
        {
            if (row.ContainsKey("__Timestamp"))
                return (DateTime)row["__Timestamp"];
            return DateTime.MinValue;
        }).ToList();

        // ─────────────────────────────────────────────────────────────
        // 3. Build column headers (Document first, then browser columns, then timestamp, then standard)
        // ─────────────────────────────────────────────────────────────
        List<string> columns = new List<string>();
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

        columns.Add("Last Opened");
        columns.Add("SheetNumber");
        columns.Add("Name");
        columns.Add("Sheet");

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
                    return view.Id == targetViewId && doc == activeDoc;
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
        if (CustomGUIs.HasPendingEdits())
        {
            CustomGUIs.ApplyCellEditsToEntities();
            editsWereApplied = true;
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
                if (viewDoc != currentDoc)
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
                    catch (Exception ex)
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
