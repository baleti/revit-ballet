using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CloseViewsInSession : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument activeUidoc = uiApp.ActiveUIDocument;
        Document activeDoc = activeUidoc?.Document;

        if (activeDoc == null)
        {
            TaskDialog.Show("Error", "No active document.");
            return Result.Failed;
        }

        // Get SessionId for database operations
        string sessionId = RevitBallet.LogViewChanges.GetSessionId();

        // ─────────────────────────────────────────────────────────────
        // 1. Collect open views from ALL open documents in the session
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
            bool isActiveDoc = doc.Equals(activeDoc);

            // Get currently open view IDs
            var openViewIds = new HashSet<ElementId>();

            if (isActiveDoc)
            {
                // For active document, use GetOpenUIViews() for accuracy
                IList<UIView> openUIViews = activeUidoc.GetOpenUIViews();
                foreach (UIView uiView in openUIViews)
                {
                    openViewIds.Add(uiView.ViewId);
                }
            }
            else
            {
                // For non-active documents, use view history database
                var viewHistory = LogViewChangesDatabase.GetViewHistoryForDocument(sessionId, doc.Title, limit: 1000);
                foreach (var entry in viewHistory)
                {
                    try
                    {
                        openViewIds.Add(entry.ViewId.ToElementId());
                    }
                    catch (Exception)
                    {
                        // Skip invalid ViewId entries (e.g., 0, -1, or corrupted data)
                        continue;
                    }
                }
            }

            // Skip this document if no open views found
            if (openViewIds.Count == 0)
                continue;

            // Get the View objects for open views
            var openViews = new List<View>();
            foreach (ElementId viewId in openViewIds)
            {
                View view = doc.GetElement(viewId) as View;
                if (view != null && !view.IsTemplate &&
                    view.ViewType != ViewType.ProjectBrowser &&
                    view.ViewType != ViewType.SystemBrowser)
                {
                    openViews.Add(view);
                }
            }

            if (openViews.Count == 0)
                continue;

            // Get browser organization columns for this document
            List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, openViews);
            browserColumnsByDoc[doc] = browserColumns;

            // Add each view to the combined list
            foreach (View v in openViews)
            {
                var dict = new Dictionary<string, object>();

                // Add document name as first column
                dict["Document"] = documentTitle;

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
                    dict["SheetNumber"] = "";
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

                dict["ViewType"] = v.ViewType;
                dict["ElementIdObject"] = v.Id;
                dict["__OriginalObject"] = v;
                dict["__Document"] = doc; // Store document reference

                gridData.Add(dict);
            }
        }

        if (gridData.Count == 0)
        {
            TaskDialog.Show("Info", "No open views found in any documents.");
            return Result.Failed;
        }

        // ─────────────────────────────────────────────────────────────
        // 2. Sort by Document first, then by browser columns
        //    (same approach as OpenViewsInSession)
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
                // Sort by browser columns using the helper method
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

        columns.Add("SheetNumber");
        columns.Add("Name");
        columns.Add("ViewType");
        columns.Add("Sheet");

        // ─────────────────────────────────────────────────────────────
        // 4. Show the grid
        // ─────────────────────────────────────────────────────────────
        CustomGUIs.SetCurrentUIDocument(activeUidoc);
        var selectedDicts = CustomGUIs.DataGrid(gridData, columns, false);

        if (selectedDicts == null || selectedDicts.Count == 0)
            return Result.Failed;

        // ─────────────────────────────────────────────────────────────
        // 5. Close selected views (handling cross-document operations)
        // ─────────────────────────────────────────────────────────────
        foreach (var row in selectedDicts)
        {
            View view = row["__OriginalObject"] as View;
            Document viewDoc = row["__Document"] as Document;

            if (view == null || viewDoc == null)
                continue;

            // Close the view
            if (viewDoc.Equals(activeDoc))
            {
                // Same document - can close directly
                foreach (UIView openedUIView in activeUidoc.GetOpenUIViews())
                {
                    if (openedUIView.ViewId.Equals(view.Id))
                    {
                        // Remove the closed view from history
                        try
                        {
                            LogViewChangesDatabase.RemoveViewFromHistory(sessionId, viewDoc.Title, view.Title);
                        }
                        catch
                        {
                            // Silently fail - don't interrupt the close operation
                        }

                        openedUIView.Close();
                        break;
                    }
                }
            }
            else
            {
                // Different document - just remove from history
                // (Cannot programmatically close views in non-active documents)
                try
                {
                    LogViewChangesDatabase.RemoveViewFromHistory(sessionId, viewDoc.Title, view.Title);
                }
                catch
                {
                    // Silently fail
                }
            }
        }

        return Result.Succeeded;
    }
}
