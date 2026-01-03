using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SwitchViewInSession : IExternalCommand
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

        // Collect views from all open documents
        var allViewsData = new List<Dictionary<string, object>>();

        // Iterate through all open documents in the session
        foreach (Document doc in uiApp.Application.Documents)
        {
            // Skip linked documents (read-only references)
            if (doc.IsLinked)
                continue;

            // Skip family documents
            if (doc.IsFamilyDocument)
                continue;

            string projectName = doc.Title;
            string sessionId = RevitBallet.LogViewChanges.GetSessionId();
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
                // The database tracks open views via LogViewActivation (when opened)
                // and RemoveViewFromHistory (when closed via CloseViews command)
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
            var viewsInDoc = new List<View>();
            foreach (ElementId viewId in openViewIds)
            {
                View view = doc.GetElement(viewId) as View;
                if (view != null && !view.IsTemplate &&
                    view.ViewType != ViewType.ProjectBrowser &&
                    view.ViewType != ViewType.SystemBrowser)
                {
                    viewsInDoc.Add(view);
                }
            }

            // Get browser organization columns for this document
            List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                BrowserOrganizationHelper.GetBrowserColumnsForViews(doc);

            // Add each view to the combined list
            foreach (View view in viewsInDoc)
            {
                var dict = new Dictionary<string, object>();

                // Add document name as first column
                dict["Document"] = projectName;

                // Add browser organization columns
                BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, view, doc, browserColumns);

                // Add standard columns
                if (view is ViewSheet sheet)
                {
                    dict["SheetNumber"] = sheet.SheetNumber;
                    dict["Name"] = sheet.Name;
                    dict["Sheet"] = ""; // Empty for sheets
                }
                else
                {
                    dict["SheetNumber"] = "";
                    dict["Name"] = view.Name;

                    // Check if view is placed on a sheet
                    var viewport = new FilteredElementCollector(doc)
                        .OfClass(typeof(Viewport))
                        .Cast<Viewport>()
                        .FirstOrDefault(vp => vp.ViewId == view.Id);

                    if (viewport != null)
                    {
                        ViewSheet containingSheet = doc.GetElement(viewport.SheetId) as ViewSheet;
                        dict["Sheet"] = containingSheet != null ? $"{containingSheet.SheetNumber} - {containingSheet.Name}" : "";
                    }
                    else
                    {
                        dict["Sheet"] = ""; // Empty for views not on sheets
                    }
                }

                dict["ViewType"] = view.ViewType;
                dict["ElementIdObject"] = view.Id;
                dict["__OriginalObject"] = view;
                dict["__Document"] = doc; // Store document reference

                allViewsData.Add(dict);
            }
        }

        if (allViewsData.Count == 0)
        {
            // No views found in any document history - documents just opened
            // Fallback: show only the currently active view from the active document
            View fallbackView = activeUidoc.ActiveView;

            // Get browser organization columns for active document
            List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                BrowserOrganizationHelper.GetBrowserColumnsForViews(activeDoc);

            var dict = new Dictionary<string, object>();
            dict["Document"] = activeDoc.Title;

            // Add browser organization columns
            BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, fallbackView, activeDoc, browserColumns);

            // Add standard columns
            if (fallbackView is ViewSheet sheet)
            {
                dict["SheetNumber"] = sheet.SheetNumber;
                dict["Name"] = sheet.Name;
                dict["Sheet"] = ""; // Empty for sheets
            }
            else
            {
                dict["SheetNumber"] = "";
                dict["Name"] = fallbackView.Name;

                // Check if view is placed on a sheet
                var viewport = new FilteredElementCollector(activeDoc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .FirstOrDefault(vp => vp.ViewId == fallbackView.Id);

                if (viewport != null)
                {
                    ViewSheet containingSheet = activeDoc.GetElement(viewport.SheetId) as ViewSheet;
                    dict["Sheet"] = containingSheet != null ? $"{containingSheet.SheetNumber} - {containingSheet.Name}" : "";
                }
                else
                {
                    dict["Sheet"] = ""; // Empty for views not on sheets
                }
            }

            dict["ViewType"] = fallbackView.ViewType;
            dict["ElementIdObject"] = fallbackView.Id;
            dict["__OriginalObject"] = fallbackView;
            dict["__Document"] = activeDoc;

            allViewsData.Add(dict);
        }

        // Get browser columns from all views (all should have same structure per document)
        var allBrowserColumnNames = allViewsData
            .SelectMany(d => d.Keys)
            .Where(k => k != "Document" && k != "SheetNumber" && k != "Name" &&
                       k != "ViewType" && k != "ElementIdObject" &&
                       k != "__OriginalObject" && k != "__Document")
            .Distinct()
            .ToList();

        // Sort by Document first, then by browser organization columns
        // Group by document and sort each group separately
        var groupedByDocument = allViewsData.GroupBy(row => row["Document"]?.ToString() ?? "").ToList();
        allViewsData.Clear();

        foreach (var docGroup in groupedByDocument.OrderBy(g => g.Key))
        {
            var viewsInDoc = docGroup.ToList();

            // Get browser columns for this document's views
            var docBrowserColumnNames = viewsInDoc.First().Keys
                .Where(k => k != "Document" && k != "SheetNumber" && k != "Name" &&
                           k != "ViewType" && k != "ElementIdObject" &&
                           k != "__OriginalObject" && k != "__Document")
                .ToList();

            if (docBrowserColumnNames.Count > 0)
            {
                // Sort by browser columns using the helper method
                var browserColumns = docBrowserColumnNames
                    .Select(name => new BrowserOrganizationHelper.BrowserColumn { Name = name })
                    .ToList();
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

            allViewsData.AddRange(viewsInDoc);
        }

        // Find currentViewIndex after sorting using Equals for document comparison
        int currentViewIndex = allViewsData.FindIndex(row =>
        {
            if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View view &&
                row.ContainsKey("__Document") && row["__Document"] is Document doc)
            {
                return doc.Equals(activeDoc) &&
                       activeUidoc.ActiveView != null &&
                       view.Id == activeUidoc.ActiveView.Id;
            }
            return false;
        });

        // Build property names - Document first, then browser columns, then standard columns
        var propertyNames = new List<string>();
        propertyNames.Add("Document");
        propertyNames.AddRange(allBrowserColumnNames);
        propertyNames.Add("SheetNumber");
        propertyNames.Add("Name");
        propertyNames.Add("ViewType");
        propertyNames.Add("Sheet");

        // Set initial selection
        List<int> initialSelectionIndices = currentViewIndex >= 0
            ? new List<int> { currentViewIndex }
            : new List<int>();

        // Show the grid
        CustomGUIs.SetCurrentUIDocument(activeUidoc);
        var selectedDicts = CustomGUIs.DataGrid(allViewsData, propertyNames, false, initialSelectionIndices);

        if (selectedDicts == null || selectedDicts.Count == 0)
            return Result.Failed;

        // Get the selected view and its document
        var selectedDict = selectedDicts.First();
        View selectedView = selectedDict["__OriginalObject"] as View;
        Document targetDoc = selectedDict["__Document"] as Document;

        if (selectedView == null || targetDoc == null)
            return Result.Failed;

        // Check if the selected view is in the active document
        if (targetDoc != activeDoc)
        {
            // Need to switch documents first
            if (string.IsNullOrEmpty(targetDoc.PathName))
            {
                // Document is unsaved - cannot switch to it programmatically
                TaskDialog.Show("Unsaved Document",
                    $"The selected view is in document '{targetDoc.Title}' which has not been saved.\n\n" +
                    $"Revit can only switch to saved documents programmatically. " +
                    $"Please save '{targetDoc.Title}' or switch to it manually.");
                return Result.Failed;
            }

            // Switch to the target document
            try
            {
                // Suppress logging to avoid recording the intermediate view when document opens
                RevitBallet.LogViewChanges.SuppressLogging();

                // CRITICAL FIX: Call OpenDocumentFile first to prevent close/reopen cycle
                // Per Revit API guidance: If OpenDocumentFile is called before OpenAndActivateDocument,
                // the latter will ONLY activate the document without closing/reopening it.
                uiApp.Application.OpenDocumentFile(targetDoc.PathName);

                UIDocument newUidoc = uiApp.OpenAndActivateDocument(targetDoc.PathName);

                // Switch to the selected view (still suppressed to avoid intermediate view logging)
                newUidoc.ActiveView = selectedView;

                // Resume logging
                RevitBallet.LogViewChanges.ResumeLogging();

                // Manually log the intended view activation to ensure it's recorded
                // (even if selectedView was already the active view and didn't trigger ViewActivated)
                string sessionId = RevitBallet.LogViewChanges.GetSessionId();
                LogViewChangesDatabase.LogViewActivation(
                    sessionId: sessionId,
                    documentSessionId: sessionId,
                    documentTitle: targetDoc.Title,
                    documentPath: targetDoc.PathName ?? "",
                    viewId: selectedView.Id,
                    viewTitle: selectedView.Title,
                    viewType: selectedView.ViewType.ToString(),
                    timestamp: DateTime.Now
                );
            }
            catch (Exception ex)
            {
                // Make sure to resume logging even if an error occurs
                RevitBallet.LogViewChanges.ResumeLogging();

                TaskDialog.Show("Error", $"Failed to switch to document '{targetDoc.Title}':\n\n{ex.Message}");
                return Result.Failed;
            }
        }
        else
        {
            // Same document - just switch the view
            activeUidoc.ActiveView = selectedView;
        }

        return Result.Succeeded;
    }
}
