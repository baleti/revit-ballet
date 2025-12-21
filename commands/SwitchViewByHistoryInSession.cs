using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SwitchViewByHistoryInSession : IExternalCommand
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

        // Get session-wide view history from database
        string sessionId = RevitBallet.RevitBallet.SessionId;
        var history = LogViewChangesDatabase.GetViewHistoryForSession(sessionId, limit: 1000);

        if (history.Count == 0)
        {
            TaskDialog.Show("Info", "No views found in session history.");
            return Result.Failed;
        }

        // Build lookup of document titles to open documents
        var documentsByTitle = new Dictionary<string, Document>();
        foreach (Document doc in uiApp.Application.Documents)
        {
            if (!doc.IsLinked && !doc.IsFamilyDocument)
            {
                documentsByTitle[doc.Title] = doc;
            }
        }

        // Build grid data from history
        var allViewsData = new List<Dictionary<string, object>>();
        int currentViewIndex = -1;
        int viewIndex = 0;

        foreach (var entry in history)
        {
            // Check if this document is still open
            if (!documentsByTitle.TryGetValue(entry.DocumentTitle, out Document doc))
                continue;

            // Get the view element
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

            Element viewElement = doc.GetElement(viewId);

            // If element not found by ID, try finding by title
            if (viewElement == null || !(viewElement is View))
            {
                viewElement = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .FirstOrDefault(v => v.Title == entry.ViewTitle);
            }

            if (viewElement == null || !(viewElement is View view))
                continue;

            // Skip templates and system views
            if (view.IsTemplate ||
                view.ViewType == ViewType.ProjectBrowser ||
                view.ViewType == ViewType.SystemBrowser)
                continue;

            // Build view data dictionary
            var dict = new Dictionary<string, object>();
            dict["Document"] = entry.DocumentTitle;

            // Get browser organization columns for this document
            List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                BrowserOrganizationHelper.GetBrowserColumnsForViews(doc);

            // Add browser organization columns
            BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, view, doc, browserColumns);

            // Add standard columns
            if (view is ViewSheet sheet)
            {
                dict["SheetNumber"] = sheet.SheetNumber;
                dict["Name"] = sheet.Name;
            }
            else
            {
                dict["SheetNumber"] = "";
                dict["Name"] = view.Name;
            }

            dict["ViewType"] = view.ViewType;
            dict["ElementIdObject"] = view.Id;
            dict["__OriginalObject"] = view;
            dict["__Document"] = doc;

            // Check if this is the currently active view using Equals for document comparison
            bool isCurrentView = doc.Equals(activeDoc) &&
                                (activeUidoc.ActiveView != null) &&
                                (view.Id == activeUidoc.ActiveView.Id);

            if (isCurrentView)
            {
                currentViewIndex = viewIndex;
            }

            allViewsData.Add(dict);
            viewIndex++;
        }

        if (allViewsData.Count == 0)
        {
            TaskDialog.Show("Info", "No views from open documents found in history.");
            return Result.Failed;
        }

        // Build property names - Document first, then browser columns, then standard columns
        var propertyNames = new List<string>();
        propertyNames.Add("Document");

        // Get unique browser column names across all documents
        var allBrowserColumnNames = allViewsData
            .SelectMany(d => d.Keys)
            .Where(k => k != "Document" && k != "SheetNumber" && k != "Name" &&
                       k != "ViewType" && k != "ElementIdObject" &&
                       k != "__OriginalObject" && k != "__Document")
            .Distinct()
            .ToList();

        propertyNames.AddRange(allBrowserColumnNames);
        propertyNames.Add("SheetNumber");
        propertyNames.Add("Name");
        propertyNames.Add("ViewType");

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

        // Check if the selected view is in the active document using Equals
        if (!targetDoc.Equals(activeDoc))
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

                UIDocument newUidoc = uiApp.OpenAndActivateDocument(targetDoc.PathName);

                // Switch to the selected view (still suppressed to avoid intermediate view logging)
                newUidoc.ActiveView = selectedView;

                // Resume logging
                RevitBallet.LogViewChanges.ResumeLogging();

                // Manually log the intended view activation to ensure it's recorded
                // (even if selectedView was already the active view and didn't trigger ViewActivated)
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
