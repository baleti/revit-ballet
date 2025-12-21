using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SwitchViewByHistoryInDocument : IExternalCommand
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

        // Build grid data from history - only for the active document
        var allViewsData = new List<Dictionary<string, object>>();
        int currentViewIndex = -1;
        int viewIndex = 0;

        foreach (var entry in history)
        {
            // Only include views from the active document
            if (entry.DocumentTitle != activeDoc.Title)
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

            Element viewElement = activeDoc.GetElement(viewId);

            // If element not found by ID, try finding by title
            if (viewElement == null || !(viewElement is View))
            {
                viewElement = new FilteredElementCollector(activeDoc)
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

            // Get browser organization columns for this document
            List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                BrowserOrganizationHelper.GetBrowserColumnsForViews(activeDoc);

            // Add browser organization columns
            BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, view, activeDoc, browserColumns);

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

            // Check if this is the currently active view
            bool isCurrentView = (activeUidoc.ActiveView != null) &&
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
            TaskDialog.Show("Info", "No views from current document found in history.");
            return Result.Failed;
        }

        // Build property names - browser columns first, then standard columns (no Document column)
        var propertyNames = new List<string>();

        // Get unique browser column names
        var allBrowserColumnNames = allViewsData
            .SelectMany(d => d.Keys)
            .Where(k => k != "SheetNumber" && k != "Name" &&
                       k != "ViewType" && k != "ElementIdObject" &&
                       k != "__OriginalObject")
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

        // Get the selected view
        var selectedDict = selectedDicts.First();
        View selectedView = selectedDict["__OriginalObject"] as View;

        if (selectedView == null)
            return Result.Failed;

        // Switch to the selected view (always in the active document)
        activeUidoc.ActiveView = selectedView;

        return Result.Succeeded;
    }
}
