using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class OpenLastSessionViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        string projectName = doc != null ? doc.Title : "UnknownProject";

        // Get current session ID to exclude it
        string currentSessionId = RevitBallet.RevitBallet.SessionId;

        // Get view history from previous sessions for this document
        var previousHistory = LogViewChangesDatabase.GetPreviousSessionViews(projectName, currentSessionId, limit: 1000);

        if (previousHistory.Count == 0)
        {
            TaskDialog.Show("Info", "No views found from previous sessions.");
            return Result.Failed;
        }

        // Get all views in the document
        List<View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .ToList();

        // Build list of views from previous session history with metadata
        List<View> savedViews = new List<View>();
        Dictionary<ElementId, ViewHistoryEntry> viewMetadata = new Dictionary<ElementId, ViewHistoryEntry>();
        HashSet<ElementId> seenIds = new HashSet<ElementId>();

        // Diagnostic file for debugging
        string diagnosticPath = System.IO.Path.Combine(
            RevitBallet.Commands.PathHelper.RuntimeDirectory,
            "diagnostics",
            $"OpenLastSessionViews-{System.DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");
        var diagnosticLines = new List<string>();
        diagnosticLines.Add($"=== OpenLastSessionViews Diagnostics at {System.DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ===");
        diagnosticLines.Add($"Project: {projectName}");
        diagnosticLines.Add($"Current Session: {currentSessionId}");
        diagnosticLines.Add($"Previous History Count: {previousHistory.Count}");
        diagnosticLines.Add("");

        foreach (var entry in previousHistory)
        {
            ElementId viewId;

            try
            {
                diagnosticLines.Add($"Processing entry: ViewId={entry.ViewId}, ViewTitle={entry.ViewTitle}, SessionId={entry.SessionId.Substring(0, 8)}");

                viewId = entry.ViewId.ToElementId();
                diagnosticLines.Add($"  Successfully converted to ElementId: {viewId.AsLong()}");

                // Skip duplicates
                if (seenIds.Contains(viewId))
                {
                    diagnosticLines.Add($"  Skipping duplicate ViewId: {viewId.AsLong()}");
                    continue;
                }
            }
            catch (Exception ex)
            {
                diagnosticLines.Add($"  ERROR converting ViewId {entry.ViewId}: {ex.Message}");
                diagnosticLines.Add($"  Exception Type: {ex.GetType().Name}");
                diagnosticLines.Add($"  Stack Trace: {ex.StackTrace}");

                // Write diagnostics immediately when error occurs
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
                System.IO.File.WriteAllLines(diagnosticPath, diagnosticLines);

                // Skip this entry instead of crashing
                continue;
            }

            View matchingView = allViews.FirstOrDefault(v => v.Id == viewId);

            // If not found by ID, try by title
            if (matchingView == null)
            {
                matchingView = allViews.FirstOrDefault(v =>
                    v.Title.Equals(entry.ViewTitle, StringComparison.OrdinalIgnoreCase));
            }

            if (matchingView != null)
            {
                diagnosticLines.Add($"  Found matching view: {matchingView.Name} (Id: {matchingView.Id.AsLong()})");
                savedViews.Add(matchingView);
                viewMetadata[matchingView.Id] = entry;
                seenIds.Add(matchingView.Id);
            }
            else
            {
                diagnosticLines.Add($"  No matching view found for ViewId={entry.ViewId}, Title={entry.ViewTitle}");
            }
        }

        // Write diagnostics file
        diagnosticLines.Add("");
        diagnosticLines.Add($"Total saved views found: {savedViews.Count}");
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
        System.IO.File.WriteAllLines(diagnosticPath, diagnosticLines);

        if (savedViews.Count == 0)
        {
            TaskDialog.Show("Info", "No matching views found in current document.");
            return Result.Failed;
        }

        // Get browser organization columns
        List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
            BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, savedViews) ?? new List<BrowserOrganizationHelper.BrowserColumn>();

        // Create session ID to friendly name mapping
        var uniqueSessionIds = viewMetadata.Values.Select(e => e.SessionId).Distinct().OrderByDescending(sid => {
            var maxTimestamp = viewMetadata.Values.Where(e => e.SessionId == sid).Max(e => e.Timestamp);
            return maxTimestamp;
        }).ToList();

        var sessionIdToName = new Dictionary<string, string>();
        for (int i = 0; i < uniqueSessionIds.Count; i++)
        {
            sessionIdToName[uniqueSessionIds[i]] = $"Session {i + 1}";
        }

        // Create grid data with browser organization columns
        List<Dictionary<string, object>> gridData = new List<Dictionary<string, object>>();
        foreach (View view in savedViews)
        {
            var dict = new Dictionary<string, object>();

            // Add browser organization columns first
            BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, view, doc, browserColumns);

            // Add session and timestamp metadata
            if (viewMetadata.TryGetValue(view.Id, out var metadata))
            {
                dict["Session"] = sessionIdToName.TryGetValue(metadata.SessionId, out var sessionName)
                    ? sessionName
                    : metadata.SessionId.Substring(0, 8);
                dict["Last Accessed"] = metadata.Timestamp.ToString("yyyy-MM-dd HH:mm:ss");
            }
            else
            {
                dict["Session"] = "";
                dict["Last Accessed"] = "";
            }

            // Then add standard columns
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

            gridData.Add(dict);
        }

        // Sort by browser organization columns before showing in DataGrid
        if (browserColumns != null && browserColumns.Count > 0)
        {
            gridData = BrowserOrganizationHelper.SortByBrowserColumns(gridData, browserColumns);
        }

        // Build property names - browser columns first, then session metadata, then standard columns
        var propertyNames = new List<string>();
        if (browserColumns != null && browserColumns.Count > 0)
        {
            propertyNames.AddRange(browserColumns.Select(bc => bc.Name));
        }
        propertyNames.Add("Session");
        propertyNames.Add("Last Accessed");
        propertyNames.Add("SheetNumber");
        propertyNames.Add("Name");
        propertyNames.Add("ViewType");

        // Display the saved views using CustomGUIs.DataGrid
        CustomGUIs.SetCurrentUIDocument(uidoc);
        var selectedDicts = CustomGUIs.DataGrid(gridData, propertyNames, false);
        var selectedViews = selectedDicts;

        // Open the selected views in Revit
        foreach (var viewEntry in selectedViews)
        {
            if (viewEntry.ContainsKey("__OriginalObject") && viewEntry["__OriginalObject"] is View view)
            {
                uidoc.RequestViewChange(view);
            }
        }

        return Result.Succeeded;
    }
}
