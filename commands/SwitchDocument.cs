using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SwitchDocument : IExternalCommand
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

        // Get SessionId once for all database operations
        string sessionId = RevitBallet.LogViewChanges.GetSessionId();

        // Collect all open documents with their last viewed view
        var documentsData = new List<Dictionary<string, object>>();
        int currentDocIndex = -1;
        int docIndex = 0;

        foreach (Document doc in uiApp.Application.Documents)
        {
            // Skip linked documents (read-only references)
            if (doc.IsLinked)
                continue;

            // Skip family documents
            if (doc.IsFamilyDocument)
                continue;

            string projectName = doc.Title;

            // Get the last viewed view from the database
            string lastViewName = "";
            ElementId lastViewId = ElementId.InvalidElementId;

            var history = LogViewChangesDatabase.GetViewHistoryForDocument(sessionId, doc.Title, limit: 1);
            if (history.Count > 0)
            {
                lastViewName = history[0].ViewTitle;
                try
                {
                    lastViewId = history[0].ViewId.ToElementId();
                }
                catch (Exception)
                {
                    // Skip invalid ViewId (e.g., 0, -1, or corrupted data)
                    lastViewId = ElementId.InvalidElementId;
                }
            }

            // Check if this is the active document
            bool isActiveDoc = (doc == activeDoc);
            if (isActiveDoc)
            {
                currentDocIndex = docIndex;
            }

            var dict = new Dictionary<string, object>
            {
                ["Document"] = projectName,
                ["LastView"] = lastViewName,
                ["__Document"] = doc,
                ["__LastViewId"] = lastViewId,
                ["__LastViewName"] = lastViewName
            };

            documentsData.Add(dict);
            docIndex++;
        }

        if (documentsData.Count == 0)
        {
            TaskDialog.Show("Info", "No documents open.");
            return Result.Failed;
        }

        // Sort by document name
        documentsData = documentsData.OrderBy(d => d["Document"].ToString()).ToList();

        // Update currentDocIndex after sorting using Equals (not reference equality)
        currentDocIndex = documentsData.FindIndex(d =>
        {
            Document doc = d["__Document"] as Document;
            return doc != null && doc.Equals(activeDoc);
        });

        // Build property names
        var propertyNames = new List<string> { "Document", "LastView" };

        // Set initial selection
        List<int> initialSelectionIndices = currentDocIndex >= 0
            ? new List<int> { currentDocIndex }
            : new List<int>();

        // Show the grid
        CustomGUIs.SetCurrentUIDocument(activeUidoc);
        var selectedDicts = CustomGUIs.DataGrid(documentsData, propertyNames, false, initialSelectionIndices);

        if (selectedDicts == null || selectedDicts.Count == 0)
            return Result.Failed;

        // Get the selected document
        var selectedDict = selectedDicts.First();
        Document targetDoc = selectedDict["__Document"] as Document;
        ElementId targetViewId = selectedDict["__LastViewId"] as ElementId;
        string targetViewName = selectedDict["__LastViewName"]?.ToString();

        if (targetDoc == null)
            return Result.Failed;

        // If same document, do nothing
        if (targetDoc == activeDoc)
            return Result.Succeeded;

        // Switch to the target document
        if (string.IsNullOrEmpty(targetDoc.PathName))
        {
            // Document is unsaved - cannot switch to it programmatically
            TaskDialog.Show("Unsaved Document",
                $"The selected document '{targetDoc.Title}' has not been saved.\n\n" +
                $"Revit can only switch to saved documents programmatically. " +
                $"Please save '{targetDoc.Title}' or switch to it manually.");
            return Result.Failed;
        }

        try
        {
            // Suppress logging to avoid recording the intermediate view when document opens
            RevitBallet.LogViewChanges.SuppressLogging();

            UIDocument newUidoc = uiApp.OpenAndActivateDocument(targetDoc.PathName);

            // Try to switch to the last viewed view (still suppressed to avoid intermediate view logging)
            View finalView = null;
            if (targetViewId != null && targetViewId != ElementId.InvalidElementId)
            {
                View targetView = targetDoc.GetElement(targetViewId) as View;

                // If view not found by ID, try by name
                if (targetView == null && !string.IsNullOrEmpty(targetViewName))
                {
                    targetView = new FilteredElementCollector(targetDoc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .FirstOrDefault(v => v.Title == targetViewName);
                }

                if (targetView != null)
                {
                    newUidoc.ActiveView = targetView;
                    finalView = targetView;
                }
            }

            // If no specific view was set, use whatever view is currently active
            if (finalView == null)
            {
                finalView = newUidoc.ActiveView;
            }

            // Resume logging
            RevitBallet.LogViewChanges.ResumeLogging();

            // Manually log the final view activation to ensure it's recorded
            // (even if it was already the active view and didn't trigger ViewActivated)
            if (finalView != null)
            {
                LogViewChangesDatabase.LogViewActivation(
                    sessionId: sessionId,
                    documentSessionId: sessionId,
                    documentTitle: targetDoc.Title,
                    documentPath: targetDoc.PathName ?? "",
                    viewId: finalView.Id,
                    viewTitle: finalView.Title,
                    viewType: finalView.ViewType.ToString(),
                    timestamp: DateTime.Now
                );
            }
        }
        catch (Exception ex)
        {
            // Make sure to resume logging even if an error occurs
            RevitBallet.LogViewChanges.ResumeLogging();

            TaskDialog.Show("Error", $"Failed to switch to document '{targetDoc.Title}':\n\n{ex.Message}");
            return Result.Failed;
        }

        return Result.Succeeded;
    }
}
