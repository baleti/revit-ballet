using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SwitchToLastView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument activeUidoc = uiApp.ActiveUIDocument;
        Document activeDoc = activeUidoc.Document;

        if (activeDoc == null)
        {
            message = "No active document.";
            return Result.Failed;
        }

        // Get session-wide view history (all documents, most recent first)
        string sessionId = RevitBallet.RevitBallet.SessionId;
        var history = LogViewChangesDatabase.GetViewHistoryForSession(sessionId, limit: 100);

        if (history.Count < 2)
        {
            message = "Not enough views in history.";
            return Result.Failed;
        }

        // Find the current view in the history
        ElementId currentViewId = activeUidoc.ActiveView?.Id;
        int currentViewIndex = -1;

        if (currentViewId != null)
        {
            currentViewIndex = history.FindIndex(entry =>
                entry.ViewId == currentViewId.AsLong() &&
                entry.DocumentTitle == activeDoc.Title);
        }

        // Iterate through history FORWARD (towards older views)
        // History is DESC (newest first), so previous view is at HIGHER index (older in time)
        // Start from the entry AFTER the current view
        int startIndex = currentViewIndex >= 0 ? currentViewIndex + 1 : 1;

        for (int i = startIndex; i < history.Count; i++)
        {
            var entry = history[i];

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

            // Try to find the document containing this view
            Document targetDoc = null;

            // First, check if it's in the active document
            if (entry.DocumentTitle == activeDoc.Title)
            {
                targetDoc = activeDoc;
            }
            else
            {
                // Search through all open documents by title
                foreach (Document doc in uiApp.Application.Documents)
                {
                    if (!doc.IsLinked && !doc.IsFamilyDocument && doc.Title == entry.DocumentTitle)
                    {
                        targetDoc = doc;
                        break;
                    }
                }
            }

            if (targetDoc == null)
                continue; // Document not open, skip this entry

            // Try to get the view from the target document
            View view = targetDoc.GetElement(viewId) as View;

            if (view == null)
                continue; // View not found, skip this entry

            // Found a valid view - now switch to it
            if (targetDoc != activeDoc)
            {
                // Need to switch documents first
                if (string.IsNullOrEmpty(targetDoc.PathName))
                {
                    // Unsaved document - skip to next entry
                    continue;
                }

                try
                {
                    // Suppress logging to avoid recording the intermediate view when document opens
                    RevitBallet.LogViewChanges.SuppressLogging();

                    // Switch to the target document
                    UIDocument newUidoc = uiApp.OpenAndActivateDocument(targetDoc.PathName);

                    // Switch to the view (still suppressed to avoid intermediate view logging)
                    newUidoc.ActiveView = view;

                    // Resume logging
                    RevitBallet.LogViewChanges.ResumeLogging();

                    // Manually log the intended view activation to ensure it's recorded
                    // (even if view was already the active view and didn't trigger ViewActivated)
                    LogViewChangesDatabase.LogViewActivation(
                        sessionId: sessionId,
                        documentSessionId: sessionId,
                        documentTitle: targetDoc.Title,
                        documentPath: targetDoc.PathName ?? "",
                        viewId: view.Id,
                        viewTitle: view.Title,
                        viewType: view.ViewType.ToString(),
                        timestamp: DateTime.Now
                    );

                    return Result.Succeeded;
                }
                catch
                {
                    // Make sure to resume logging even if an error occurs
                    RevitBallet.LogViewChanges.ResumeLogging();

                    // Failed to switch - try next entry
                    continue;
                }
            }
            else
            {
                // Same document - just switch the view
                activeUidoc.ActiveView = view;
                return Result.Succeeded;
            }
        }

        message = "No valid previous view found in history.";
        return Result.Failed;
    }
}
