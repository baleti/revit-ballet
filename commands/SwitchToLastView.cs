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

        // DIAGNOSTIC: Log initial state
        var diagnosticLines = new System.Collections.Generic.List<string>();
        diagnosticLines.Add($"=== SwitchToLastView Diagnostic at {System.DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ===");
        diagnosticLines.Add($"Active Document: {activeDoc.Title}");
        diagnosticLines.Add($"Active View: {activeUidoc.ActiveView?.Name} (ID: {activeUidoc.ActiveView?.Id?.AsLong()})");
        diagnosticLines.Add($"Session ID: {sessionId}");
        diagnosticLines.Add($"History Count: {history.Count}");
        diagnosticLines.Add("");
        diagnosticLines.Add("=== View History (most recent first) ===");
        for (int i = 0; i < history.Count && i < 10; i++)
        {
            var entry = history[i];
            diagnosticLines.Add($"  [{i}] {entry.DocumentTitle} | {entry.ViewTitle} (ID: {entry.ViewId}) | {entry.Timestamp:yyyy-MM-dd HH:mm:ss.fff}");
        }
        diagnosticLines.Add("");

        if (history.Count < 2)
        {
            message = "Not enough views in history.";
            diagnosticLines.Add($"FAILED: Not enough views in history (count: {history.Count})");
            WriteDiagnostic(diagnosticLines);
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

        diagnosticLines.Add($"Current View Index in History: {currentViewIndex}");
        diagnosticLines.Add("");

        // Iterate through history FORWARD (towards older views)
        // History is DESC (newest first), so previous view is at HIGHER index (older in time)
        // Start from the entry AFTER the current view
        int startIndex = currentViewIndex >= 0 ? currentViewIndex + 1 : 1;
        diagnosticLines.Add($"Start Index: {startIndex}");

        // IMPORTANT: After a cross-document switch, we want to go back to the PREVIOUS DOCUMENT,
        // not to an intermediate view (like Landing Page) in the CURRENT document.
        // So we prefer the first view from a DIFFERENT document.
        bool preferDifferentDocument = startIndex < history.Count &&
                                       history[startIndex].DocumentTitle == activeDoc.Title;

        if (preferDifferentDocument)
        {
            diagnosticLines.Add($"Next entry is in same document - will prefer different document");
        }

        diagnosticLines.Add("");
        diagnosticLines.Add("=== Searching for valid previous view ===");

        for (int i = startIndex; i < history.Count; i++)
        {
            var entry = history[i];
            diagnosticLines.Add($"[{i}] Trying: {entry.DocumentTitle} | {entry.ViewTitle} (ID: {entry.ViewId})");

            // If we're preferring a different document, skip entries in the current document
            if (preferDifferentDocument && entry.DocumentTitle == activeDoc.Title)
            {
                diagnosticLines.Add($"    SKIP: Same document (preferring different document)");
                continue;
            }

            ElementId viewId;
            try
            {
                viewId = entry.ViewId.ToElementId();
            }
            catch (Exception ex)
            {
                // Skip invalid ViewId entries (e.g., 0, -1, or corrupted data)
                diagnosticLines.Add($"    SKIP: Invalid ViewId - {ex.Message}");
                continue;
            }

            // Try to find the document containing this view
            Document targetDoc = null;

            // First, check if it's in the active document
            if (entry.DocumentTitle == activeDoc.Title)
            {
                targetDoc = activeDoc;
                diagnosticLines.Add($"    Target doc: Active document");
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
                diagnosticLines.Add($"    Target doc: {(targetDoc != null ? "Found in open documents" : "NOT FOUND")}");
            }

            if (targetDoc == null)
            {
                diagnosticLines.Add($"    SKIP: Document not open");
                continue; // Document not open, skip this entry
            }

            // Try to get the view from the target document
            View view = targetDoc.GetElement(viewId) as View;

            if (view == null)
            {
                diagnosticLines.Add($"    SKIP: View not found in document");
                continue; // View not found, skip this entry
            }

            diagnosticLines.Add($"    Found valid view: {view.Name}");

            // Found a valid view - now switch to it
            if (targetDoc != activeDoc)
            {
                diagnosticLines.Add($"    Switching documents: {activeDoc.Title} → {targetDoc.Title}");

                // Need to switch documents first
                if (string.IsNullOrEmpty(targetDoc.PathName))
                {
                    // Unsaved document - skip to next entry
                    diagnosticLines.Add($"    SKIP: Target document is unsaved");
                    continue;
                }

                diagnosticLines.Add($"    PathName: {targetDoc.PathName}");

                try
                {
                    // Suppress logging to avoid recording the intermediate view when document opens
                    RevitBallet.LogViewChanges.SuppressLogging();

                    diagnosticLines.Add($"    Calling OpenAndActivateDocument...");

                    // Switch to the target document
                    UIDocument newUidoc = uiApp.OpenAndActivateDocument(targetDoc.PathName);

                    diagnosticLines.Add($"    Setting active view to: {view.Name}");

                    // Switch to the view (still suppressed to avoid intermediate view logging)
                    newUidoc.ActiveView = view;

                    // Resume logging
                    RevitBallet.LogViewChanges.ResumeLogging();

                    diagnosticLines.Add($"    Manually logging view activation");

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

                    diagnosticLines.Add($"SUCCESS: Switched to {targetDoc.Title} | {view.Name}");
                    WriteDiagnostic(diagnosticLines);

                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    // Make sure to resume logging even if an error occurs
                    RevitBallet.LogViewChanges.ResumeLogging();

                    diagnosticLines.Add($"    ERROR: {ex.Message}");
                    diagnosticLines.Add($"    Trying next entry...");

                    // Failed to switch - try next entry
                    continue;
                }
            }
            else
            {
                diagnosticLines.Add($"    Same document - just switching view");

                // Same document - just switch the view
                activeUidoc.ActiveView = view;

                diagnosticLines.Add($"SUCCESS: Switched to view {view.Name}");
                WriteDiagnostic(diagnosticLines);

                return Result.Succeeded;
            }
        }

        message = "No valid previous view found in history.";
        diagnosticLines.Add("");
        diagnosticLines.Add($"FAILED: No valid previous view found in history");
        WriteDiagnostic(diagnosticLines);

        return Result.Failed;
    }

    private static void WriteDiagnostic(System.Collections.Generic.List<string> lines)
    {
        try
        {
            string diagnosticPath = System.IO.Path.Combine(
                PathHelper.RuntimeDirectory,
                "diagnostics",
                $"SwitchToLastView-{System.DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");

            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
            System.IO.File.WriteAllLines(diagnosticPath, lines);
        }
        catch
        {
            // Silently fail - don't interrupt the command
        }
    }
}
