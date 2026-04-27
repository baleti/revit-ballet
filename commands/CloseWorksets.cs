using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;

// CloseWorksets: close/hide worksets in host doc (no selection), in views (views selected),
//               or reload linked models with fewer open worksets (links selected).
[Transaction(TransactionMode.Manual)]
[CommandMeta("")]
public class CloseWorksets : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        => WorksetToggleHelper.Run(commandData, ref message, closing: true);
}

// OpenWorksets: open/show worksets in host doc (no selection), in views (views selected),
//              or reload linked models with more open worksets (links selected).
[Transaction(TransactionMode.Manual)]
[CommandMeta("")]
public class OpenWorksets : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        => WorksetToggleHelper.Run(commandData, ref message, closing: false);
}

internal static class WorksetToggleHelper
{
    public static Result Run(ExternalCommandData commandData, ref string message, bool closing)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        if (uidoc == null) { message = "No active document."; return Result.Failed; }
        Document doc = uidoc.Document;

        if (!doc.IsWorkshared)
        {
            TaskDialog.Show("Info", "This document is not workshared. Worksets are not available.");
            return Result.Cancelled;
        }

        ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();

        // Links take priority over views in selection
        var selectedLinks = selectedIds
            .Select(id => doc.GetElement(id))
            .OfType<RevitLinkInstance>()
            .Where(li => li.GetLinkDocument() != null)
            .ToList();

        if (selectedLinks.Count > 0)
            return HandleLinkedModels(doc, selectedLinks, closing, ref message);

        var targetViews = ExtractTargetViews(doc, selectedIds);
        if (targetViews.Count > 0)
            return HandleViewWorksets(doc, targetViews, closing, ref message);

        return HandleDocumentWorksets(uidoc, doc, closing, ref message);
    }

    // ── Document-level global visibility ─────────────────────────────────────
    // WorksetDefaultVisibilitySettings controls whether each workset is visible
    // by default across all views. This is the API-accessible equivalent of
    // opening/closing worksets globally (the Revit API cannot truly unload
    // worksets from an already-open document; WorksetConfiguration only applies
    // when initially loading a document).

    private static Result HandleDocumentWorksets(UIDocument uidoc, Document doc, bool closing, ref string message)
    {
        var worksets = new FilteredWorksetCollector(doc)
            .OfKind(WorksetKind.UserWorkset)
            .ToWorksets()
            .OrderBy(w => w.Name)
            .ToList();

        if (worksets.Count == 0)
        {
            TaskDialog.Show("Info", "No user worksets in this document.");
            return Result.Cancelled;
        }

        var wdvs = new FilteredElementCollector(doc)
            .OfClass(typeof(WorksetDefaultVisibilitySettings))
            .Cast<WorksetDefaultVisibilitySettings>()
            .FirstOrDefault();

        var rows = worksets.Select(ws =>
        {
            bool globallyVisible = wdvs != null ? wdvs.IsWorksetVisible(ws.Id) : ws.IsVisibleByDefault;
            return new Dictionary<string, object>
            {
                { "Workset",    ws.Name },
                { "Global",     globallyVisible ? "Visible" : "Hidden" },
                { "Editable",   ws.IsEditable ? "Yes" : "No" },
                { "WorksetId",  ws.Id }
            };
        }).ToList();

        var picked = CustomGUIs.DataGrid(rows, new List<string> { "Workset", "Global", "Editable" }, false);
        if (picked == null || picked.Count == 0) return Result.Cancelled;

        if (wdvs == null)
        {
            TaskDialog.Show("Error", "Could not find WorksetDefaultVisibilitySettings in this document.");
            return Result.Failed;
        }

        using (var t = new Transaction(doc, closing ? "Close Worksets" : "Open Worksets"))
        {
            t.Start();
            foreach (var row in picked)
            {
                if (!row.TryGetValue("WorksetId", out object wsIdObj) || !(wsIdObj is WorksetId wsId)) continue;
                wdvs.SetWorksetVisibility(wsId, !closing); // true = visible (open), false = hidden (close)
            }
            t.Commit();
        }

        return Result.Succeeded;
    }

    // ── View-level workset visibility ─────────────────────────────────────────

    private static Result HandleViewWorksets(Document doc, List<View> targetViews, bool closing, ref string message)
    {
        var worksets = new FilteredWorksetCollector(doc)
            .OfKind(WorksetKind.UserWorkset)
            .ToWorksets()
            .OrderBy(w => w.Name)
            .ToList();

        if (worksets.Count == 0)
        {
            TaskDialog.Show("Info", "No user worksets in this document.");
            return Result.Cancelled;
        }

        var rows = worksets.Select(ws =>
        {
            var visPerView = targetViews.Select(v => GetVisibilityLabel(v, ws.Id)).Distinct().ToList();
            return new Dictionary<string, object>
            {
                { "Workset",    ws.Name },
                { "Visibility", visPerView.Count == 1 ? visPerView[0] : "Mixed" },
                { "WorksetId",  ws.Id }
            };
        }).ToList();

        var picked = CustomGUIs.DataGrid(rows, new List<string> { "Workset", "Visibility" }, false);
        if (picked == null || picked.Count == 0) return Result.Cancelled;

        var targetVisibility = closing ? WorksetVisibility.Hidden : WorksetVisibility.Visible;
        string txName = targetViews.Count == 1
            ? $"{(closing ? "Hide" : "Show")} Worksets in {targetViews[0].Name}"
            : $"{(closing ? "Hide" : "Show")} Worksets in {targetViews.Count} Views";

        using (var t = new Transaction(doc, txName))
        {
            t.Start();
            foreach (var row in picked)
            {
                if (!row.TryGetValue("WorksetId", out object wsIdObj) || !(wsIdObj is WorksetId wsId)) continue;
                foreach (var view in targetViews)
                    view.SetWorksetVisibility(wsId, targetVisibility);
            }
            t.Commit();
        }

        return Result.Succeeded;
    }

    // ── Linked model open/close via reload ───────────────────────────────────
    // RevitLinkType.LoadFrom(ModelPath, WorksetConfiguration) reloads the link
    // with a new set of open worksets. Multiple instances of the same link type
    // are processed once (they share one loaded document).

    private static Result HandleLinkedModels(Document doc, List<RevitLinkInstance> selectedLinks, bool closing, ref string message)
    {
        var processedTypeIds = new HashSet<ElementId>();

        foreach (var linkInstance in selectedLinks)
        {
            ElementId typeId = linkInstance.GetTypeId();
            if (!processedTypeIds.Add(typeId)) continue;

            var linkType = doc.GetElement(typeId) as RevitLinkType;
            if (linkType == null) continue;

            var linkedDoc = linkInstance.GetLinkDocument();
            if (linkedDoc == null) continue;

            if (!linkedDoc.IsWorkshared)
            {
                TaskDialog.Show("Info", $"{linkType.Name}: not workshared, no worksets to change.");
                continue;
            }

            var worksets = new FilteredWorksetCollector(linkedDoc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .OrderBy(w => w.Name)
                .ToList();

            if (worksets.Count == 0)
            {
                TaskDialog.Show("Info", $"{linkType.Name}: no user worksets.");
                continue;
            }

            var rows = worksets.Select(ws => new Dictionary<string, object>
            {
                { "Workset",   ws.Name },
                { "Status",    ws.IsOpen ? "Open" : "Closed" },
                { "WorksetId", ws.Id }
            }).ToList();

            string title = $"{(closing ? "Close" : "Open")} Worksets in {linkType.Name}";
            var picked = CustomGUIs.DataGrid(rows, new List<string> { "Workset", "Status" }, false);
            if (picked == null || picked.Count == 0) continue;

            var selectedWsIds = picked
                .Where(r => r.TryGetValue("WorksetId", out object v) && v is WorksetId)
                .Select(r => (WorksetId)r["WorksetId"])
                .ToHashSet();

            // Build new WorksetConfiguration.
            // Start with all worksets open, then close what should be closed.
            var alreadyClosed = worksets.Where(ws => !ws.IsOpen).Select(ws => ws.Id).ToList();

            List<WorksetId> toClose;
            if (closing)
                toClose = alreadyClosed.Concat(selectedWsIds).Distinct().ToList();
            else
                toClose = alreadyClosed.Except(selectedWsIds).ToList();

            var config = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
            if (toClose.Count > 0)
                config.Close(toClose);

            ModelPath path = linkType.GetExternalFileReference().GetAbsolutePath();

            using (var t = new Transaction(doc, title))
            {
                t.Start();
                linkType.LoadFrom(path, config);
                t.Commit();
            }
        }

        return Result.Succeeded;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    // Views with a template applied are redirected to the template so the change
    // takes effect (follows the pattern established in HideWorksetsInViews).
    private static List<View> ExtractTargetViews(Document doc, ICollection<ElementId> selectedIds)
    {
        var seen = new HashSet<ElementId>();
        var result = new List<View>();

        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            View view = null;

            if (elem is View v && !(v is ViewSheet || v is ViewSchedule))
                view = v;
            else if (elem is Viewport vp)
                view = doc.GetElement(vp.ViewId) as View;

            if (view == null || view is ViewSheet || view is ViewSchedule) continue;

            View target = view;
            if (!view.IsTemplate && view.ViewTemplateId != ElementId.InvalidElementId)
                target = doc.GetElement(view.ViewTemplateId) as View ?? view;

            if (seen.Add(target.Id))
                result.Add(target);
        }

        return result;
    }

    private static string GetVisibilityLabel(View view, WorksetId wsId)
    {
        try
        {
            WorksetVisibility vis = view.GetWorksetVisibility(wsId);
            switch (vis)
            {
                case WorksetVisibility.Visible:          return "Shown";
                case WorksetVisibility.Hidden:           return "Hidden";
                case WorksetVisibility.UseGlobalSetting:
                    Workset ws = view.Document.GetWorksetTable().GetWorkset(wsId);
                    return ws?.IsVisibleByDefault == true ? "Global (Visible)" : "Global (Hidden)";
                default:                                 return vis.ToString();
            }
        }
        catch { return "N/A"; }
    }
}
