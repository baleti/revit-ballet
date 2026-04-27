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
[CommandMeta("View, Link")]
public class CloseWorksets : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        => WorksetToggleHelper.Run(commandData, ref message, closing: true);
}

// OpenWorksets: open/show worksets in host doc (no selection), in views (views selected),
//              or reload linked models with more open worksets (links selected).
[Transaction(TransactionMode.Manual)]
[CommandMeta("View, Link")]
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

        // Links take priority over views in selection.
        // Accept both RevitLinkInstance and RevitLinkType (e.g. from SelectFamilyTypesInDocument).
        var selectedLinks = new List<RevitLinkInstance>();
        var seenTypeIds   = new HashSet<ElementId>();

        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            if (elem is RevitLinkInstance li && li.GetLinkDocument() != null)
            {
                if (seenTypeIds.Add(li.GetTypeId()))
                    selectedLinks.Add(li);
            }
            else if (elem is RevitLinkType && seenTypeIds.Add(id))
            {
                // Find any loaded instance of this type
                RevitLinkInstance inst = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .FirstOrDefault(x => x.GetTypeId() == id && x.GetLinkDocument() != null);
                if (inst != null)
                    selectedLinks.Add(inst);
            }
        }

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
    // Shows one DataGrid for all selected links. Workset names are the key;
    // the Links column lists which link types have that workset (type name only,
    // not instance name which includes location). Hidden for single-link.

    private static Result HandleLinkedModels(Document doc, List<RevitLinkInstance> selectedLinks, bool closing, ref string message)
    {
        // Resolve type names and collect worksets from every link.
        // worksetsByName: name → list of (workset, owning link instance)
        var worksetsByName  = new Dictionary<string, List<(Workset ws, RevitLinkInstance link)>>();
        var linkTypeNameFor = new Dictionary<ElementId, string>();

        foreach (var linkInst in selectedLinks)
        {
            var linkedDoc = linkInst.GetLinkDocument();
            if (linkedDoc == null) continue;

            if (!linkedDoc.IsWorkshared)
            {
                TaskDialog.Show("Info", $"{doc.GetElement(linkInst.GetTypeId())?.Name ?? linkInst.Name}: not workshared, skipping.");
                continue;
            }

            ElementId typeId = linkInst.GetTypeId();
            if (!linkTypeNameFor.ContainsKey(typeId))
                linkTypeNameFor[typeId] = (doc.GetElement(typeId) as RevitLinkType)?.Name ?? linkInst.Name;

            foreach (Workset ws in new FilteredWorksetCollector(linkedDoc)
                .OfKind(WorksetKind.UserWorkset).ToWorksets())
            {
                if (!worksetsByName.TryGetValue(ws.Name, out var list))
                    worksetsByName[ws.Name] = list = new List<(Workset, RevitLinkInstance)>();
                list.Add((ws, linkInst));
            }
        }

        if (worksetsByName.Count == 0)
        {
            TaskDialog.Show("Info", "No user worksets found in the selected linked models.");
            return Result.Cancelled;
        }

        bool multiLink = selectedLinks.Count > 1;

        var rows = worksetsByName.OrderBy(kv => kv.Key).Select(kv =>
        {
            var entries = kv.Value;
            string status = entries.All(e => e.ws.IsOpen)  ? "Open"  :
                            entries.All(e => !e.ws.IsOpen) ? "Closed" : "Mixed";
            var row = new Dictionary<string, object>
            {
                { "Workset", kv.Key },
                { "Status",  status }
            };
            if (multiLink)
                row["Links"] = string.Join(", ", entries
                    .Select(e => linkTypeNameFor[e.link.GetTypeId()])
                    .Distinct().OrderBy(n => n));
            return row;
        }).ToList();

        var columns = multiLink
            ? new List<string> { "Workset", "Status", "Links" }
            : new List<string> { "Workset", "Status" };

        var picked = CustomGUIs.DataGrid(rows, columns, false);
        if (picked == null || picked.Count == 0) return Result.Cancelled;

        var selectedNames = picked
            .Select(r => r.TryGetValue("Workset", out object n) ? n as string : null)
            .Where(n => n != null)
            .ToHashSet();

        // Apply one WorksetConfiguration reload per link.
        foreach (var linkInst in selectedLinks)
        {
            var linkedDoc = linkInst.GetLinkDocument();
            if (linkedDoc == null || !linkedDoc.IsWorkshared) continue;

            var linkType = doc.GetElement(linkInst.GetTypeId()) as RevitLinkType;
            if (linkType == null) continue;

            var allWorksets = new FilteredWorksetCollector(linkedDoc)
                .OfKind(WorksetKind.UserWorkset).ToWorksets().ToList();

            // Use CloseAllWorksets + Open: OpenAllWorksets ignores Close() calls.
            var toOpen = allWorksets
                .Where(ws => closing ? (ws.IsOpen && !selectedNames.Contains(ws.Name))
                                     : (ws.IsOpen || selectedNames.Contains(ws.Name)))
                .Select(ws => ws.Id).ToList();

            var config = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            if (toOpen.Count > 0) config.Open(toOpen);

            linkType.LoadFrom(linkType.GetExternalFileReference().GetAbsolutePath(), config);
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
