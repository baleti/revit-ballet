// C# 7.3 — Revit 2024 API
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[CommandMeta("Sheet")]
public class UnsetRevisionToSheet : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        // ─────────────────────────────────────────────
        // Revit context
        // ─────────────────────────────────────────────
        UIApplication uiApp = commandData.Application;
        UIDocument    uiDoc = uiApp.ActiveUIDocument;
        Document      doc   = uiDoc.Document;

        // ─────────────────────────────────────────────
        // 1. Get sheets currently selected in Revit
        // ─────────────────────────────────────────────
        ICollection<ElementId> pickIds = uiDoc.GetSelectionIds();

        List<ViewSheet> targetSheets = new List<ViewSheet>();
        foreach (ElementId id in pickIds)
        {
            Element e = doc.GetElement(id);
            if (e is ViewSheet vs && !targetSheets.Contains(vs)) targetSheets.Add(vs);
            else if (e is Viewport vp && vp.SheetId != ElementId.InvalidElementId)
            {
                var vs2 = doc.GetElement(vp.SheetId) as ViewSheet;
                if (vs2 != null && !targetSheets.Contains(vs2)) targetSheets.Add(vs2);
            }
        }

        if (targetSheets.Count == 0)
        {
            CustomGUIs.SetCurrentUIDocument(uiDoc);
            var allSheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet)).Cast<ViewSheet>().ToList();
            var gridData = CustomGUIs.ConvertToDataGridFormat(allSheets, new List<string> { "Sheet Number", "Name" });
            var chosen = CustomGUIs.DataGrid(gridData, new List<string> { "Sheet Number", "Name" }, false);
            if (chosen == null) return Result.Cancelled;
            targetSheets = CustomGUIs.ExtractOriginalObjects<ViewSheet>(chosen) ?? new List<ViewSheet>();
            if (targetSheets.Count == 0) return Result.Succeeded;
        }

        // ─────────────────────────────────────────────
        // 2. Collect union of revisions already on them
        // ─────────────────────────────────────────────
        HashSet<ElementId> currentRevIds = new HashSet<ElementId>();

        foreach (ViewSheet sheet in targetSheets)
            foreach (ElementId rid in sheet.GetAdditionalRevisionIds())
                currentRevIds.Add(rid);

        if (currentRevIds.Count == 0)
        {
            TaskDialog.Show("Unset Revision",
                "None of the selected sheets have additional revisions.");
            return Result.Cancelled;
        }

        List<Revision> currentRevisions = currentRevIds
            .Select(id => doc.GetElement(id) as Revision)
            .Where(r => r != null)
            .ToList();

        // ─────────────────────────────────────────────
        // 3. Build the grid from ONLY those revisions
        // ─────────────────────────────────────────────
        var revEntries = currentRevisions
            .Select(r => new Dictionary<string, object>
            {
                { "Revision Sequence", r.SequenceNumber },
                { "Revision Date",     r.RevisionDate   },
                { "Description",       r.Description    },
                { "Issued By",         r.IssuedBy       },
                { "Issued To",         r.IssuedTo       }
            })
            .ToList();

        var revProps = new List<string>
        {
            "Revision Sequence",
            "Revision Date",
            "Description",
            "Issued By",
            "Issued To"
        };

        List<Dictionary<string, object>> selectedRevisions =
            CustomGUIs.DataGrid(
                revEntries,
                revProps,
                false,
                /* no pre-selection */ null);

        if (selectedRevisions == null || selectedRevisions.Count == 0)
        {
            TaskDialog.Show("Unset Revision", "No revision was selected.");
            return Result.Cancelled;
        }

        // ─────────────────────────────────────────────
        // 4. Resolve selected revision objects
        // ─────────────────────────────────────────────
        List<Revision> revisionsToRemove = new List<Revision>();

        foreach (Dictionary<string, object> selRev in selectedRevisions)
        {
            int seq = Convert.ToInt32(selRev["Revision Sequence"]);
            Revision rev = currentRevisions
                .FirstOrDefault(r => r.SequenceNumber == seq);
            if (rev != null)
                revisionsToRemove.Add(rev);
        }

        if (revisionsToRemove.Count == 0)
        {
            TaskDialog.Show("Unset Revision", "No valid revisions were selected.");
            return Result.Failed;
        }

        // ─────────────────────────────────────────────
        // 5. Remove chosen revisions from selected sheets
        // ─────────────────────────────────────────────
        using (Transaction tx = new Transaction(doc,
            "Remove Revisions from Selected Sheets"))
        {
            tx.Start();

            foreach (ViewSheet sheet in targetSheets)
            {
                IList<ElementId> revIds = sheet.GetAdditionalRevisionIds().ToList();

                foreach (Revision rev in revisionsToRemove)
                    revIds.Remove(rev.Id);

                sheet.SetAdditionalRevisionIds(revIds);
            }

            tx.Commit();
        }

        return Result.Succeeded;
    }
}
