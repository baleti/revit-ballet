// C# 7.3 — Revit 2024 API
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class UnsetRevisionToSelectedSheets : IExternalCommand
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

        if (pickIds == null || pickIds.Count == 0)
        {
            TaskDialog.Show("Unset Revision",
                "Select one or more sheets before running the command.");
            return Result.Cancelled;
        }

        List<ViewSheet> targetSheets = pickIds
            .Select(id => doc.GetElement(id))
            .OfType<ViewSheet>()
            .ToList();

        if (targetSheets.Count == 0)
        {
            TaskDialog.Show("Unset Revision",
                "No sheets were found in the current selection.");
            return Result.Cancelled;
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
                spanAllScreens: false,
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
