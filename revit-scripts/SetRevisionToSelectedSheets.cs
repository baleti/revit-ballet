// C# 7.3 — Revit 2024 API
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SetRevisionToSelectedSheets : IExternalCommand
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
        UIDocument uiDoc   = uiApp.ActiveUIDocument;
        Document   doc     = uiDoc.Document;

        // ─────────────────────────────────────────────
        // 1. Gather all revisions and display DataGrid
        // ─────────────────────────────────────────────
        List<Revision> revisions = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Revisions)
            .WhereElementIsNotElementType()
            .Cast<Revision>()
            .ToList();

        List<Dictionary<string, object>> revisionEntries = revisions
            .Select(r => new Dictionary<string, object>
            {
                { "Revision Sequence", r.SequenceNumber },
                { "Revision Date",     r.RevisionDate   },
                { "Description",       r.Description    },
                { "Issued By",         r.IssuedBy       },
                { "Issued To",         r.IssuedTo       }
            })
            .ToList();

        List<string> revisionProps = new List<string>
        {
            "Revision Sequence",
            "Revision Date",
            "Description",
            "Issued By",
            "Issued To"
        };

        // Show grid; pre-select most-recent row
        List<Dictionary<string, object>> selectedRevisions =
            CustomGUIs.DataGrid(
                revisionEntries,
                revisionProps,
                spanAllScreens: false,
                new List<int> { revisionEntries.Count - 1 });

        if (selectedRevisions == null || selectedRevisions.Count == 0)
        {
            TaskDialog.Show("Set Revision", "No revision was selected.");
            return Result.Cancelled;
        }

        // ─────────────────────────────────────────────
        // 2. Use the sheets currently selected in Revit
        // ─────────────────────────────────────────────
        ICollection<ElementId> pickIds = uiDoc.Selection.GetElementIds();

        if (pickIds == null || pickIds.Count == 0)
        {
            TaskDialog.Show("Set Revision", "Select one or more sheets before running the command.");
            return Result.Cancelled;
        }

        List<ViewSheet> targetSheets = pickIds
            .Select(id => doc.GetElement(id))
            .OfType<ViewSheet>()
            .ToList();

        if (targetSheets.Count == 0)
        {
            TaskDialog.Show("Set Revision", "No sheets were found in the current selection.");
            return Result.Cancelled;
        }

        // ─────────────────────────────────────────────
        // 3. Resolve selected revision objects
        // ─────────────────────────────────────────────
        List<Revision> revisionsToAssign = new List<Revision>();

        foreach (Dictionary<string, object> selRev in selectedRevisions)
        {
            int seq = Convert.ToInt32(selRev["Revision Sequence"]);
            Revision rev = revisions.FirstOrDefault(r => r.SequenceNumber == seq);
            if (rev != null)
                revisionsToAssign.Add(rev);
        }

        if (revisionsToAssign.Count == 0)
        {
            TaskDialog.Show("Set Revision", "No valid revisions were selected.");
            return Result.Failed;
        }

        // ─────────────────────────────────────────────
        // 4. Add chosen revisions to the selected sheets
        // ─────────────────────────────────────────────
        using (Transaction tx = new Transaction(doc, "Assign Revisions to Selected Sheets"))
        {
            tx.Start();

            foreach (ViewSheet sheet in targetSheets)
            {
                IList<ElementId> revIds = sheet.GetAdditionalRevisionIds().ToList();

                foreach (Revision rev in revisionsToAssign)
                {
                    if (!revIds.Contains(rev.Id))
                        revIds.Add(rev.Id);
                }

                sheet.SetAdditionalRevisionIds(revIds);
            }

            tx.Commit();
        }

        return Result.Succeeded;
    }
}
