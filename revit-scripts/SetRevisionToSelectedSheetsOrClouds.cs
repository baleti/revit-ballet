// C# 7.3 — Revit 2024 API
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SetRevisionToSelectedSheetsOrClouds : IExternalCommand
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
        // 2. Use the elements currently selected in Revit
        // ─────────────────────────────────────────────
        ICollection<ElementId> pickIds = uiDoc.GetSelectionIds();

        if (pickIds == null || pickIds.Count == 0)
        {
            TaskDialog.Show("Set Revision", "Select one or more sheets or revision clouds before running the command.");
            return Result.Cancelled;
        }

        // Separate sheets and revision clouds from selection
        List<ViewSheet> targetSheets = new List<ViewSheet>();
        List<RevisionCloud> targetClouds = new List<RevisionCloud>();

        foreach (ElementId id in pickIds)
        {
            Element element = doc.GetElement(id);
            if (element is ViewSheet sheet)
            {
                targetSheets.Add(sheet);
            }
            else if (element is RevisionCloud cloud)
            {
                targetClouds.Add(cloud);
            }
        }

        if (targetSheets.Count == 0 && targetClouds.Count == 0)
        {
            TaskDialog.Show("Set Revision", "No sheets or revision clouds were found in the current selection.");
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
        // 4. Add chosen revisions to the selected elements
        // ─────────────────────────────────────────────
        using (Transaction tx = new Transaction(doc, "Assign Revisions to Selected Sheets and Clouds"))
        {
            tx.Start();

            // Process sheets
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

            // Process revision clouds
            foreach (RevisionCloud cloud in targetClouds)
            {
                foreach (Revision rev in revisionsToAssign)
                {
                    // For revision clouds, we set the RevisionId parameter
                    // Note: A revision cloud can only be associated with one revision at a time
                    // So we'll use the first revision from the selection
                    cloud.RevisionId = revisionsToAssign.First().Id;
                    break; // Only assign the first revision to the cloud
                }
            }

            tx.Commit();
        }

        // ─────────────────────────────────────────────
        // 5. Show summary of what was processed
        // ─────────────────────────────────────────────
        string summary = $"Successfully assigned revisions to:\n";
        if (targetSheets.Count > 0)
            summary += $"• {targetSheets.Count} sheet(s)\n";
        if (targetClouds.Count > 0)
            summary += $"• {targetClouds.Count} revision cloud(s)\n";
        
        return Result.Succeeded;
    }
}
