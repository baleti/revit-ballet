using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class ListSheetsByRevisions : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get all revisions in the project
        IList<Revision> revisions = new FilteredElementCollector(doc)
            .OfClass(typeof(Revision))
            .Cast<Revision>()
            .ToList();

        if (revisions.Count == 0)
        {
            TaskDialog.Show("Revisions", "No revisions found in the project.");
            return Result.Cancelled;
        }

        // Create data entries for the DataGrid
        List<Dictionary<string, object>> revisionEntries = revisions
            .Select(revision => new Dictionary<string, object>
            {
                { "Revision Sequence", revision.SequenceNumber },
                { "Revision Date", revision.RevisionDate },
                { "Revision Description", revision.Description },
                { "Issued To", revision.IssuedTo },
                { "Issued By", revision.IssuedBy }
            })
            .ToList();

        // Display the revisions and let the user select one or more revisions
        List<string> revisionProperties = new List<string> { "Revision Sequence", "Revision Date", "Revision Description", "Issued To", "Issued By"};
        List<Dictionary<string, object>> selectedRevisions = CustomGUIs.DataGrid(revisionEntries, revisionProperties, false, new List<int> { revisionEntries.Count - 1 });

        if (selectedRevisions.Count == 0)
        {
            return Result.Cancelled;
        }

        // Get the selected revision sequence numbers
        var selectedRevisionSequences = selectedRevisions
            .Select(revision => Convert.ToInt32(revision["Revision Sequence"]))
            .ToHashSet();

        // Find sheets with any of the selected revisions
        IList<ViewSheet> sheetsWithRevisions = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet))
            .Cast<ViewSheet>()
            .Where(sheet => sheet.GetAllRevisionIds()
                                 .Any(revId => doc.GetElement(revId) is Revision rev && selectedRevisionSequences.Contains(rev.SequenceNumber)))
            .ToList();

        if (sheetsWithRevisions.Count == 0)
        {
            TaskDialog.Show("Sheets", "No sheets found with the selected revisions.");
            return Result.Cancelled;
        }

        // Create data entries for the sheets
        List<Dictionary<string, object>> sheetEntries = sheetsWithRevisions
            .Select(sheet => new Dictionary<string, object>
            {
                { "Sheet Title", sheet.Title },
                { "Sheet Number", sheet.SheetNumber }
            })
            .ToList();

        // Display the sheets and let the user select one or more sheets
        List<string> sheetProperties = new List<string> { "Sheet Title", "Sheet Number" };
        List<Dictionary<string, object>> selectedSheets = CustomGUIs.DataGrid(sheetEntries, sheetProperties, false);

        if (selectedSheets.Count == 0)
        {
            return Result.Cancelled;
        }

        // Open the selected sheets
        foreach (var selectedSheetEntry in selectedSheets)
        {
            string selectedSheetNumber = selectedSheetEntry["Sheet Number"].ToString();

            // Find the sheet by its number
            ViewSheet selectedSheet = sheetsWithRevisions
                .FirstOrDefault(sheet => sheet.SheetNumber == selectedSheetNumber);

            if (selectedSheet != null)
            {
                uiDoc.ActiveView = selectedSheet;
            }
        }

        return Result.Succeeded;
    }
}
