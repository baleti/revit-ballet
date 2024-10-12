using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SetRevisionsToSheets : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        // Access Revit application and document
        UIApplication uiapp = commandData.Application;
        Document doc = uiapp.ActiveUIDocument.Document;

        // Get all revisions in the document
        var revisions = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Revisions)
                        .WhereElementIsNotElementType()
                        .Cast<Revision>()
                        .ToList();

        // Prepare the list for DataGrid
        List<Dictionary<string, object>> revisionEntries = new List<Dictionary<string, object>>();
        foreach (var rev in revisions)
        {
            var entry = new Dictionary<string, object>
            {
                { "Revision Sequence", rev.SequenceNumber },
                { "Revision Date", rev.RevisionDate },
                { "Description", rev.Description },
                { "Issued By", rev.IssuedBy },
                { "Issued To", rev.IssuedTo }
            };
            revisionEntries.Add(entry);
        }

        // Show revision selection dialog
        List<string> revisionProperties = new List<string> 
        { 
            "Revision Sequence", 
            "Revision Date", 
            "Description", 
            "Issued By", 
            "Issued To" 
        };
        var selectedRevisions = CustomGUIs.DataGrid(revisionEntries, revisionProperties, spanAllScreens: false, new List<int> { revisionEntries.Count - 1 });

        if (!selectedRevisions.Any())
        {
            TaskDialog.Show("Selection", "No revisions were selected.");
            return Result.Cancelled;
        }

        // Get all sheets in the document
        var sheets = new FilteredElementCollector(doc)
                     .OfCategory(BuiltInCategory.OST_Sheets)
                     .WhereElementIsNotElementType()
                     .Cast<ViewSheet>()
                     .ToList();

        // Prepare the list for DataGrid
        List<Dictionary<string, object>> sheetEntries = new List<Dictionary<string, object>>();
        foreach (var sheet in sheets)
        {
            // Get the current revision sequence for the sheet
            ICollection<ElementId> revisionIds = sheet.GetAdditionalRevisionIds();
            string currentRevision = revisionIds.Any() 
                ? revisions.FirstOrDefault(r => r.Id == revisionIds.Last())?.SequenceNumber.ToString() 
                : "None";
            
            string currentRevisionIssuedTo = revisionIds.Any() 
                ? revisions.FirstOrDefault(r => r.Id == revisionIds.Last())?.IssuedTo 
                : "None";

            var entry = new Dictionary<string, object>
            {
                { "Sheet Number", sheet.SheetNumber },
                { "Sheet Name", sheet.Name },
                { "Current Revision", currentRevision },
                { "Current Revision Issued To", currentRevisionIssuedTo }
            };
            sheetEntries.Add(entry);
        }

        // Show sheet selection dialog
        List<string> sheetProperties = new List<string> { "Sheet Number", "Sheet Name", "Current Revision", "Current Revision Issued To" };
        var selectedSheets = CustomGUIs.DataGrid(sheetEntries, sheetProperties, spanAllScreens: false);

        if (!selectedSheets.Any())
        {
            TaskDialog.Show("Selection", "No sheets were selected.");
            return Result.Cancelled;
        }

        // Prepare to assign selected revisions to the selected sheets
        var revisionsToAssign = new List<Revision>();
        foreach (var selectedRevision in selectedRevisions)
        {
            int selectedSequence = Convert.ToInt32(selectedRevision["Revision Sequence"]);
            var revision = revisions.FirstOrDefault(r => r.SequenceNumber == selectedSequence);

            if (revision != null)
            {
                revisionsToAssign.Add(revision);
            }
        }

        if (!revisionsToAssign.Any())
        {
            TaskDialog.Show("Error", "No valid revisions were selected.");
            return Result.Failed;
        }

        // Assign the selected revisions to the selected sheets
        using (Transaction trans = new Transaction(doc, "Assign Revisions to Sheets"))
        {
            trans.Start();

            foreach (var sheetEntry in selectedSheets)
            {
                string sheetNumber = sheetEntry["Sheet Number"].ToString();
                ViewSheet sheet = sheets.FirstOrDefault(s => s.SheetNumber == sheetNumber);

                if (sheet != null)
                {
                    ICollection<ElementId> revisionIds = sheet.GetAdditionalRevisionIds().ToList();

                    foreach (var revision in revisionsToAssign)
                    {
                        if (!revisionIds.Contains(revision.Id))
                        {
                            revisionIds.Add(revision.Id);
                        }
                    }

                    sheet.SetAdditionalRevisionIds(revisionIds);
                }
            }

            trans.Commit();
        }

        return Result.Succeeded;
    }
}
