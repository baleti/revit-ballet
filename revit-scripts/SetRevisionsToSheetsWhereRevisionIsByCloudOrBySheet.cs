using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SetRevisionsToSheetsWhereRevisionIsByCloudOrBySheet : IExternalCommand
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

        // Get the selected revision sequence numbers
        var selectedRevisionSequences = selectedRevisions
            .Select(r => Convert.ToInt32(r["Revision Sequence"]))
            .ToHashSet();

        // Find sheets with any of the selected revisions
        var sheetsWithSelectedRevisions = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet))
            .Cast<ViewSheet>()
            .Where(sheet => sheet.GetAllRevisionIds()
                                 .Select(revId => doc.GetElement(revId) as Revision)
                                 .Any(rev => rev != null && selectedRevisionSequences.Contains(rev.SequenceNumber)))
            .ToList();

        if (!sheetsWithSelectedRevisions.Any())
        {
            TaskDialog.Show("Sheets", "No sheets found with the selected revisions.");
            return Result.Cancelled;
        }

        // Prepare the list for DataGrid
        List<Dictionary<string, object>> sheetEntries = new List<Dictionary<string, object>>();
        foreach (var sheet in sheetsWithSelectedRevisions)
        {
            var entry = new Dictionary<string, object>
            {
                { "Sheet Number", sheet.SheetNumber },
                { "Sheet Name", sheet.Name }
            };
            sheetEntries.Add(entry);
        }

        // Show sheet selection dialog
        List<string> sheetProperties = new List<string> { "Sheet Number", "Sheet Name" };
        var selectedSheets = CustomGUIs.DataGrid(sheetEntries, sheetProperties, spanAllScreens: false);

        if (!selectedSheets.Any())
        {
            TaskDialog.Show("Selection", "No sheets were selected.");
            return Result.Cancelled;
        }

        // Assign the selected revisions to the selected sheets
        using (Transaction trans = new Transaction(doc, "Assign Revisions to Sheets"))
        {
            trans.Start();

            foreach (var sheetEntry in selectedSheets)
            {
                string sheetNumber = sheetEntry["Sheet Number"].ToString();
                ViewSheet sheet = sheetsWithSelectedRevisions.FirstOrDefault(s => s.SheetNumber == sheetNumber);

                if (sheet != null)
                {
                    ICollection<ElementId> revisionIds = sheet.GetAdditionalRevisionIds().ToList();

                    foreach (var rev in revisions.Where(r => selectedRevisionSequences.Contains(r.SequenceNumber)))
                    {
                        if (!revisionIds.Contains(rev.Id))
                        {
                            revisionIds.Add(rev.Id);
                        }
                    }

                    sheet.SetAdditionalRevisionIds(revisionIds);
                }
            }

            trans.Commit();
        }

        TaskDialog.Show("Success", "Selected revisions have been assigned to selected sheets.");
        return Result.Succeeded;
    }
}
