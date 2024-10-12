using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SetRevisionsToSheetsWhereRevisionIsByCloud : IExternalCommand
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

        // Find sheets with revisions from clouds that are not explicitly added
        var sheetsWithCloudOnlyRevisions = new List<ViewSheet>();
        var allSheets = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .ToList();

        var sheetToCloudRevisionsMap = new Dictionary<ViewSheet, List<ElementId>>();

        foreach (var sheet in allSheets)
        {
            var allRevisionIds = sheet.GetAllRevisionIds();
            var additionalRevisionIds = sheet.GetAdditionalRevisionIds();

            var cloudRevisionIds = allRevisionIds
                .Where(id => !additionalRevisionIds.Contains(id))
                .ToList();

            var relevantCloudRevisionIds = cloudRevisionIds
                .Where(id => selectedRevisionSequences.Contains(((Revision)doc.GetElement(id)).SequenceNumber))
                .ToList();

            if (relevantCloudRevisionIds.Any())
            {
                sheetsWithCloudOnlyRevisions.Add(sheet);
                sheetToCloudRevisionsMap[sheet] = relevantCloudRevisionIds;
            }
        }

        if (!sheetsWithCloudOnlyRevisions.Any())
        {
            TaskDialog.Show("Sheets", "No sheets found with revisions added only by clouds.");
            return Result.Cancelled;
        }

        // Prepare the list for DataGrid
        List<Dictionary<string, object>> sheetEntries = new List<Dictionary<string, object>>();
        foreach (var sheet in sheetsWithCloudOnlyRevisions)
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

        // Assign the cloud-assigned revisions to the selected sheets
        using (Transaction trans = new Transaction(doc, "Assign Cloud Revisions to Sheets"))
        {
            trans.Start();

            foreach (var sheetEntry in selectedSheets)
            {
                string sheetNumber = sheetEntry["Sheet Number"].ToString();
                ViewSheet sheet = sheetsWithCloudOnlyRevisions.FirstOrDefault(s => s.SheetNumber == sheetNumber);

                if (sheet != null)
                {
                    // Get existing additional revisions
                    ICollection<ElementId> revisionIds = sheet.GetAdditionalRevisionIds().ToList();

                    // Add only the revisions that were previously assigned by cloud
                    foreach (var revId in sheetToCloudRevisionsMap[sheet])
                    {
                        if (!revisionIds.Contains(revId))
                        {
                            revisionIds.Add(revId);
                        }
                    }

                    // Update the sheet with the new set of additional revisions
                    sheet.SetAdditionalRevisionIds(revisionIds);
                }
            }

            trans.Commit();
        }

        return Result.Succeeded;
    }
}
