using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ListSheetsByRevisionsWhereRevisionIsByCloud : IExternalCommand
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

        foreach (var sheet in allSheets)
        {
            var allRevisionIds = sheet.GetAllRevisionIds();
            var additionalRevisionIds = sheet.GetAdditionalRevisionIds();

            var cloudRevisionIds = allRevisionIds
                .Where(id => !additionalRevisionIds.Contains(id))
                .ToList();

            if (cloudRevisionIds.Any(id => selectedRevisionSequences.Contains(((Revision)doc.GetElement(id)).SequenceNumber)))
            {
                sheetsWithCloudOnlyRevisions.Add(sheet);
            }
        }

        if (!sheetsWithCloudOnlyRevisions.Any())
        {
            return Result.Cancelled;
        }

        // Prepare the list for DataGrid
        List<Dictionary<string, object>> sheetEntries = new List<Dictionary<string, object>>();
        foreach (var sheet in sheetsWithCloudOnlyRevisions)
        {
            // Collect revision sequence numbers assigned by cloud and present in selected revisions
            var revisionSequenceNumbers = sheet.GetAllRevisionIds()
                .Where(revId => !sheet.GetAdditionalRevisionIds().Contains(revId))  // Filter out revisions that are not assigned by cloud
                .Select(revId => doc.GetElement(revId) as Revision)
                .Where(rev => rev != null && selectedRevisionSequences.Contains(rev.SequenceNumber))
                .Select(rev => rev.SequenceNumber.ToString())
                .ToList();

            var entry = new Dictionary<string, object>
            {
                { "Sheet Number", sheet.SheetNumber },
                { "Sheet Name", sheet.Name },
                { "Assigned Cloud Revisions", string.Join(", ", revisionSequenceNumbers) } // New column for cloud-assigned revision sequence numbers
            };
            sheetEntries.Add(entry);
        }

        // Show sheet selection dialog with the new column
        List<string> sheetProperties = new List<string> { "Sheet Number", "Sheet Name", "Assigned Cloud Revisions" };
        var selectedSheets = CustomGUIs.DataGrid(sheetEntries, sheetProperties, spanAllScreens: false);

        if (!selectedSheets.Any())
        {
            return Result.Cancelled;
        }

        // Open the selected sheet
        string selectedSheetNumber = selectedSheets[0]["Sheet Number"].ToString();
        ViewSheet selectedSheet = sheetsWithCloudOnlyRevisions.FirstOrDefault(s => s.SheetNumber == selectedSheetNumber);

        if (selectedSheet != null)
        {
            uiapp.ActiveUIDocument.ActiveView = selectedSheet;
        }
        else
        {
            TaskDialog.Show("Error", "Selected sheet could not be found.");
            return Result.Failed;
        }

        return Result.Succeeded;
    }
}
