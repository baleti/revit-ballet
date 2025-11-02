using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SetRevisionsToCurrentSheet : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        // Access Revit application and document
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get the currently opened sheet
        ViewSheet currentSheet = doc.ActiveView as ViewSheet;
        if (currentSheet == null)
        {
            TaskDialog.Show("Error", "Please open a sheet before running this command.");
            return Result.Failed;
        }

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

        // Prepare to assign selected revisions to the current sheet
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

        // Assign the selected revisions to the current sheet
        using (Transaction trans = new Transaction(doc, "Assign Revisions to Current Sheet"))
        {
            trans.Start();

            ICollection<ElementId> revisionIds = currentSheet.GetAdditionalRevisionIds().ToList();

            foreach (var revision in revisionsToAssign)
            {
                if (!revisionIds.Contains(revision.Id))
                {
                    revisionIds.Add(revision.Id);
                }
            }

            currentSheet.SetAdditionalRevisionIds(revisionIds);

            trans.Commit();
        }

        return Result.Succeeded;
    }
}
