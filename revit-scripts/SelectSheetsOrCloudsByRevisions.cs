using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class SelectSheetsOrCloudsByRevisions : IExternalCommand
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

        // Find revision clouds with any of the selected revisions
        IList<RevisionCloud> cloudsWithRevisions = new FilteredElementCollector(doc)
            .OfClass(typeof(RevisionCloud))
            .Cast<RevisionCloud>()
            .Where(cloud => cloud.RevisionId != null && 
                           doc.GetElement(cloud.RevisionId) is Revision rev && 
                           selectedRevisionSequences.Contains(rev.SequenceNumber))
            .ToList();

        if (sheetsWithRevisions.Count == 0 && cloudsWithRevisions.Count == 0)
        {
            TaskDialog.Show("Elements", "No sheets or revision clouds found with the selected revisions.");
            return Result.Cancelled;
        }

        // Create combined data entries for sheets and revision clouds
        List<Dictionary<string, object>> elementEntries = new List<Dictionary<string, object>>();

        // Add sheet entries
        foreach (var sheet in sheetsWithRevisions)
        {
            // Get all revisions for this sheet
            var sheetRevisionIds = sheet.GetAllRevisionIds();
            var sheetRevisions = sheetRevisionIds
                .Select(revId => doc.GetElement(revId) as Revision)
                .Where(rev => rev != null)
                .OrderBy(rev => rev.SequenceNumber)
                .ToList();

            // Create comma-delimited strings for all revision data
            string allRevNums = string.Join(", ", sheetRevisions.Select(r => r.SequenceNumber.ToString()));
            string allRevDates = string.Join(", ", sheetRevisions.Select(r => r.RevisionDate ?? "N/A"));
            string allRevDescriptions = string.Join(", ", sheetRevisions.Select(r => r.Description ?? "N/A"));
            string allIssuedTo = string.Join(", ", sheetRevisions.Select(r => r.IssuedTo ?? "N/A"));
            string allIssuedBy = string.Join(", ", sheetRevisions.Select(r => r.IssuedBy ?? "N/A"));
            
            elementEntries.Add(new Dictionary<string, object>
            {
                { "Element Type", "Sheet" },
                { "Sheet Name", sheet.Title },
                { "Sheet Number", sheet.SheetNumber },
                { "All Rev Nums", allRevNums },
                { "All Rev Dates", allRevDates },
                { "All Rev Descriptions", allRevDescriptions },
                { "All Issued To", allIssuedTo },
                { "All Issued By", allIssuedBy },
                { "Element ID", sheet.Id.IntegerValue }
            });
        }

        // Add revision cloud entries
        foreach (var cloud in cloudsWithRevisions)
        {
            var cloudRevision = doc.GetElement(cloud.RevisionId) as Revision;
            string sheetName = "N/A";
            string sheetNumber = "N/A";
            
            // Try to get the sheet info where the cloud is located
            if (cloud.OwnerViewId != null && cloud.OwnerViewId != ElementId.InvalidElementId)
            {
                var view = doc.GetElement(cloud.OwnerViewId) as View;
                if (view != null)
                {
                    // Check if the view is a sheet
                    if (view is ViewSheet viewSheet)
                    {
                        sheetName = viewSheet.Title;
                        sheetNumber = viewSheet.SheetNumber;
                    }
                    else
                    {
                        // For other views, try to find if they're placed on any sheet
                        var allSheets = new FilteredElementCollector(doc)
                            .OfClass(typeof(ViewSheet))
                            .Cast<ViewSheet>();
                        
                        foreach (var checkSheet in allSheets)
                        {
                            var viewports = new FilteredElementCollector(doc)
                                .OfClass(typeof(Viewport))
                                .Cast<Viewport>()
                                .Where(vp => vp.SheetId == checkSheet.Id && vp.ViewId == view.Id);
                            
                            if (viewports.Any())
                            {
                                sheetName = checkSheet.Title;
                                sheetNumber = checkSheet.SheetNumber;
                                break;
                            }
                        }
                    }
                }
            }
            
            // For revision clouds, we only have one revision, but format it consistently
            string revNum = cloudRevision?.SequenceNumber.ToString() ?? "N/A";
            string revDate = cloudRevision?.RevisionDate ?? "N/A";
            string revDescription = cloudRevision?.Description ?? "N/A";
            string issuedTo = cloudRevision?.IssuedTo ?? "N/A";
            string issuedBy = cloudRevision?.IssuedBy ?? "N/A";
            
            elementEntries.Add(new Dictionary<string, object>
            {
                { "Element Type", "Revision Cloud" },
                { "Sheet Name", sheetName },
                { "Sheet Number", sheetNumber },
                { "All Rev Nums", revNum },
                { "All Rev Dates", revDate },
                { "All Rev Descriptions", revDescription },
                { "All Issued To", issuedTo },
                { "All Issued By", issuedBy },
                { "Element ID", cloud.Id.IntegerValue }
            });
        }

        // Sort entries by element type and then by sheet name
        elementEntries = elementEntries
            .OrderBy(entry => entry["Element Type"].ToString())
            .ThenBy(entry => entry["Sheet Name"].ToString())
            .ToList();

        // Display the elements and let the user select one or more
        List<string> elementProperties = new List<string> { "Sheet Name", "Element Type", "Sheet Number", "All Rev Nums", "All Rev Dates", "All Rev Descriptions", "All Issued To", "All Issued By", "Element ID" };
        List<Dictionary<string, object>> selectedElements = CustomGUIs.DataGrid(elementEntries, elementProperties, false);

        if (selectedElements.Count == 0)
        {
            return Result.Cancelled;
        }

        // Process the selected elements
        List<ElementId> elementsToSelect = new List<ElementId>();
        
        foreach (var selectedElementEntry in selectedElements)
        {
            if (!selectedElementEntry.ContainsKey("Element Type") || !selectedElementEntry.ContainsKey("Element ID"))
                continue;

            string elementType = selectedElementEntry["Element Type"].ToString();
            int elementId = Convert.ToInt32(selectedElementEntry["Element ID"]);
            ElementId revitElementId = new ElementId(elementId);

            if (elementType == "Sheet")
            {
                // For sheets, just add them to selection (don't open them)
                ViewSheet selectedSheet = doc.GetElement(revitElementId) as ViewSheet;
                if (selectedSheet != null)
                {
                    elementsToSelect.Add(selectedSheet.Id);
                }
            }
            else if (elementType == "Revision Cloud")
            {
                // For revision clouds, add them to the selection list
                RevisionCloud selectedCloud = doc.GetElement(revitElementId) as RevisionCloud;
                if (selectedCloud != null)
                {
                    elementsToSelect.Add(selectedCloud.Id);
                }
            }
        }
        
        // Add all selected elements to the current selection (append to existing selection)
        if (elementsToSelect.Count > 0)
        {
            // Get current selection
            var currentSelection = uiDoc.GetSelectionIds().ToList();
            
            // Add new elements to current selection
            currentSelection.AddRange(elementsToSelect);
            
            // Set the combined selection
            uiDoc.SetSelectionIds(currentSelection);
        }

        return Result.Succeeded;
    }
}
