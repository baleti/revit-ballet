using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CopySelectedRevisionCloudsToSheetsAndInheritRevision : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        // Get the current document
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View currentView = doc.ActiveView;
        
        // Get the selected elements in the active view
        ICollection<ElementId> selectedElementIds = uidoc.GetSelectionIds();
        if (!selectedElementIds.Any())
        {
            TaskDialog.Show("Error", "No elements selected. Please select revision clouds to copy.");
            return Result.Failed;
        }
        
        // Filter to get only revision clouds from selection
        List<RevisionCloud> selectedRevisionClouds = new List<RevisionCloud>();
        foreach (ElementId id in selectedElementIds)
        {
            Element elem = doc.GetElement(id);
            if (elem is RevisionCloud revCloud)
            {
                selectedRevisionClouds.Add(revCloud);
            }
        }
        
        if (!selectedRevisionClouds.Any())
        {
            TaskDialog.Show("Error", "No revision clouds selected. Please select revision clouds to copy.");
            return Result.Failed;
        }
        
        // Get source sheets from selected revision clouds
        HashSet<ElementId> sourceSheetIds = new HashSet<ElementId>();
        foreach (var revCloud in selectedRevisionClouds)
        {
            View ownerView = doc.GetElement(revCloud.OwnerViewId) as View;
            if (ownerView is ViewSheet)
            {
                sourceSheetIds.Add(ownerView.Id);
            }
            else if (ownerView != null)
            {
                // Check if this view is placed on any sheet
                var viewSheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Where(vs => vs.GetAllPlacedViews().Contains(ownerView.Id))
                    .ToList();
                foreach (var vs in viewSheets)
                {
                    sourceSheetIds.Add(vs.Id);
                }
            }
        }
        
        // Get all sheets in the document, excluding source sheets
        var sheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Where(s => !sourceSheetIds.Contains(s.Id))
                    .OrderBy(s => s.Title)
                    .ToList();
        
        // Convert sheets to dictionary format for DataGrid
        var sheetDictionaries = sheets.Select(s => new Dictionary<string, object>
        {
            { "Title", s.Title },
            { "Element", s }
        }).ToList();
        
        // Use the CustomGUIs.DataGrid to prompt the user to select target sheets
        var selectedSheetDicts = CustomGUIs.DataGrid(
            sheetDictionaries,
            new List<string> { "Title" },
            spanAllScreens: false
        );
        
        if (!selectedSheetDicts.Any())
        {
            message = "No target sheets selected.";
            return Result.Failed;
        }
        
        // Extract ViewSheet objects from selected dictionaries
        List<ViewSheet> selectedSheets = selectedSheetDicts
            .Select(dict => dict["Element"] as ViewSheet)
            .Where(s => s != null)
            .ToList();
        
        if (!selectedSheets.Any())
        {
            message = "No target sheets selected.";
            return Result.Failed;
        }
        
        // Start a transaction to copy and paste the elements
        using (Transaction transaction = new Transaction(doc, "Copy Revision Clouds to Sheets with Latest Revision"))
        {
            transaction.Start();
            
            CopyPasteOptions options = new CopyPasteOptions();
            
            foreach (ViewSheet targetSheet in selectedSheets)
            {
                try
                {
                    // Get the latest revision for this sheet
                    ElementId latestRevisionId = GetLatestRevisionForSheet(doc, targetSheet);
                    
                    if (latestRevisionId == null || latestRevisionId == ElementId.InvalidElementId)
                    {
                        TaskDialog.Show("Warning", $"No revisions found on sheet {targetSheet.SheetNumber}. Skipping this sheet.");
                        continue;
                    }
                    
                    // Copy the revision clouds to the target sheet
                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                        currentView, 
                        selectedRevisionClouds.Select(rc => rc.Id).ToList(), 
                        targetSheet, 
                        Transform.Identity, 
                        options);
                    
                    // Update the revision of each copied cloud
                    foreach (ElementId copiedId in copiedIds)
                    {
                        Element copiedElem = doc.GetElement(copiedId);
                        if (copiedElem is RevisionCloud copiedCloud)
                        {
                            copiedCloud.RevisionId = latestRevisionId;
                        }
                    }
                }
                catch (Exception ex)
                {
                    message = $"Error copying revision clouds to sheet {targetSheet.SheetNumber}: {ex.Message}";
                    transaction.RollBack();
                    return Result.Failed;
                }
            }
            
            transaction.Commit();
        }
        
        return Result.Succeeded;
    }
    
    private ElementId GetLatestRevisionForSheet(Document doc, ViewSheet sheet)
    {
        // First, check revisions assigned directly to the sheet
        ICollection<ElementId> sheetRevisionIds = sheet.GetAdditionalRevisionIds();
        
        // Also get all revision clouds on the sheet and its views to find their revisions
        HashSet<ElementId> allRevisionIds = new HashSet<ElementId>(sheetRevisionIds);
        
        // Add the sheet itself to check for revision clouds
        var viewsToCheck = new List<ElementId> { sheet.Id };
        
        // Add all views placed on this sheet (without triggering regeneration)
        var placedViews = sheet.GetAllPlacedViews();
        viewsToCheck.AddRange(placedViews);
        
        // Collect revision clouds in all these views
        var revisionClouds = new FilteredElementCollector(doc)
            .WhereElementIsNotElementType()
            .OfClass(typeof(RevisionCloud))
            .Cast<RevisionCloud>()
            .Where(rc => viewsToCheck.Contains(rc.OwnerViewId))
            .ToList();
        
        // Add revision IDs from revision clouds
        foreach (var cloud in revisionClouds)
        {
            if (cloud.RevisionId != ElementId.InvalidElementId)
            {
                allRevisionIds.Add(cloud.RevisionId);
            }
        }
        
        if (!allRevisionIds.Any())
        {
            return ElementId.InvalidElementId;
        }
        
        // Get all revision elements and find the one with the highest sequence number
        ElementId latestRevisionId = ElementId.InvalidElementId;
        int highestSequence = -1;
        
        foreach (ElementId revId in allRevisionIds)
        {
            Revision revision = doc.GetElement(revId) as Revision;
            if (revision != null)
            {
                int sequence = revision.SequenceNumber;
                if (sequence > highestSequence)
                {
                    highestSequence = sequence;
                    latestRevisionId = revId;
                }
            }
        }
        
        return latestRevisionId;
    }
}
