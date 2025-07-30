using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectRevisionCloudsOfSelectedSheets : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;
            
            // Get current selection using SelectionModeManager
            var selectedIds = uidoc.GetSelectionIds();
            
            if (!selectedIds.Any())
            {
                message = "Please select at least one sheet.";
                return Result.Failed;
            }
            
            // Collect all view IDs we're interested in
            var viewsToCheck = new HashSet<ElementId>();
            int sheetCount = 0;
            
            foreach (var id in selectedIds)
            {
                var element = doc.GetElement(id);
                if (element is ViewSheet sheet)
                {
                    sheetCount++;
                    // Add the sheet itself
                    viewsToCheck.Add(sheet.Id);
                    
                    // Add all views placed on this sheet
                    // GetAllPlacedViews() should not trigger regeneration
                    var placedViews = sheet.GetAllPlacedViews();
                    foreach (var viewId in placedViews)
                    {
                        viewsToCheck.Add(viewId);
                    }
                }
            }
            
            if (sheetCount == 0)
            {
                message = "No sheets found in the current selection.";
                return Result.Failed;
            }
            
            // Collect ALL revision clouds in the document
            var revisionCloudsToSelect = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(RevisionCloud))
                .Cast<RevisionCloud>()
                .Where(rc => viewsToCheck.Contains(rc.OwnerViewId))
                .Select(rc => rc.Id)
                .ToList();
            
            if (!revisionCloudsToSelect.Any())
            {
                TaskDialog.Show("Result", 
                    $"No revision clouds found on the selected {sheetCount} sheet(s) or their views.");
                return Result.Succeeded;
            }
            
            // Select the revision clouds using SelectionModeManager
            uidoc.SetSelectionIds(revisionCloudsToSelect);
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Error: {ex.Message}";
            return Result.Failed;
        }
    }
}
