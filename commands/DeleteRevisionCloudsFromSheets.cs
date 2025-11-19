using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class DeleteRevisionCloudsFromSheets : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Collect all sheets in the project
        IList<ViewSheet> sheets = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet))
            .Cast<ViewSheet>()
            .ToList();

        if (sheets.Count == 0)
        {
            TaskDialog.Show("Sheets", "No sheets found in the project.");
            return Result.Cancelled;
        }

        // Define the parameters to display
        List<string> sheetProperties = new List<string>
        {
            "Sheet Number",
            "Sheet Name",
            "Current Revision",
            "Current Revision Issued To",
            "Current Revision Description"
        };

        // Create a dictionary to map sheets to their entries
        Dictionary<string, ViewSheet> sheetMap = new Dictionary<string, ViewSheet>();
        List<Dictionary<string, object>> sheetEntries = new List<Dictionary<string, object>>();

        foreach (var sheet in sheets)
        {
            Dictionary<string, object> sheetData = new Dictionary<string, object>();

            // Add the specified parameters
            sheetData.Add("Sheet Number", sheet.SheetNumber);
            sheetData.Add("Sheet Name", sheet.Name);

            // Get the Current Revision parameter values
            string currentRevision = GetParameterValue(sheet, BuiltInParameter.SHEET_CURRENT_REVISION);
            string currentRevisionIssuedTo = GetParameterValue(sheet, BuiltInParameter.SHEET_CURRENT_REVISION_ISSUED_TO);
            string currentRevisionDescription = GetParameterValue(sheet, BuiltInParameter.SHEET_CURRENT_REVISION_DESCRIPTION);

            sheetData.Add("Current Revision", currentRevision);
            sheetData.Add("Current Revision Issued To", currentRevisionIssuedTo);
            sheetData.Add("Current Revision Description", currentRevisionDescription);

            // Create a unique key for each sheet (e.g., "SheetNumber|SheetName") for easy lookup later
            string sheetKey = sheet.SheetNumber + "|" + sheet.Name;
            sheetMap.Add(sheetKey, sheet); // Map the unique key to the actual ViewSheet object

            sheetEntries.Add(sheetData);
        }

        // Display the DataGrid and let the user select one or more sheets
        List<Dictionary<string, object>> selectedSheets = CustomGUIs.DataGrid(sheetEntries, sheetProperties, false);

        if (selectedSheets.Count == 0)
        {
            return Result.Cancelled;
        }

        // Begin a transaction to delete revision clouds
        using (Transaction t = new Transaction(doc, "Delete Revision Clouds"))
        {
            t.Start();

            // Delete revision clouds on the selected sheets
            foreach (var selectedSheetEntry in selectedSheets)
            {
                // Build the same unique key from the selected sheet's number and name
                string selectedSheetKey = selectedSheetEntry["Sheet Number"].ToString() + "|" + selectedSheetEntry["Sheet Name"].ToString();

                // Use the key to retrieve the actual ViewSheet from the map
                if (sheetMap.ContainsKey(selectedSheetKey))
                {
                    ViewSheet selectedSheet = sheetMap[selectedSheetKey];

                    if (selectedSheet != null)
                    {
                        // Delete revision clouds from the selected sheet
                        DeleteRevisionCloudsFromSheet(doc, selectedSheet);

                        // Also check views placed on the sheet and delete revision clouds from them
                        DeleteRevisionCloudsFromViewsOnSheet(doc, selectedSheet);
                    }
                }
            }

            t.Commit();
        }

        return Result.Succeeded;
    }

    // Helper method to delete revision clouds from a given sheet
    private void DeleteRevisionCloudsFromSheet(Document doc, ViewSheet sheet)
    {
        // Find all revision clouds in the sheet
        var revisionClouds = new FilteredElementCollector(doc)
            .OfClass(typeof(RevisionCloud))
            .WhereElementIsNotElementType()
            .Cast<RevisionCloud>()
            .Where(rc => rc.OwnerViewId == sheet.Id);

        // Collect all the revision cloud IDs into a list
        List<ElementId> cloudIdsToDelete = new List<ElementId>();

        foreach (var cloud in revisionClouds)
        {
            cloudIdsToDelete.Add(cloud.Id);
        }

        // Delete all the revision clouds after collecting their IDs
        foreach (var cloudId in cloudIdsToDelete)
        {
            doc.Delete(cloudId);
        }
    }

    // Helper method to delete revision clouds from views placed on the sheet
    private void DeleteRevisionCloudsFromViewsOnSheet(Document doc, ViewSheet sheet)
    {
        // Get the views placed on the sheet
        var viewportIds = sheet.GetAllViewports();
        
        foreach (ElementId vpId in viewportIds)
        {
            Viewport viewport = doc.GetElement(vpId) as Viewport;
            if (viewport != null)
            {
                View view = doc.GetElement(viewport.ViewId) as View;
                if (view != null)
                {
                    // Find all revision clouds in the view
                    var revisionClouds = new FilteredElementCollector(doc)
                        .OfClass(typeof(RevisionCloud))
                        .WhereElementIsNotElementType()
                        .Cast<RevisionCloud>()
                        .Where(rc => rc.OwnerViewId == view.Id);

                    // Collect all the revision cloud IDs into a list
                    List<ElementId> cloudIdsToDelete = new List<ElementId>();

                    foreach (var cloud in revisionClouds)
                    {
                        cloudIdsToDelete.Add(cloud.Id);
                    }

                    // Delete all the revision clouds after collecting their IDs
                    foreach (var cloudId in cloudIdsToDelete)
                    {
                        doc.Delete(cloudId);
                    }
                }
            }
        }
    }

    // Helper method to get parameter values
    private string GetParameterValue(ViewSheet sheet, BuiltInParameter paramId)
    {
        Parameter param = sheet.get_Parameter(paramId);
        if (param != null && param.HasValue)
        {
            return param.AsString() ?? param.AsValueString() ?? "";
        }
        else
        {
            return "";
        }
    }
}
