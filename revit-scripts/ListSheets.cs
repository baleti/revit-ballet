using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ListSheets : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Step 1: Prepare a list of sheets and their relevant parameters in the current project
        List<Dictionary<string, object>> sheetEntries = new List<Dictionary<string, object>>();
        Dictionary<string, ViewSheet> sheetElementMap = new Dictionary<string, ViewSheet>(); // Map unique keys to ViewSheets

        // Collect all ViewSheet elements in the project
        var viewSheets = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet))
            .Cast<ViewSheet>()
            .ToList();

        // Iterate through the ViewSheets and collect their relevant parameters
        foreach (var sheet in viewSheets)
        {
            // Get relevant parameters
            string sheetNumber = sheet.SheetNumber;
            string sheetName = sheet.Name;
            string currentRevision = GetParameterValue(sheet, BuiltInParameter.SHEET_CURRENT_REVISION);
            string currentRevisionIssuedTo = GetParameterValue(sheet, BuiltInParameter.SHEET_CURRENT_REVISION_ISSUED_TO);
            string currentRevisionDescription = GetParameterValue(sheet, BuiltInParameter.SHEET_CURRENT_REVISION_DESCRIPTION);

            // Create an entry for the sheet
            var entry = new Dictionary<string, object>
            {
                { "Sheet Number", sheetNumber },
                { "Sheet Name", sheetName },
                { "Current Revision", currentRevision },
                { "Current Revision Issued To", currentRevisionIssuedTo },
                { "Current Revision Description", currentRevisionDescription }
            };

            // Store the ViewSheet with a unique key (SheetNumber:SheetName)
            string uniqueKey = $"{sheetNumber}:{sheetName}";
            sheetElementMap[uniqueKey] = sheet;

            sheetEntries.Add(entry);
        }

        // Step 2: Display the list of sheets using CustomGUIs.DataGrid
        var propertyNames = new List<string>
        {
            "Sheet Number", 
            "Sheet Name", 
            "Current Revision", 
            "Current Revision Issued To", 
            "Current Revision Description"
        };

        var selectedEntries = CustomGUIs.DataGrid(sheetEntries, propertyNames, spanAllScreens: false);

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled; // No selection made
        }

        // Step 3: Collect the selected sheets using Sheet Number and Sheet Name
        List<ViewSheet> selectedSheets = selectedEntries
            .Select(entry =>
            {
                string sheetNumber = entry["Sheet Number"].ToString();
                string sheetName = entry["Sheet Name"].ToString();
                string uniqueKey = $"{sheetNumber}:{sheetName}";

                // Match the sheet from the dictionary
                if (sheetElementMap.ContainsKey(uniqueKey))
                {
                    return sheetElementMap[uniqueKey];
                }
                return null;
            })
            .Where(sheet => sheet != null)
            .ToList();

        // Step 4: Open the selected sheets
        foreach (ViewSheet sheet in selectedSheets)
        {
            if (sheet != null)
            {
                uidoc.ActiveView = sheet; // Set the active view to the selected sheet
            }
        }

        return Result.Succeeded;
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
