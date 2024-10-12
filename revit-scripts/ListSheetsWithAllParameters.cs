using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class ListSheetsWithAllParameters : IExternalCommand
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

        // Create a dictionary to map sheets to their entries
        Dictionary<string, ViewSheet> sheetMap = new Dictionary<string, ViewSheet>();
        List<Dictionary<string, object>> sheetEntries = new List<Dictionary<string, object>>();

        // Collect all unique parameter names across all sheets
        HashSet<string> uniqueParameterNames = new HashSet<string>();

        foreach (var sheet in sheets)
        {
            Dictionary<string, object> sheetData = new Dictionary<string, object>();

            // Add basic info like Title and Sheet Number
            sheetData.Add("Sheet Title", sheet.Title);
            sheetData.Add("Sheet Number", sheet.SheetNumber);

            // Create a unique key for each sheet (e.g., "SheetNumber|SheetTitle") for easy lookup later
            string sheetKey = sheet.SheetNumber + "|" + sheet.Title;
            sheetMap.Add(sheetKey, sheet); // Map the unique key to the actual ViewSheet object

            // Iterate over all parameters for each sheet and add them to the dictionary
            foreach (Parameter param in sheet.Parameters)
            {
                string paramName = param.Definition.Name;
                if (!sheetData.ContainsKey(paramName)) // Avoid duplicates
                {
                    sheetData[paramName] = GetParameterValue(param);
                    uniqueParameterNames.Add(paramName);
                }
            }

            sheetEntries.Add(sheetData);
        }

        // Convert HashSet to List to display all unique parameter names as column headers in the DataGrid
        List<string> sheetProperties = new List<string> { "Sheet Title", "Sheet Number" };
        sheetProperties.AddRange(uniqueParameterNames);

        // Display the DataGrid and let the user select one or more sheets
        List<Dictionary<string, object>> selectedSheets = CustomGUIs.DataGrid(sheetEntries, sheetProperties, false);

        if (selectedSheets.Count == 0)
        {
            return Result.Cancelled;
        }

        // Open the selected sheets
        foreach (var selectedSheetEntry in selectedSheets)
        {
            // Build the same unique key from the selected sheet's title and number
            string selectedSheetKey = selectedSheetEntry["Sheet Number"].ToString() + "|" + selectedSheetEntry["Sheet Title"].ToString();

            // Use the key to retrieve the actual ViewSheet from the map
            if (sheetMap.ContainsKey(selectedSheetKey))
            {
                ViewSheet selectedSheet = sheetMap[selectedSheetKey];

                if (selectedSheet != null)
                {
                    uiDoc.ActiveView = selectedSheet;
                }
            }
        }

        return Result.Succeeded;
    }

    // Helper method to extract parameter values based on parameter type
    private object GetParameterValue(Parameter param)
    {
        switch (param.StorageType)
        {
            case StorageType.Double:
                return param.AsDouble().ToString("0.##"); // Formatting double values
            case StorageType.Integer:
                return param.AsInteger();
            case StorageType.String:
                return param.AsString();
            case StorageType.ElementId:
                ElementId id = param.AsElementId();
                return id.IntegerValue != -1 ? id.IntegerValue.ToString() : "None";
            default:
                return "Unknown";
        }
    }
}
