using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ExportSelectedViewsCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get all sheets in the project
        var sheets = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Sheets)
            .WhereElementIsNotElementType()
            .ToList();

        // Prepare data for the DataGrid GUI
        var sheetData = sheets.Select(sheet => new
        {
            SheetNumber = sheet.get_Parameter(BuiltInParameter.SHEET_NUMBER).AsString(),
            SheetName = sheet.get_Parameter(BuiltInParameter.SHEET_NAME).AsString(),
            Id = sheet.Id
        }).ToList();

        // Prompt user to select sheets
        var selectedSheets = CustomGUIs.DataGrid(sheetData, new List<string> { "SheetNumber", "SheetName" }, Title: "Select Sheets");

        if (selectedSheets == null || !selectedSheets.Any())
        {
            TaskDialog.Show("Export Sheets", "No sheets were selected.");
            return Result.Cancelled;
        }

        // Get the selected sheet IDs
        var selectedSheetIds = selectedSheets.Select(sheet => (ElementId)sheet.Id).ToList();

        // Use PostableCommand to save sheets as library views
        foreach (ElementId sheetId in selectedSheetIds)
        {
            ViewSheet sheet = doc.GetElement(sheetId) as ViewSheet;
            if (sheet != null)
            {
                // Activate the sheet
                uiDoc.ActiveView = sheet;

                // Lookup the PostableCommandId for SaveAsLibraryView
                RevitCommandId commandId = RevitCommandId.LookupPostableCommandId(PostableCommand.SaveAsLibraryView);
                if (uiApp.CanPostCommand(commandId))
                {
                    uiApp.PostCommand(commandId);
                }
                else
                {
                    TaskDialog.Show("Export Sheets", $"Unable to save sheet {sheet.SheetNumber} - {sheet.Name} as library view.");
                }
            }
        }

        TaskDialog.Show("Export Sheets", "Selected sheets have been exported as library views. Follow the on-screen prompts to complete the process.");
        return Result.Succeeded;
    }
}
