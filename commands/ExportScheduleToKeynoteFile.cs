using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ExportScheduleToKeynoteFile : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get keynote file path
        KeynoteTable keynoteTable = KeynoteTable.GetKeynoteTable(doc);
        if (keynoteTable == null)
        {
            TaskDialog.Show("Error", "No keynote file is currently loaded.");
            return Result.Failed;
        }

        ExternalFileReference extFileRef = keynoteTable.GetExternalFileReference();
        ModelPath modelPath = extFileRef.GetAbsolutePath();
        string keynoteFilePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);

        var schedules = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSchedule))
            .Cast<ViewSchedule>()
            .Where(vs => !vs.Name.StartsWith("<") && !vs.Name.Contains("Keynote Legend"))
            .Select(vs => new Dictionary<string, object>
            {
                { "Name", vs.Name },
                { "Id", vs.Id.IntegerValue }
            })
            .ToList();

        List<string> propertyNames = new List<string> { "Name", "Id" };
        var selectedSchedules = CustomGUIs.DataGrid(schedules, propertyNames, false);

        if (selectedSchedules.Count == 0)
            return Result.Cancelled;

        var selectedScheduleId = new ElementId(Convert.ToInt32(selectedSchedules[0]["Id"]));
        ViewSchedule selectedSchedule = doc.GetElement(selectedScheduleId) as ViewSchedule;

        if (selectedSchedule == null)
        {
            message = "Selected schedule could not be found.";
            return Result.Failed;
        }

        var options = new ViewScheduleExportOptions
        {
            Title = false,
            TextQualifier = ExportTextQualifier.None,
            FieldDelimiter = "\t",
            ColumnHeaders = ExportColumnHeaders.None
        };

        uiDoc.ActiveView = selectedSchedule;

        using (Transaction transaction = new Transaction(doc, "Export Schedule"))
        {
            transaction.Start();
            selectedSchedule.Export(
                System.IO.Path.GetDirectoryName(keynoteFilePath),
                System.IO.Path.GetFileName(keynoteFilePath),
                options);
            transaction.Commit();
        }

        return Result.Succeeded;
    }
}
