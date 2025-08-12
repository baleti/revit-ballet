using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class ExportSchedule : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Retrieve all schedule views in the document, excluding those that start with '<'
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

        // Show custom GUI to select a schedule
        List<string> propertyNames = new List<string> { "Name", "Id" };
        var selectedSchedules = CustomGUIs.DataGrid(schedules, propertyNames, false);

        if (selectedSchedules.Count == 0)
        {
            return Result.Cancelled;
        }

        var selectedScheduleId = new ElementId(Convert.ToInt32(selectedSchedules[0]["Id"]));
        ViewSchedule selectedSchedule = doc.GetElement(selectedScheduleId) as ViewSchedule;

        if (selectedSchedule != null)
        {
            // Define export options
            var options = new ViewScheduleExportOptions
            {
                Title = false,
                TextQualifier = ExportTextQualifier.None,
                FieldDelimiter = "\t",
                ColumnHeaders = ExportColumnHeaders.None
            };

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string exportConfigFolder = Path.Combine(appDataPath, "revit-scripts");
            string exportConfigPath = Path.Combine(exportConfigFolder, $"ExportSchedule - {selectedSchedule.Name}.txt");

            string exportPath;

            // Check if previously saved path exists
            if (File.Exists(exportConfigPath))
            {
                exportPath = File.ReadAllText(exportConfigPath);
            }
            else
            {
                // Prompt user to choose a file location
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    saveFileDialog.FileName = $"{selectedSchedule.Name}.txt";
                    saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        exportPath = saveFileDialog.FileName;

                        // Ensure the directory exists and save the path
                        if (!Directory.Exists(exportConfigFolder))
                        {
                            Directory.CreateDirectory(exportConfigFolder);
                        }

                        File.WriteAllText(exportConfigPath, exportPath);
                    }
                    else
                    {
                        TaskDialog.Show("Export Cancelled", "No file location was chosen.");
                        return Result.Cancelled;
                    }
                }
            }

            // Ensure the directory exists
            string exportFolder = Path.GetDirectoryName(exportPath);
            if (!Directory.Exists(exportFolder))
            {
                Directory.CreateDirectory(exportFolder);
            }

            // Activate the selected schedule view
            uiDoc.ActiveView = selectedSchedule;

            using (Transaction transaction = new Transaction(doc, "Export Schedule"))
            {
                transaction.Start();

                // Correctly call the Export method with the folder and file name
                selectedSchedule.Export(exportFolder, Path.GetFileName(exportPath), options);

                transaction.Commit();
            }

            return Result.Succeeded;
        }

        message = "Selected schedule could not be found.";
        return Result.Failed;
    }
}
