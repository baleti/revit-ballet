using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class OpenLastSessionViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        // Get the project name
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        string projectName = doc != null ? doc.Title : "UnknownProject";

        // Construct the log file path
        string logFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
            "revit-scripts", 
            "LogViewChanges",
            $"{projectName}.last"
        );

        if (!File.Exists(logFilePath))
        {
            TaskDialog.Show("Error", "Log file does not exist.");
            return Result.Failed;
        }

        // Read the data from the log file
        List<Dictionary<string, object>> savedViewTitles = new List<Dictionary<string, object>>();
        using (StreamReader reader = new StreamReader(logFilePath))
        {
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                var parts = line.Split(new[] { ' ' }, 2); // Split into two parts: ID and Title
                if (parts.Length == 2)
                {
                    var entry = new Dictionary<string, object> { { "Title", parts[1] } };
                    savedViewTitles.Add(entry);
                }
            }
        }

        // Display the saved views using CustomGUIs.DataGrid
        var selectedViews = CustomGUIs.DataGrid(savedViewTitles, new List<string> { "Title" }, false);

        // Open the selected views in Revit
        UIApplication uiapp = commandData.Application;
        foreach (var viewEntry in selectedViews)
        {
            if (viewEntry["Title"] == null)
                continue;

            string viewTitle = viewEntry["Title"].ToString();
            Autodesk.Revit.DB.View view = new FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .FirstOrDefault(v => v.Title.Equals(viewTitle, StringComparison.OrdinalIgnoreCase));

            if (view != null)
            {
                uidoc.RequestViewChange(view);
            }
        }

        return Result.Succeeded;
    }
}
