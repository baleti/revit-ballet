using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CloseCurrentView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        string projectName = doc != null ? doc.Title : "UnknownProject";

        // Remove the closed view from the log file
        string logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "revit-scripts", "LogViewChanges", $"{projectName}");
        if (File.Exists(logFilePath))
        {
            List<string> logEntries = File.ReadAllLines(logFilePath).ToList();
            string activeViewTitle = uidoc.ActiveView.Title;

            // Filter out the entry with the current active view's title
            logEntries = logEntries
                .Where(entry => !entry.Contains($" {activeViewTitle}"))
                .ToList();

            File.WriteAllLines(logFilePath, logEntries);
        }

        UIView activeUIView = uidoc.GetOpenUIViews().FirstOrDefault(u => u.ViewId == uidoc.ActiveView.Id);
        activeUIView.Close();

        return Result.Succeeded;
    }
}
