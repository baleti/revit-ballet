using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SwitchToLastView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        string projectName = doc != null ? doc.Title : "UnknownProject";

        string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string logFilePath = Path.Combine(appDataPath, "revit-scripts", "LogViewChanges", $"{projectName}");

        if (!File.Exists(logFilePath))
        {
            message = "Log file does not exist.";
            return Result.Failed;
        }

        List<string> logLines = File.ReadLines(logFilePath).ToList();

        if (logLines.Count < 2)
        {
            message = "Not enough entries in the log file.";
            return Result.Failed;
        }

        for (int i = logLines.Count - 2; i >= 0; i--)
        {
            string lastViewIdStr = logLines[i].Split(' ')[0];

            if (int.TryParse(lastViewIdStr, out int lastViewId))
            {
                ElementId viewElementId = new ElementId(lastViewId);
                View lastView = doc.GetElement(viewElementId) as View;

                if (lastView != null)
                {
                    uidoc.ActiveView = lastView;
                    return Result.Succeeded;
                }
            }
        }

        message = "No valid view found in the log file.";
        return Result.Failed;
    }
}
