using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SwitchToLastPhase : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the active UIDocument and Document.
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        string projectName = doc != null ? doc.Title : "UnknownProject";

        // Build the log file path from app data under "revit-scripts\LogPhaseChanges".
        string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string logFolderPath = Path.Combine(appDataPath, "revit-scripts", "LogPhaseChanges");
        if (!Directory.Exists(logFolderPath))
        {
            message = "Log folder does not exist.";
            return Result.Failed;
        }
        string logFilePath = Path.Combine(logFolderPath, $"{projectName}");

        if (!File.Exists(logFilePath))
        {
            message = "Log file does not exist.";
            return Result.Failed;
        }

        // Read all lines from the log file.
        List<string> logLines = File.ReadLines(logFilePath).ToList();
        if (logLines.Count < 2)
        {
            message = "Not enough entries in the log file.";
            return Result.Failed;
        }

        // Loop from the second-to-last entry backward to find a valid phase.
        Phase chosenPhase = null;
        int chosenIndex = -1;
        for (int i = logLines.Count - 2; i >= 0; i--)
        {
            string[] parts = logLines[i].Split(new[] { ' ' }, 2);
            if (parts.Length < 2)
                continue;

            string phaseIdStr = parts[0];
            if (int.TryParse(phaseIdStr, out int phaseId))
            {
                ElementId phaseElementId = new ElementId(phaseId);
                Phase phase = doc.GetElement(phaseElementId) as Phase;
                if (phase != null)
                {
                    chosenPhase = phase;
                    chosenIndex = i;
                    break;
                }
            }
        }

        if (chosenPhase == null)
        {
            message = "No valid phase found in the log file.";
            return Result.Failed;
        }

        // Get the active view and its phase parameter.
        View activeView = uidoc.ActiveView;
        Parameter phaseParam = activeView.get_Parameter(BuiltInParameter.VIEW_PHASE);
        if (phaseParam == null || phaseParam.IsReadOnly)
        {
            message = "Active view does not support phase changes.";
            return Result.Failed;
        }

        // Update the active view's phase parameter within a transaction.
        using (Transaction tx = new Transaction(doc, "Switch Phase"))
        {
            tx.Start();
            phaseParam.Set(chosenPhase.Id);
            tx.Commit();
        }

        // "Move" the last log entry back by one by swapping it with the chosen entry.
        int lastIndex = logLines.Count - 1;
        if (chosenIndex != lastIndex)
        {
            string temp = logLines[lastIndex];
            logLines[lastIndex] = logLines[chosenIndex];
            logLines[chosenIndex] = temp;
            try
            {
                File.WriteAllLines(logFilePath, logLines);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Logging Error", $"Failed to update log file: {ex.Message}");
                // Continue execution even if the log update fails.
            }
        }

        return Result.Succeeded;
    }
}
