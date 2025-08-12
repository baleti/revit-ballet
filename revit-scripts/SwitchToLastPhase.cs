using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

[Transaction(TransactionMode.Manual)]
public class SwitchToLastPhase : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Helper: Formats a string by adding quotes if it contains spaces.
        string FormatString(string input)
        {
            return input.Contains(" ") ? $"\"{input}\"" : input;
        }

        // Helper: Removes surrounding quotes from a string.
        string Unquote(string text)
        {
            if (text.StartsWith("\"") && text.EndsWith("\""))
            {
                return text.Substring(1, text.Length - 2);
            }
            return text;
        }

        // Helper: Tokenizes a log line into tokens, respecting quoted substrings.
        List<string> TokenizeLine(string line)
        {
            var tokens = new List<string>();
            foreach (Match m in Regex.Matches(line, @"(""[^""]+""|\S+)"))
            {
                tokens.Add(m.Value);
            }
            return tokens;
        }

        // Get the active UIDocument and Document.
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        string projectName = doc != null ? doc.Title : "UnknownProject";

        // Build the log file path.
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
        List<string> logLines = File.ReadAllLines(logFilePath).ToList();
        if (logLines.Count == 0)
        {
            message = "Log file is empty.";
            return Result.Failed;
        }

        // Get active view and its phase parameter.
        View activeView = uidoc.ActiveView;
        Parameter phaseParam = activeView.get_Parameter(BuiltInParameter.VIEW_PHASE);
        if (phaseParam == null || phaseParam.IsReadOnly)
        {
            message = "Active view does not support phase changes.";
            return Result.Failed;
        }

        // Capture the current phase from the active view (to be logged later).
        ElementId currentPhaseId = phaseParam.AsElementId();
        Phase currentPhase = doc.GetElement(currentPhaseId) as Phase;
        if (currentPhase == null)
        {
            message = "Current phase not found.";
            return Result.Failed;
        }

        // Build a list of log entries for the current view.
        // Each log line is expected to have the format: "phaseId phaseName viewName"
        var matchingEntries = new List<(int lineIndex, int phaseId, string phaseName, string viewName)>();
        for (int i = 0; i < logLines.Count; i++)
        {
            string line = logLines[i];
            List<string> tokens = TokenizeLine(line);
            if (tokens.Count != 3)
                continue;
            if (!int.TryParse(tokens[0], out int loggedPhaseId))
                continue;
            string loggedPhaseName = Unquote(tokens[1]);
            string loggedViewName = Unquote(tokens[2]);
            if (string.Equals(loggedViewName, activeView.Name, StringComparison.Ordinal))
            {
                matchingEntries.Add((i, loggedPhaseId, loggedPhaseName, loggedViewName));
            }
        }

        Phase chosenPhase = null;
        // We'll also remember the current phase (before switching) to log later.
        // If there are at least two matching entries, swap the last two.
        if (matchingEntries.Count >= 2)
        {
            // Get the last two entries for the current view.
            var secondLastEntry = matchingEntries[matchingEntries.Count - 2];
            var lastEntry = matchingEntries[matchingEntries.Count - 1];

            // Use the older of the two (the second last) as the chosen phase.
            chosenPhase = doc.GetElement(new ElementId(secondLastEntry.phaseId)) as Phase;
            if (chosenPhase == null)
            {
                message = "Invalid phase found in the log.";
                return Result.Failed;
            }

            // Begin transaction to update the active view's phase.
            using (Transaction tx = new Transaction(doc, "Switch Phase"))
            {
                tx.Start();
                phaseParam.Set(chosenPhase.Id);
                tx.Commit();
            }

            // Remove both the second last and last entries from the log (for this view).
            // Remove by descending index to avoid reindexing issues.
            logLines.RemoveAt(lastEntry.lineIndex);
            logLines.RemoveAt(secondLastEntry.lineIndex);

            // Append two new log entries for this view:
            // First, log the phase that was active before switching (currentPhase).
            string currentLogLine = $"{currentPhase.Id.IntegerValue} {FormatString(currentPhase.Name)} {FormatString(activeView.Name)}";
            // Second, log the chosen phase (the one we just switched to).
            string chosenLogLine = $"{chosenPhase.Id.IntegerValue} {FormatString(chosenPhase.Name)} {FormatString(activeView.Name)}";

            logLines.Add(currentLogLine);
            logLines.Add(chosenLogLine);
        }
        else if (matchingEntries.Count == 1)
        {
            // Only one matching entry exists; use it as the chosen phase.
            var entry = matchingEntries[0];
            chosenPhase = doc.GetElement(new ElementId(entry.phaseId)) as Phase;
            if (chosenPhase == null)
            {
                message = "Invalid phase found in the log.";
                return Result.Failed;
            }
            // Begin transaction to update the active view's phase.
            using (Transaction tx = new Transaction(doc, "Switch Phase"))
            {
                tx.Start();
                phaseParam.Set(chosenPhase.Id);
                tx.Commit();
            }
            // Append a new log entry for the current phase.
            string currentLogLine = $"{currentPhase.Id.IntegerValue} {FormatString(currentPhase.Name)} {FormatString(activeView.Name)}";
            logLines.Add(currentLogLine);
        }
        else
        {
            message = "No valid log entry found for the current view.";
            return Result.Failed;
        }

        // Update the log file.
        try
        {
            File.WriteAllLines(logFilePath, logLines);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Logging Error", $"Failed to update log file: {ex.Message}");
            // Continue execution even if the log update fails.
        }

        return Result.Succeeded;
    }
}
