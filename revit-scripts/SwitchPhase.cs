using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

[Transaction(TransactionMode.Manual)]
public class SwitchPhase : IExternalCommand
{
    // Wrapper class for displaying phases with a Number column.
    public class PhaseWrapper
    {
        public int i { get; set; }
        public string Name { get; set; }
        public Phase Phase { get; set; }
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Helper: Wrap a string in quotes if it contains spaces.
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

        // Get the active document, UIDocument, and active view.
        Document doc = commandData.Application.ActiveUIDocument.Document;
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        View activeView = uidoc.ActiveView;

        // Retrieve phases in their natural chronological order from the document's Phases collection.
        List<Phase> phases = new List<Phase>();
        foreach (Phase phase in doc.Phases)
        {
            phases.Add(phase);
        }

        if (phases == null || phases.Count == 0)
        {
            TaskDialog.Show("Error", "No phases found in the project.");
            return Result.Failed;
        }

        // Determine the current phase of the active view (if applicable).
        ElementId currentPhaseId = ElementId.InvalidElementId;
        Parameter phaseParam = activeView.get_Parameter(BuiltInParameter.VIEW_PHASE);
        if (phaseParam != null)
        {
            currentPhaseId = phaseParam.AsElementId();
        }

        // Create a list of PhaseWrapper objects with a Number column (1-based index).
        List<PhaseWrapper> phaseWrappers = phases.Select((p, index) => new PhaseWrapper
        {
            i = index + 1,
            Name = p.Name,
            Phase = p
        }).ToList();

        // Define the property names for the DataGrid; "Number" appears first.
        List<string> propertyNames = new List<string> { "i", "Name" };

        // Determine the current phase index in the original phases list.
        int selectedIndex = phases.FindIndex(p => p.Id == currentPhaseId);
        List<int> initialSelectionIndices = selectedIndex >= 0 ? new List<int> { selectedIndex } : new List<int>();

        // Display the phases using the custom DataGrid UI.
        List<PhaseWrapper> selectedPhaseWrappers = CustomGUIs.DataGrid(phaseWrappers, propertyNames, initialSelectionIndices);
        if (selectedPhaseWrappers == null || selectedPhaseWrappers.Count == 0)
            return Result.Failed;

        // Retrieve the underlying Phase from the selected wrapper.
        Phase chosenPhase = selectedPhaseWrappers.First().Phase;

        // Start a transaction to update the active view's phase parameter.
        using (Transaction tx = new Transaction(doc, "Switch Phase"))
        {
            tx.Start();
            if (phaseParam != null && !phaseParam.IsReadOnly)
            {
                phaseParam.Set(chosenPhase.Id);
            }
            else
            {
                TaskDialog.Show("Error", "Cannot change the phase for the active view.");
                tx.RollBack();
                return Result.Failed;
            }
            tx.Commit();
        }

        // Logging the phase change.
        try
        {
            string projectName = doc != null ? doc.Title : "UnknownProject";
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logFolderPath = Path.Combine(appDataPath, "revit-scripts", "LogPhaseChanges");
            if (!Directory.Exists(logFolderPath))
            {
                Directory.CreateDirectory(logFolderPath);
            }
            string logFilePath = Path.Combine(logFolderPath, $"{projectName}");

            // Read existing log lines, if any.
            List<string> logLines = new List<string>();
            bool logFileExists = File.Exists(logFilePath);
            if (logFileExists)
            {
                logLines = File.ReadAllLines(logFilePath).ToList();
            }

            string activeViewName = activeView.Name;
            int chosenPhaseIdInt = chosenPhase.Id.IntegerValue;
            string formattedViewName = FormatString(activeViewName);

            // Remove any existing log entry for the chosen ("to") phase in the current view.
            logLines = logLines.Where(line =>
            {
                List<string> tokens = TokenizeLine(line);
                if (tokens.Count < 3)
                    return true;
                if (!int.TryParse(tokens[0], out int loggedPhaseId))
                    return true;
                string loggedViewName = Unquote(tokens[2]);
                // Remove the line if it belongs to the active view and its phase equals the chosen phase.
                if (string.Equals(loggedViewName, activeViewName, StringComparison.Ordinal) &&
                    loggedPhaseId == chosenPhaseIdInt)
                    return false;
                return true;
            }).ToList();

            // Check if any entries exist for the current view.
            bool hasEntriesForCurrentView = logLines.Any(line =>
            {
                List<string> tokens = TokenizeLine(line);
                if (tokens.Count < 3)
                    return false;
                string loggedViewName = Unquote(tokens[2]);
                return string.Equals(loggedViewName, activeViewName, StringComparison.Ordinal);
            });

            // Format the new ("to") phase log line.
            string toPhaseLogLine = $"{chosenPhase.Id.IntegerValue} {FormatString(chosenPhase.Name)} {formattedViewName}";

            // Prepare the "from" phase log line.
            Phase fromPhase = phases.FirstOrDefault(p => p.Id == currentPhaseId);
            string fromPhaseLogLine = fromPhase != null 
                ? $"{fromPhase.Id.IntegerValue} {FormatString(fromPhase.Name)} {formattedViewName}" 
                : null;

            // Decide what new lines to add:
            if (!logFileExists || logLines.Count == 0 || !hasEntriesForCurrentView)
            {
                // If the file is missing/empty or there are no entries for the current view,
                // add both a "from" phase entry and the "to" phase entry.
                var newEntries = new List<string>();
                if (fromPhaseLogLine != null)
                    newEntries.Add(fromPhaseLogLine);
                newEntries.Add(toPhaseLogLine);
                // Append these new entries to the log (preserving entries for other views).
                logLines.AddRange(newEntries);
            }
            else
            {
                // Otherwise, simply append the new ("to") phase entry.
                logLines.Add(toPhaseLogLine);
            }

            // Write out the complete updated log file.
            File.WriteAllLines(logFilePath, logLines);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Logging Error", $"Failed to log phase change: {ex.Message}");
        }

        return Result.Succeeded;
    }
}
