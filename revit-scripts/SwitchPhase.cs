using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SwitchPhase : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the active document, UIDocument, and active view.
        Document doc = commandData.Application.ActiveUIDocument.Document;
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        View activeView = uidoc.ActiveView;

        // Retrieve all phases in the project and sort them alphabetically.
        List<Phase> phases = new FilteredElementCollector(doc)
            .OfClass(typeof(Phase))
            .Cast<Phase>()
            .OrderBy(p => p.Name)
            .ToList();

        if (phases == null || phases.Count == 0)
        {
            TaskDialog.Show("Error", "No phases found in the project.");
            return Result.Failed;
        }

        // Define the property names to display in the custom UI (here, only the phase Name).
        List<string> propertyNames = new List<string> { "Name" };

        // Determine the current phase of the active view (if applicable).
        ElementId currentPhaseId = ElementId.InvalidElementId;
        Parameter phaseParam = activeView.get_Parameter(BuiltInParameter.VIEW_PHASE);
        if (phaseParam != null)
        {
            currentPhaseId = phaseParam.AsElementId();
        }
        int selectedIndex = phases.FindIndex(p => p.Id == currentPhaseId);
        List<int> initialSelectionIndices = selectedIndex >= 0 ? new List<int> { selectedIndex } : new List<int>();

        // Display the phases using a custom DataGrid UI.
        List<Phase> selectedPhases = CustomGUIs.DataGrid(phases, propertyNames, initialSelectionIndices);
        if (selectedPhases == null || selectedPhases.Count == 0)
            return Result.Failed;

        Phase chosenPhase = selectedPhases.First();

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

        // Log the chosen phase so that SwitchToLastPhase can reference it.
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
            // Format the log entry as "PhaseID PhaseName".
            string logEntry = $"{chosenPhase.Id.IntegerValue} {chosenPhase.Name}";
            File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Logging Error", $"Failed to log phase change: {ex.Message}");
            // Continue execution even if logging fails.
        }

        return Result.Succeeded;
    }
}
