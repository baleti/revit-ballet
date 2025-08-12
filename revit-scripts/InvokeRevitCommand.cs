using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class InvokeRevitCommand : IExternalCommand
{
    private const string FolderName = "revit-scripts";
    private const string HistoryFileName = "InvokeRevitCommand-history";
    private static readonly string ConfigFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName);
    private static readonly string HistoryFilePath = Path.Combine(ConfigFolderPath, HistoryFileName);

    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        // Get all built-in Revit postable command identifiers
        var commandEntries = new List<Dictionary<string, object>>();
        var commandIds = new Dictionary<string, RevitCommandId>();
        
        // Get all values from PostableCommand enum
        foreach (var commandName in Enum.GetNames(typeof(PostableCommand)))
        {
            var postableCommand = (PostableCommand)Enum.Parse(typeof(PostableCommand), commandName);
            var commandId = RevitCommandId.LookupPostableCommandId(postableCommand);
            if (commandId != null)
            {
                commandEntries.Add(new Dictionary<string, object> { { "Command", commandName } });
                commandIds[commandName] = commandId; // Store the command id
            }
        }
        
        // Use CustomGUIs.DataGrid to let user select a command
        List<string> propertyNames = new List<string> { "Command" };
        var selectedCommand = CustomGUIs.DataGrid(commandEntries, propertyNames, false).FirstOrDefault();
        
        if (selectedCommand != null)
        {
            string commandName = selectedCommand["Command"].ToString();
            RevitCommandId commandId = commandIds[commandName]; // Get the command id
            
            try
            {
                // Append to history file
                AppendToCommandHistory(commandName);
                
                // Invoke the selected command using PostCommand
                commandData.Application.PostCommand(commandId);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"An error occurred: {ex.Message}";
                return Result.Failed;
            }
        }
        
        return Result.Failed;
    }

    private void AppendToCommandHistory(string commandName)
    {
        try
        {
            // Ensure directory exists
            if (!Directory.Exists(ConfigFolderPath))
            {
                Directory.CreateDirectory(ConfigFolderPath);
            }

            // Append command to history file
            using (StreamWriter sw = File.AppendText(HistoryFilePath))
            {
                sw.WriteLine(commandName);
            }
        }
        catch (Exception ex)
        {
            // Log error but don't fail the command execution
            TaskDialog.Show("Warning", $"Failed to update command history: {ex.Message}");
        }
    }
}
