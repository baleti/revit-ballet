using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class InvokeRevitCommand : IExternalCommand
{
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
}
