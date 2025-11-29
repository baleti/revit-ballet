using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

/// <summary>
/// Utility class to manage the select-in-links mode state.
/// This controls whether selection commands should search scope boxes within linked models.
/// </summary>
public static class SelectInLinksMode
{
    private static readonly string StateFilePath = PathHelper.GetRuntimeFilePath("select-in-links-mode");

    /// <summary>
    /// Gets the current select-in-links mode state.
    /// </summary>
    /// <returns>True if selection commands should check scope boxes in linked models, false otherwise. Default is false.</returns>
    public static bool IsEnabled()
    {
        try
        {
            if (!File.Exists(StateFilePath))
            {
                return false; // Default: don't search in linked models
            }

            string content = File.ReadAllText(StateFilePath).Trim().ToLower();
            return content == "true";
        }
        catch
        {
            return false; // Default on error
        }
    }

    /// <summary>
    /// Sets the select-in-links mode state.
    /// </summary>
    /// <param name="enabled">True to enable checking scope boxes in linked models, false to disable.</param>
    public static void SetEnabled(bool enabled)
    {
        try
        {
            // Write state to file (PathHelper ensures directory exists)
            File.WriteAllText(StateFilePath, enabled ? "true" : "false");
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to save select-in-links mode state: {ex.Message}", ex);
        }
    }
}

[Transaction(TransactionMode.Manual)]
public class ToggleSelectInLinksMode : IExternalCommand
{
    public Result Execute(ExternalCommandData cData, ref string message, ElementSet elements)
    {
        try
        {
            var uiDoc = cData.Application.ActiveUIDocument;

            // Get current state
            bool currentState = SelectInLinksMode.IsEnabled();

            // Build DataGrid entries for the two states
            var entries = new List<Dictionary<string, object>>();

            entries.Add(new Dictionary<string, object>
            {
                { "State", "Enabled" },
                { "Description", "Filter commands will check scope boxes in linked models" },
                { "IsEnabled", true }
            });

            entries.Add(new Dictionary<string, object>
            {
                { "State", "Disabled" },
                { "Description", "Filter commands will NOT check scope boxes in linked models" },
                { "IsEnabled", false }
            });

            // Determine initial selection based on current state
            var initialSelection = new List<int> { currentState ? 0 : 1 };

            // Show DataGrid
            var propertyNames = new List<string> { "State", "Description" };
            var chosenRows = CustomGUIs.DataGrid(
                entries,
                propertyNames,
                spanAllScreens: false,
                initialSelectionIndices: initialSelection);

            if (chosenRows == null || chosenRows.Count == 0)
            {
                TaskDialog.Show("Info", "Select-in-links mode not changed.");
                return Result.Cancelled;
            }

            // Get the selected state
            var selectedRow = chosenRows[0];
            bool newState = (bool)selectedRow["IsEnabled"];

            // Only update if state changed
            if (newState != currentState)
            {
                SelectInLinksMode.SetEnabled(newState);
                string modeDescription = newState
                    ? "enabled (filter commands will check scope boxes in linked models)"
                    : "disabled (filter commands will NOT check scope boxes in linked models)";
                TaskDialog.Show("Select-in-Links Mode", $"Select-in-links mode {modeDescription}");
            }
            else
            {
                TaskDialog.Show("Info", "Select-in-links mode unchanged.");
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Error changing select-in-links mode: {ex.Message}";
            return Result.Failed;
        }
    }
}
