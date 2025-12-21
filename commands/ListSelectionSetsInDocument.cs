#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class ListSelectionSetsInDocument : IExternalCommand
{
    // Static initializer to register the Name column handler for SelectionFilterElement
    static ListSelectionSetsInDocument()
    {
        RegisterSelectionSetNameHandler();
    }

    private static void RegisterSelectionSetNameHandler()
    {
        // Register handler for "Name" column to edit SelectionFilterElement names
        CustomGUIs.ColumnHandlerRegistry.Register(new CustomGUIs.ColumnHandler
        {
            ColumnName = "Name",
            IsEditable = true,
            Description = "Selection set name",
            Validator = CustomGUIs.ColumnValidators.All(
                CustomGUIs.ColumnValidators.NotEmpty,
                CustomGUIs.ColumnValidators.NoInvalidCharacters,
                CustomGUIs.ColumnValidators.NoLeadingTrailingWhitespace
            ),
            Getter = (elem, doc) =>
            {
                if (elem is SelectionFilterElement selectionSet)
                    return selectionSet.Name;
                return elem?.Name ?? "";
            },
            Setter = (elem, doc, newValue) =>
            {
                if (elem is SelectionFilterElement selectionSet)
                {
                    string strValue = newValue?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(strValue))
                        return false;

                    try
                    {
                        selectionSet.Name = strValue;
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }
                return false;
            }
        });
    }

    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get all selection sets
        var selectionSets = new FilteredElementCollector(doc)
            .OfClass(typeof(SelectionFilterElement))
            .Cast<SelectionFilterElement>()
            .OrderBy(s => s.Name)
            .ToList();

        if (selectionSets.Count == 0)
        {
            TaskDialog.Show("List Selection Sets",
                "No selection sets found in the current document.");
            return Result.Cancelled;
        }

        // Prepare data for DataGrid
        var gridData = new List<Dictionary<string, object>>();
        foreach (var set in selectionSets)
        {
            gridData.Add(new Dictionary<string, object>
            {
                ["Name"] = set.Name,
                ["Count"] = set.GetElementIds().Count,
                ["ElementIdObject"] = set.Id, // Store ElementId for edit system
                ["Id"] = set.Id.AsLong() // Also store as long for compatibility
            });
        }

        // Set UIDocument for automatic edit application
        CustomGUIs.SetCurrentUIDocument(uidoc);

        // Define columns to display
        var columns = new List<string> { "Name", "Count" };

        // Show DataGrid with delete callback
        var selected = CustomGUIs.DataGrid(
            entries: gridData,
            propertyNames: columns,
            spanAllScreens: false,
            onDeleteEntries: (entriesToDelete) =>
            {
                // Delete selection sets in a transaction
                try
                {
                    using (var trans = new Transaction(doc, "Delete Selection Sets"))
                    {
                        trans.Start();

                        foreach (var entry in entriesToDelete)
                        {
                            // Get the ElementId from the entry
                            if (entry.ContainsKey("ElementIdObject") &&
                                entry["ElementIdObject"] is ElementId elemId)
                            {
                                doc.Delete(elemId);
                            }
                            else if (entry.ContainsKey("Id"))
                            {
                                long id = Convert.ToInt64(entry["Id"]);
                                doc.Delete(id.ToElementId());
                            }
                        }

                        trans.Commit();
                    }
                    return true; // Deletion successful
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Delete Error", $"Failed to delete selection sets: {ex.Message}");
                    return false; // Deletion failed
                }
            }
        );

        return Result.Succeeded;
    }
}

#endif
