using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

/// <summary>
/// Shows a datagrid with all levels in the current document.
/// Adds to selection all elements in the active document that are associated with the selected levels.
/// </summary>
[Transaction(TransactionMode.ReadOnly)]
public class SelectByAssociatedLevelsInDocument : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Collect all Level elements in the document, ordered by elevation
        List<Level> levels = new FilteredElementCollector(doc)
            .OfClass(typeof(Level))
            .Cast<Level>()
            .OrderBy(l => l.Elevation)
            .ToList();

        if (levels.Count == 0)
        {
            message = "No levels found in the document.";
            return Result.Failed;
        }

        // Prepare level data for DataGrid
        List<Dictionary<string, object>> levelEntries = new List<Dictionary<string, object>>();
        List<string> propertyNames = new List<string> { "Id", "Name", "Elevation (mm)" };

        foreach (Level lvl in levels)
        {
            Dictionary<string, object> entry = new Dictionary<string, object>
            {
                ["Id"] = lvl.Id.ToString(),
                ["Name"] = lvl.Name,
                // Convert elevation from feet to millimeters
                ["Elevation (mm)"] = (int)Math.Round(lvl.Elevation * 304.8, MidpointRounding.AwayFromZero)
            };

            levelEntries.Add(entry);
        }

        // Display the data grid and get selected levels
        List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(levelEntries, propertyNames, false);

        // If no levels selected, return
        if (selectedEntries == null || selectedEntries.Count == 0)
        {
            return Result.Succeeded;
        }

        // Extract selected level IDs
        HashSet<ElementId> selectedLevelIds = new HashSet<ElementId>();
        foreach (var entry in selectedEntries)
        {
            if (entry.ContainsKey("Id") && entry["Id"] != null)
            {
                if (int.TryParse(entry["Id"].ToString(), out int idInt))
                {
                    selectedLevelIds.Add(idInt.ToElementId());
                }
            }
        }

        if (selectedLevelIds.Count == 0)
        {
            message = "No valid level IDs found in selection.";
            return Result.Failed;
        }

        // Find all elements in the document associated with the selected levels
        List<ElementId> associatedElementIds = new List<ElementId>();

        // Collect all elements that are not element types
        FilteredElementCollector collector = new FilteredElementCollector(doc)
            .WhereElementIsNotElementType();

        foreach (Element elem in collector)
        {
            // Skip the levels themselves
            if (elem is Level)
                continue;

            // Check if element has a level parameter
            Parameter levelParam = elem.get_Parameter(BuiltInParameter.LEVEL_PARAM);
            if (levelParam != null && !levelParam.IsReadOnly)
            {
                ElementId elemLevelId = levelParam.AsElementId();
                if (elemLevelId != null && selectedLevelIds.Contains(elemLevelId))
                {
                    associatedElementIds.Add(elem.Id);
                    continue; // Already added, skip other checks
                }
            }

            // For walls, also check base constraint and top constraint
            if (elem is Wall wall)
            {
                Parameter baseConstraintParam = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
                if (baseConstraintParam != null)
                {
                    ElementId baseId = baseConstraintParam.AsElementId();
                    if (baseId != null && selectedLevelIds.Contains(baseId))
                    {
                        associatedElementIds.Add(elem.Id);
                        continue; // Already added, skip top constraint check
                    }
                }

                Parameter topConstraintParam = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                if (topConstraintParam != null)
                {
                    ElementId topId = topConstraintParam.AsElementId();
                    if (topId != null && selectedLevelIds.Contains(topId))
                    {
                        associatedElementIds.Add(elem.Id);
                    }
                }
            }

            // For rooms, check their level
            if (elem is Room room)
            {
                if (room.Level != null && selectedLevelIds.Contains(room.Level.Id))
                {
                    associatedElementIds.Add(elem.Id);
                }
            }

            // For views, check associated level (plan views, ceiling plans, etc.)
            if (elem is View view && view.GenLevel != null)
            {
                if (selectedLevelIds.Contains(view.GenLevel.Id))
                {
                    associatedElementIds.Add(elem.Id);
                }
            }
        }

        // Remove duplicates (e.g., walls might be added multiple times)
        associatedElementIds = associatedElementIds.Distinct().ToList();

        if (associatedElementIds.Count == 0)
        {
            return Result.Succeeded;
        }

        // Add the associated elements to the current selection
        uidoc.AddToSelection(associatedElementIds);

        return Result.Succeeded;
    }
}
