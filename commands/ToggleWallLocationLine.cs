using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

[Transaction(TransactionMode.Manual)]
[CommandMeta("Wall")]
public class ToggleWallsLocationLineFromFinishFaceExteriorToFinishFaceInterior : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        // Get the current document and selection
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        Selection sel = uidoc.Selection;

        // Get the selected elements
        var selectedElementIds = uidoc.GetSelectionIds();

        // Filter the selected elements to get only walls
        var selectedWalls = selectedElementIds
            .Select(id => doc.GetElement(id))
            .OfType<Wall>()
            .ToList();

        if (selectedWalls.Count == 0)
        {
            CustomGUIs.SetCurrentUIDocument(uidoc);
            var allWalls = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
                .Cast<Wall>().ToList();
            var gridData = CustomGUIs.ConvertToDataGridFormat(allWalls, new List<string> { "Name" });
            var chosen = CustomGUIs.DataGrid(gridData, new List<string> { "Name" }, false);
            if (chosen == null) return Result.Cancelled;
            selectedWalls = CustomGUIs.ExtractOriginalObjects<Wall>(chosen) ?? new List<Wall>();
            if (selectedWalls.Count == 0) return Result.Succeeded;
        }

        try
        {
            // Start a transaction
            using (Transaction trans = new Transaction(doc, "Toggle Wall Location Line"))
            {
                trans.Start();

                foreach (var wall in selectedWalls)
                {
                    // Get the current location line of the wall
                    WallLocationLine currentLocationLine = (WallLocationLine)wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).AsInteger();

                    // Toggle between "Finish Face: Exterior" and "Finish Face: Interior"
                    if (currentLocationLine == WallLocationLine.FinishFaceExterior)
                    {
                        wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).Set((int)WallLocationLine.FinishFaceInterior);
                    }
                    else if (currentLocationLine == WallLocationLine.FinishFaceInterior)
                    {
                        wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).Set((int)WallLocationLine.FinishFaceExterior);
                    }
                }

                // Commit the transaction
                trans.Commit();
            }

            // Re-select the walls (this ensures the same elements remain selected after the operation)
            uidoc.SetSelectionIds(selectedWalls.Select(w => w.Id).ToList());

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            // Handle any other exceptions
            message = ex.Message;
            return Result.Failed;
        }
    }
}
