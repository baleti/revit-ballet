using System;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

[Transaction(TransactionMode.Manual)]
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
        var selectedElementIds = sel.GetElementIds();

        // Filter the selected elements to get only walls
        var selectedWalls = selectedElementIds
            .Select(id => doc.GetElement(id))
            .OfType<Wall>()
            .ToList();

        // If no walls are selected, do nothing
        if (selectedWalls.Count == 0)
        {
            message = "No walls are selected.";
            return Result.Cancelled;
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
            sel.SetElementIds(selectedWalls.Select(w => w.Id).ToList());

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
