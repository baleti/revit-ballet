using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class SetSelectedGrids2D : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the active Revit application and document
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get currently selected grids
        ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
        List<Grid> selectedGrids = new List<Grid>();
        foreach (var id in selectedIds)
        {
            Element element = doc.GetElement(id);
            if (element is Grid grid)
            {
                selectedGrids.Add(grid);
            }
        }

        // Toggle DatumExtentType for selected grids
        using (Transaction transaction = new Transaction(doc, "Toggle 2D Extents for Grids"))
        {
            transaction.Start();

            foreach (Grid grid in selectedGrids)
            {
                // Toggle both ends to ViewSpecific (2D)
                grid.SetDatumExtentType(DatumEnds.End0, doc.ActiveView, DatumExtentType.ViewSpecific);
                grid.SetDatumExtentType(DatumEnds.End1, doc.ActiveView, DatumExtentType.ViewSpecific);
            }

            transaction.Commit();
        }

        return Result.Succeeded;
    }
}
