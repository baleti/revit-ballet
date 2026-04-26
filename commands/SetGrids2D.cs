using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
[CommandMeta("Grid")]
public class SetGrids2D : IExternalCommand
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
                selectedGrids.Add(grid);
        }

        if (selectedGrids.Count == 0)
        {
            CustomGUIs.SetCurrentUIDocument(uiDoc);
            var allGrids = new FilteredElementCollector(doc).OfClass(typeof(Grid)).Cast<Grid>().ToList();
            var gridData = CustomGUIs.ConvertToDataGridFormat(allGrids, new List<string> { "Name" });
            var chosen = CustomGUIs.DataGrid(gridData, new List<string> { "Name" }, false);
            if (chosen == null) return Result.Cancelled;
            selectedGrids = CustomGUIs.ExtractOriginalObjects<Grid>(chosen) ?? new List<Grid>();
            if (selectedGrids.Count == 0) return Result.Succeeded;
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
