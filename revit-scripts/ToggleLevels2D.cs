using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class ToggleLevels2D : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message,ElementSet elements)
    {
        // Get the active Revit application and document
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get currently selected levels
        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
        List<Level> selectedLevels = new List<Level>();
        foreach (var id in selectedIds)
        {
            Element element = doc.GetElement(id);
            if (element is Level level)
            {
                selectedLevels.Add(level);
            }
        }

        // Toggle DatumExtentType for selected levels
        using (Transaction transaction = new Transaction(doc, "Toggle DatumExtentType"))
        {
            transaction.Start();

            foreach (Level level in selectedLevels)
            {
                level.SetDatumExtentType(DatumEnds.End0, doc.ActiveView, DatumExtentType.ViewSpecific);
                level.SetDatumExtentType(DatumEnds.End1, doc.ActiveView, DatumExtentType.ViewSpecific);
            }

            transaction.Commit();
        }

        return Result.Succeeded;
    }
}
