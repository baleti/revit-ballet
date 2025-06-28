using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectAllElementsInViewAndInGroups : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        Document doc = commandData.Application.ActiveUIDocument.Document;
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;

        var elementIdsInView = new FilteredElementCollector(doc, doc.ActiveView.Id)
            .ToElementIds();

        // Set the selection to these filtered elements
        uiDoc.SetSelectionIds(elementIdsInView);

        return Result.Succeeded;
    }
}
