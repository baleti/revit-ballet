using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectAllElementsInView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        Document doc = commandData.Application.ActiveUIDocument.Document;
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        View currentView = doc.ActiveView;

        List<ElementId> ElementsNotInGroups = new FilteredElementCollector(doc, doc.ActiveView.Id)
            .WhereElementIsNotElementType()
            .Where(e => e.GroupId == ElementId.InvalidElementId)
            .Select(e => e.Id)
            .ToList();

        uiDoc.Selection.SetElementIds(ElementsNotInGroups);

        return Result.Succeeded;
    }
}
