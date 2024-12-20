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

        // Collect elements that are not in groups and explicitly exclude the active view and ExtentElem
        List<ElementId> ElementsNotInGroups = new FilteredElementCollector(doc, currentView.Id)
            .WhereElementIsNotElementType()
            .Where(e => e.GroupId == ElementId.InvalidElementId)
            .Where(e => !(e is View)) // Exclude views explicitly
            .Where( x => x.Category != null) // Exclude ExtentElem https://thebuildingcoder.typepad.com/blog/2017/09/extentelem-and-square-face-dimensioning-references.html
            .Select(e => e.Id)
            .ToList();

        // Set the selection to the filtered element IDs
        uiDoc.Selection.SetElementIds(ElementsNotInGroups);

        return Result.Succeeded;
    }
}
