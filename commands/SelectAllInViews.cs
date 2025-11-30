using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectAllInViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        Document doc = commandData.Application.ActiveUIDocument.Document;
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;

        // Determine target views (selected views or active view)
        ICollection<ElementId> currentSelection = uiDoc.GetSelectionIds();
        List<View> targetViews = new List<View>();

        bool hasSelectedViews = currentSelection.Any(id =>
        {
            Element elem = doc.GetElement(id);
            return elem is View;
        });

        if (hasSelectedViews)
        {
            // Use selected views
            targetViews = currentSelection
                .Select(id => doc.GetElement(id))
                .OfType<View>()
                .ToList();
        }
        else
        {
            // Use current view only
            targetViews.Add(doc.ActiveView);
        }

        // Collect elements from all target views
        HashSet<ElementId> elementsNotInGroups = new HashSet<ElementId>();

        foreach (View view in targetViews)
        {
            var viewElements = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.GroupId == ElementId.InvalidElementId)
                .Where(e => !(e is View)) // Exclude views
                .Where(x => x.Category != null) // Exclude ExtentElem
                .Where(e => !(e.Category?.Id.AsLong() == (int)BuiltInCategory.OST_Cameras)) // Exclude cameras
                .Select(e => e.Id);

            foreach (var id in viewElements)
            {
                elementsNotInGroups.Add(id);
            }
        }

        uiDoc.SetSelectionIds(elementsNotInGroups.ToList());
        return Result.Succeeded;
    }
}
