using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectAllInViewsAndInGroups : IExternalCommand
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

        // Collect all elements from all target views
        HashSet<ElementId> allElements = new HashSet<ElementId>();

        foreach (View view in targetViews)
        {
            var viewElements = new FilteredElementCollector(doc, view.Id)
                .ToElementIds();

            foreach (var id in viewElements)
            {
                allElements.Add(id);
            }
        }

        // Set the selection to these filtered elements
        uiDoc.SetSelectionIds(allElements.ToList());

        return Result.Succeeded;
    }
}
