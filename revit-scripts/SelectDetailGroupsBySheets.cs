using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectElementsWithSameCommentInCurrentView : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData, 
      ref string message, 
      ElementSet elements)
    {
        // Get the current document and UIDocument
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the currently selected element
        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
        if (selectedIds.Count != 1)
        {
            message = "Please select a single element.";
            return Result.Failed;
        }

        ElementId selectedId = selectedIds.First();
        Element selectedElement = doc.GetElement(selectedId);

        // Get the comment value of the selected element
        string commentValue = selectedElement.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.AsString();
        if (string.IsNullOrEmpty(commentValue))
        {
            message = "The selected element does not have a comment value.";
            return Result.Failed;
        }

        // Get the current view
        View currentView = doc.ActiveView;

        // Find all elements in the current view with the same comment value
        FilteredElementCollector collector = new FilteredElementCollector(doc, currentView.Id);
        List<ElementId> matchingElementIds = collector
            .WherePasses(new ElementParameterFilter(ParameterFilterRuleFactory.CreateEqualsRule(
                new ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS),
                commentValue, false)))
            .Select(e => e.Id)
            .ToList();

        // Select those elements
        uiDoc.Selection.SetElementIds(matchingElementIds);

        return Result.Succeeded;
    }
}
