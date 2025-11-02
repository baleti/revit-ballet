using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectLinkedModels : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get application and document objects
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get a list of linked models
        List<RevitLinkInstance> linkedInstances = new FilteredElementCollector(doc)
            .OfClass(typeof(RevitLinkInstance))
            .Cast<RevitLinkInstance>()
            .ToList();

        var propertyNames = new List<string> { "Name" };
        //Create a new DataGridView
        var selectedCategories = CustomGUIs.DataGrid<RevitLinkInstance>(linkedInstances, propertyNames);

        // Select the linked models
        List<ElementId> linkedModelIds = selectedCategories.Select(x => x.Id).ToList();

        // Get current selection
        ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();

        // Add new selections to the existing selection
        foreach (var id in linkedModelIds)
        {
            if (!currentSelection.Contains(id))
            {
                currentSelection.Add(id);
            }
        }

        // Set the updated selection
        uidoc.SetSelectionIds(currentSelection);

        return Result.Succeeded;
    }
}
