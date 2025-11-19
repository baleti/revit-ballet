using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectDetailItemsWithKeynote : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData, 
      ref string message, 
      ElementSet elements)
    {
        // Get the current document
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Create a filter to select all detail items in the project
        ElementCategoryFilter detailItemFilter = new ElementCategoryFilter(BuiltInCategory.OST_DetailComponents);

        // Create a collector to get all detail items
        FilteredElementCollector detailItemsCollector = new FilteredElementCollector(doc)
            .WherePasses(detailItemFilter)
            .WhereElementIsNotElementType();

        // List to hold the elements that match the keynote criteria
        var detailItemsWithKeynote = detailItemsCollector
            .Where(e => e.LookupParameter("Keynote") != null)
            .Where(e => e.LookupParameter("Keynote").AsString().StartsWith("P10"))
            .ToList();

        // If no elements found, display a message and return
        if (!detailItemsWithKeynote.Any())
        {
            TaskDialog.Show("No Elements Found", "There are no detail items with a keynote that starts with 'P10'.");
            return Result.Succeeded;
        }

        // Select the elements in the Revit UI
        uidoc.SetSelectionIds(detailItemsWithKeynote.Select(e => e.Id).ToList());

        // Return success
        return Result.Succeeded;
    }
}
