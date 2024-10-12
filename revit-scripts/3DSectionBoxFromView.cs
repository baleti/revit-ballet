using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SectionBox3DFromViewList : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData, 
      ref string message, 
      ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Check if the current view is a 3D view
        View3D current3DView = doc.ActiveView as View3D;
        if (current3DView == null)
        {
            TaskDialog.Show("Error", "The current view is not a 3D view.");
            return Result.Failed;
        }

        // Get all views in the document
        List<View> views = new FilteredElementCollector(doc)
                            .OfClass(typeof(View))
                            .Cast<View>()
                            .Where(v => !v.IsTemplate)
                            .ToList();

        // Create entries for the custom GUI
        List<Dictionary<string, object>> entries = views.Select(view => new Dictionary<string, object>
        {
            { "Name", view.Name },
            { "Id", view.Id }
        }).ToList();

        // Define property names for the custom GUI
        List<string> propertyNames = new List<string> { "Name", "Id" };

        // Show the custom GUI
        List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, false);

        if (selectedEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No view selected.");
            return Result.Cancelled;
        }

        // Get the selected view
        Dictionary<string, object> selectedEntry = selectedEntries.First();
        ElementId selectedViewId = (ElementId)selectedEntry["Id"];
        View selectedView = doc.GetElement(selectedViewId) as View;

        // Get the crop region of the selected view
        BoundingBoxXYZ cropBox = selectedView.CropBox;

        using (Transaction trans = new Transaction(doc, "Set 3D Section Box"))
        {
            trans.Start();
            current3DView.IsSectionBoxActive = true;
            current3DView.SetSectionBox(cropBox);
            trans.Commit();
        }

        return Result.Succeeded;
    }
}
