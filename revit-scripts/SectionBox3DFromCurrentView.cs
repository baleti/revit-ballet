using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SectionBox3DFromCurrentView : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData, 
      ref string message, 
      ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;
        View currentView = doc.ActiveView;
        View targetView = null;

        // Check if the current view is a sheet and if viewports are selected
        if (currentView.ViewType == ViewType.DrawingSheet)
        {
            var selectedIds = uiDoc.Selection.GetElementIds();
            var selectedViewports = selectedIds.Select(id => doc.GetElement(id)).OfType<Viewport>().ToList();

            if (selectedViewports.Count == 0)
            {
                TaskDialog.Show("Error", "No viewport selected.");
                return Result.Failed;
            }
            else if (selectedViewports.Count > 1)
            {
                TaskDialog.Show("Warning", "Multiple viewports selected. Please select a single viewport.");
                return Result.Failed;
            }

            targetView = doc.GetElement(selectedViewports.First().ViewId) as View;
        }
        else if (currentView.ViewType == ViewType.FloorPlan ||
                 currentView.ViewType == ViewType.CeilingPlan ||
                 currentView.ViewType == ViewType.Section ||
                 currentView.ViewType == ViewType.Elevation)
        {
            targetView = currentView;
        }

        if (targetView == null)
        {
            TaskDialog.Show("Error", "No suitable view or viewport selected.");
            return Result.Failed;
        }

        // Check if the target view has a crop region
        if (!targetView.CropBoxActive)
        {
            TaskDialog.Show("Error", "The selected view does not have an active crop region.");
            return Result.Failed;
        }

        // Get the crop region of the target view
        BoundingBoxXYZ cropBox = targetView.CropBox;

        // Find the default 3D view named in the format "{3D - username}"
        string default3DViewName = "{3D - " + uiApp.Application.Username + "}";
        View3D default3DView = new FilteredElementCollector(doc)
                                .OfClass(typeof(View3D))
                                .Cast<View3D>()
                                .FirstOrDefault(v => v.Name == default3DViewName);

        if (default3DView == null)
        {
            TaskDialog.Show("Error", "The default 3D view could not be found.");
            return Result.Failed;
        }

        using (Transaction trans = new Transaction(doc, "Set 3D Section Box"))
        {
            trans.Start();
            default3DView.IsSectionBoxActive = true;
            default3DView.SetSectionBox(cropBox);
            trans.Commit();
        }

        // Activate the default 3D view
        uiDoc.ActiveView = default3DView;

        return Result.Succeeded;
    }
}
