using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using System.Collections.Generic;

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

        // If we're on a sheet, get the viewport’s view.
        if (currentView.ViewType == ViewType.DrawingSheet)
        {
            var selectedIds = uiDoc.Selection.GetElementIds();
            var viewports = selectedIds.Select(id => doc.GetElement(id))
                                       .OfType<Viewport>()
                                       .ToList();
            if (viewports.Count != 1)
            {
                TaskDialog.Show("Error", "Please select a single viewport on the sheet.");
                return Result.Failed;
            }
            targetView = doc.GetElement(viewports.First().ViewId) as View;
        }
        // Otherwise, if the active view has a crop box, use it.
        else if (currentView.CropBoxActive)
        {
            targetView = currentView;
        }
        else
        {
            TaskDialog.Show("Error", "Active view does not have an active crop region.");
            return Result.Failed;
        }

        if (!targetView.CropBoxActive)
        {
            TaskDialog.Show("Error", "The target view does not have an active crop region.");
            return Result.Failed;
        }

        // Get the crop box from the target view.
        BoundingBoxXYZ cropBox = targetView.CropBox;

        // IMPORTANT:
        // The cropBox.Min and cropBox.Max are expressed in the crop box’s local coordinate system.
        // To get the actual (world space) coordinates of its corners, you need to apply cropBox.Transform.
        List<XYZ> worldPoints = new List<XYZ>();
        foreach (double x in new double[] { cropBox.Min.X, cropBox.Max.X })
        {
            foreach (double y in new double[] { cropBox.Min.Y, cropBox.Max.Y })
            {
                foreach (double z in new double[] { cropBox.Min.Z, cropBox.Max.Z })
                {
                    XYZ localPt = new XYZ(x, y, z);
                    XYZ worldPt = cropBox.Transform.OfPoint(localPt);
                    worldPoints.Add(worldPt);
                }
            }
        }

        // Compute the axis-aligned bounding box in world coordinates.
        double minX = worldPoints.Min(pt => pt.X);
        double minY = worldPoints.Min(pt => pt.Y);
        double minZ = worldPoints.Min(pt => pt.Z);
        double maxX = worldPoints.Max(pt => pt.X);
        double maxY = worldPoints.Max(pt => pt.Y);
        double maxZ = worldPoints.Max(pt => pt.Z);

        BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
        sectionBox.Min = new XYZ(minX, minY, minZ);
        sectionBox.Max = new XYZ(maxX, maxY, maxZ);
        // We now have a bounding box that is axis aligned in world coordinates.
        sectionBox.Transform = Transform.Identity;

        // Locate the default 3D view.
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

        // Set the section box on the default 3D view.
        using (Transaction trans = new Transaction(doc, "Set 3D Section Box"))
        {
            trans.Start();
            default3DView.IsSectionBoxActive = true;
            default3DView.SetSectionBox(sectionBox);
            trans.Commit();
        }

        // Activate the default 3D view.
        uiDoc.ActiveView = default3DView;
        return Result.Succeeded;
    }
}
