using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;

[Transaction(TransactionMode.ReadOnly)]
public class Print2DViewCropTransformation : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        // Get the current document and active view
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = doc.ActiveView;

        if (activeView == null)
        {
            message = "No active view found.";
            return Result.Failed;
        }

        // Check if the active view is a 2D view like Plan or Section (including Elevation)
        if (activeView is ViewPlan || activeView is ViewSection)
        {
            if (activeView.CropBoxActive && activeView.CropBoxVisible)
            {
                BoundingBoxXYZ cropBox = activeView.CropBox;
                Transform cropTransform = cropBox.Transform;

                // Get the transformation data for the crop region
                XYZ translation = cropTransform.Origin;
                XYZ basisX = cropTransform.BasisX;
                XYZ basisY = cropTransform.BasisY;
                XYZ basisZ = cropTransform.BasisZ;

                TaskDialog.Show("2D View Crop Region Transformation", 
                    $"Crop Region Translation: {FormatXYZ(translation)}\n" +
                    $"Crop Region Basis X: {FormatXYZ(basisX)}\n" +
                    $"Crop Region Basis Y: {FormatXYZ(basisY)}\n" +
                    $"Crop Region Basis Z: {FormatXYZ(basisZ)}");
            }
            else
            {
                TaskDialog.Show("Crop Region Info", "The crop region is not active or visible.");
            }
        }
        else
        {
            TaskDialog.Show("View Info", "This command is designed for 2D views (plans, sections).");
        }

        return Result.Succeeded;
    }

    // Helper method to format XYZ points
    private string FormatXYZ(XYZ point)
    {
        return $"X: {point.X}, Y: {point.Y}, Z: {point.Z}";
    }
}
