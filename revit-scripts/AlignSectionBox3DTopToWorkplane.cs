using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AlignSectionBox3DTopToWorkplane : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            
            try
            {
                // Check if current view is a 3D view
                View3D view3D = doc.ActiveView as View3D;
                if (view3D == null)
                {
                    TaskDialog.Show("Error", "Please run this command in a 3D view.");
                    return Result.Failed;
                }
                
                // Check if the view is a perspective view
                if (view3D.IsPerspective)
                {
                    TaskDialog.Show("Error", "Section boxes cannot be used in perspective views. Please use an orthographic 3D view.");
                    return Result.Failed;
                }
                
                // Get the active work plane
                var activeWorkPlane = doc.ActiveView.SketchPlane;
                if (activeWorkPlane == null)
                {
                    TaskDialog.Show("Error", "No active work plane found. Please set a work plane first.");
                    return Result.Failed;
                }
                
                // Get the plane from the work plane
                Plane plane = activeWorkPlane.GetPlane();
                
                // Calculate the Z elevation of the work plane
                // The plane origin gives us a point on the plane
                double workplaneZ = plane.Origin.Z;
                
                // If the plane is not horizontal, we need to consider the plane's normal
                // and find the appropriate Z coordinate for the section box alignment
                if (Math.Abs(plane.Normal.Z) < 0.99999)
                {
                    // Non-horizontal plane
                    TaskDialog dlg = new TaskDialog("Non-Horizontal Work Plane");
                    dlg.MainInstruction = "The current work plane is not horizontal.";
                    dlg.MainContent = "The section box top will be aligned to the Z elevation of the work plane origin.";
                    dlg.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel;
                    
                    if (dlg.Show() == TaskDialogResult.Cancel)
                    {
                        return Result.Cancelled;
                    }
                }
                
                // Enable section box if not already enabled
                if (!view3D.IsSectionBoxActive)
                {
                    using (Transaction trans = new Transaction(doc, "Enable Section Box"))
                    {
                        trans.Start();
                        view3D.IsSectionBoxActive = true;
                        trans.Commit();
                    }
                }
                
                // Get current section box
                BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
                
                // Create new section box with aligned top
                BoundingBoxXYZ newSectionBox = new BoundingBoxXYZ();
                newSectionBox.Transform = sectionBox.Transform;
                newSectionBox.Min = sectionBox.Min;
                newSectionBox.Max = new XYZ(sectionBox.Max.X, sectionBox.Max.Y, sectionBox.Max.Z);
                
                // Transform the workplaneZ to section box coordinate system
                Transform inverseTransform = sectionBox.Transform.Inverse;
                XYZ worldPoint = new XYZ(0, 0, workplaneZ);
                XYZ localPoint = inverseTransform.OfPoint(worldPoint);
                
                // Set the new maximum Z in local coordinates
                newSectionBox.Max = new XYZ(sectionBox.Max.X, sectionBox.Max.Y, localPoint.Z);
                
                // Ensure the box has valid dimensions
                if (newSectionBox.Max.Z <= newSectionBox.Min.Z)
                {
                    TaskDialog.Show("Error", "The work plane is below the bottom of the section box.");
                    return Result.Failed;
                }
                
                // Apply the new section box
                using (Transaction trans = new Transaction(doc, "Align Section Box Top to Work Plane"))
                {
                    trans.Start();
                    view3D.SetSectionBox(newSectionBox);
                    trans.Commit();
                }
                
                // Get work plane name for user feedback
                string workplaneName = "Unknown";
                if (activeWorkPlane.Name != null)
                {
                    workplaneName = activeWorkPlane.Name;
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
