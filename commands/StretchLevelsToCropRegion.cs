#region Namespaces
using System;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
#endregion

[Transaction(TransactionMode.Manual)]
public class StretchLevelsToCropRegionCommand : IExternalCommand
{
   public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
   {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;
      View activeView = doc.ActiveView;
      
      // Process only 2D views.
      if (activeView.ViewType == ViewType.ThreeD)
      {
         TaskDialog.Show("Info", "This command only applies to 2D views. The active view is 3D.");
         return Autodesk.Revit.UI.Result.Cancelled;
      }
      
      // Ensure an active crop region exists.
      if (!activeView.CropBoxActive)
      {
         TaskDialog.Show("Info", "The active view does not have an active crop region.");
         return Autodesk.Revit.UI.Result.Cancelled;
      }
      
      // Get the crop region's X boundaries.
      BoundingBoxXYZ cropBox = activeView.CropBox;
      double newX0 = cropBox.Min.X;
      double newX1 = cropBox.Max.X;
      
      using (Transaction trans = new Transaction(doc, "Stretch Datum X Extents"))
      {
         trans.Start();
         
         // Collect DatumPlane elements from the active view.
         var datumElements = new FilteredElementCollector(doc, activeView.Id)
            .OfClass(typeof(DatumPlane))
            .WhereElementIsNotElementType()
            .Cast<DatumPlane>()
            .ToList();
         
         if (!datumElements.Any())
         {
            TaskDialog.Show("Info", "No DatumPlane elements found in the active view.");
            return Autodesk.Revit.UI.Result.Cancelled;
         }
         
         foreach (DatumPlane level in datumElements)
         {
            // Retrieve the current datum curve for the X-axis.
            var curves = level.GetCurvesInView(DatumExtentType.Model, activeView);
            if (!curves.Any())
               continue;
            
            Curve originalCurve = curves.First();
            Line originalLine = originalCurve as Line;
            if (originalLine == null)
               continue;
            
            // Get a base point and direction from the original line.
            XYZ P0 = originalLine.GetEndPoint(0);
            XYZ D = originalLine.Direction;
            
            // Ensure the line's X component is significant.
            if (Math.Abs(D.X) < 1e-6)
            {
               TaskDialog.Show("Warning", "A datum plane with a nearly vertical orientation was skipped.");
               continue;
            }
            
            // Compute parameters so that:
            //   P0 + t0 * D has X = newX0, and
            //   P0 + t1 * D has X = newX1.
            double t0 = (newX0 - P0.X) / D.X;
            double t1 = (newX1 - P0.X) / D.X;
            
            // Evaluate the new endpoints along the original infinite line.
            XYZ newP0 = P0 + t0 * D;
            XYZ newP1 = P0 + t1 * D;
            
            // Create a new bound line using these endpoints.
            Line newCurve = Line.CreateBound(newP0, newP1);
            
            // Update the datum's curve.
            try
            {
               level.SetCurveInView(DatumExtentType.Model, activeView, newCurve);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException ex)
            {
               TaskDialog.Show("Error", $"Failed to update datum: {ex.Message}");
               continue;
            }
         }
         
         trans.Commit();
      }
      
      return Autodesk.Revit.UI.Result.Succeeded;
   }
}
