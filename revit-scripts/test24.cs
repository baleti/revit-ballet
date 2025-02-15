using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

namespace MyRevitCommands
{
  [Transaction(TransactionMode.Manual)]
  public class CmdCreateSectionViews : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      // Get the active UIDocument and Document
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;
      View activeView = doc.ActiveView;

      // Get currently selected elements
      ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
      if (selIds.Count == 0)
      {
        message = "Please select at least one element.";
        return Result.Failed;
      }

      // Get a section view family type
      ViewFamilyType sectionType = new FilteredElementCollector(doc)
        .OfClass(typeof(ViewFamilyType))
        .Cast<ViewFamilyType>()
        .FirstOrDefault(x => x.ViewFamily == ViewFamily.Section);
      if (sectionType == null)
      {
        message = "No section view type found.";
        return Result.Failed;
      }

      using (Transaction tx = new Transaction(doc, "Create Section Views"))
      {
        tx.Start();

        foreach (ElementId id in selIds)
        {
          Element elem = doc.GetElement(id);
          BoundingBoxXYZ bbox = elem.get_BoundingBox(activeView);
          if (bbox == null)
            continue;

          // Compute the element's center in world coordinates.
          XYZ localCenter = (bbox.Min + bbox.Max) * 0.5;
          XYZ center = bbox.Transform.OfPoint(localCenter);

          // Get the element’s rotation (if available, from a LocationPoint).
          double elementRotation = 0.0;
          if (elem.Location is LocationPoint locPoint)
          {
            elementRotation = locPoint.Rotation;
          }

          // --- Determine the coordinate system for the section view ---
          // If the active view is a plan view, override the directions to create a vertical section.
          XYZ sectionUp;
          XYZ sectionViewDir;
          if (activeView is ViewPlan)
          {
            // For a vertical section view, the up vector is vertical.
            sectionUp = XYZ.BasisZ;
            // Create a horizontal view direction based on the element's rotation.
            sectionViewDir = new XYZ(Math.Cos(elementRotation), Math.Sin(elementRotation), 0);
          }
          else
          {
            // Otherwise, use the active view's directions.
            sectionUp = activeView.UpDirection.Normalize();
            sectionViewDir = activeView.ViewDirection.Normalize();
          }

          // Compute the right direction (X axis) for the section view.
          XYZ sectionRight = sectionUp.CrossProduct(sectionViewDir).Normalize();

          // Build the transform for the section view coordinate system:
          // - Origin: element's center
          // - BasisX: sectionRight, BasisY: sectionUp, BasisZ: sectionViewDir
          Transform sectionTransform = Transform.Identity;
          sectionTransform.Origin = center;
          sectionTransform.BasisX = sectionRight;
          sectionTransform.BasisY = sectionUp;
          sectionTransform.BasisZ = sectionViewDir;

          // Compute the 8 corner points of the element's bounding box in world coordinates.
          List<XYZ> worldCorners = new List<XYZ>();
          foreach (int ix in new int[] { 0, 1 })
          {
            double x = (ix == 0) ? bbox.Min.X : bbox.Max.X;
            foreach (int iy in new int[] { 0, 1 })
            {
              double y = (iy == 0) ? bbox.Min.Y : bbox.Max.Y;
              foreach (int iz in new int[] { 0, 1 })
              {
                double z = (iz == 0) ? bbox.Min.Z : bbox.Max.Z;
                XYZ localPt = new XYZ(x, y, z);
                XYZ worldPt = bbox.Transform.OfPoint(localPt);
                worldCorners.Add(worldPt);
              }
            }
          }

          // Transform these world points into the section view coordinate system.
          List<XYZ> localPoints = worldCorners.Select(pt => sectionTransform.Inverse.OfPoint(pt)).ToList();

          // Determine the extents in the section view coordinate system.
          double minX = localPoints.Min(pt => pt.X);
          double maxX = localPoints.Max(pt => pt.X);
          double minY = localPoints.Min(pt => pt.Y);
          double maxY = localPoints.Max(pt => pt.Y);

          // Optionally, add a margin (in the section view's units)
          double margin = 1.0;
          minX -= margin;
          maxX += margin;
          minY -= margin;
          maxY += margin;

          // For the Z direction, use a thin slice.
          double minZ = -0.5;
          double maxZ = 0.5;

          // Create the BoundingBoxXYZ for the section view.
          BoundingBoxXYZ sectionBB = new BoundingBoxXYZ();
          sectionBB.Transform = sectionTransform;
          sectionBB.Min = new XYZ(minX, minY, minZ);
          sectionBB.Max = new XYZ(maxX, maxY, maxZ);

          // Create the section view using the BoundingBoxXYZ.
          ViewSection sectionView = ViewSection.CreateSection(doc, sectionType.Id, sectionBB);

          // Ensure a unique name by appending the element Id.
          sectionView.Name = $"Section of {elem.Name} ({elem.Id.IntegerValue})";
        }

        tx.Commit();
      }

      return Result.Succeeded;
    }
  }
}
