using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

// Alias WinForms and Drawing types.
using WinForms = System.Windows.Forms;
using DrawingPoint = System.Drawing.Point;
using DrawingPointF = System.Drawing.PointF;
using DrawingSize = System.Drawing.Size;

// Using ClipperLib.
using ClipperLib;
using ClipType = ClipperLib.ClipType;

// Alias Revit's View to avoid ambiguity.
using RevitView = Autodesk.Revit.DB.View;

namespace RevitAddin
{
  [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
  public class BooleanDifferenceFilledRegionsWithSelectedElements : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      // Prompt the user for an offset in millimeters.
      double? offsetFeet = PromptForOffset();
      if (!offsetFeet.HasValue)
        return Result.Cancelled;
      
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      // Partition selection: filled regions versus subtracting elements.
      List<FilledRegion> selectedRegions = new List<FilledRegion>();
      List<ElementId> subtractElementIds = new List<ElementId>();

      ICollection<ElementId> selIds = uidoc.GetSelectionIds();
      if (selIds == null || selIds.Count == 0)
      {
        WinForms.MessageBox.Show("Please select at least one filled region and one subtracting element.", "No Selection");
        return Result.Failed;
      }
      foreach (ElementId id in selIds)
      {
        Element elem = doc.GetElement(id);
        if (elem is FilledRegion)
          selectedRegions.Add(elem as FilledRegion);
        else
          subtractElementIds.Add(id);
      }
      if (selectedRegions.Count == 0)
      {
        WinForms.MessageBox.Show("Please select at least one filled region.", "No Filled Regions");
        return Result.Failed;
      }
      if (subtractElementIds.Count == 0)
      {
        WinForms.MessageBox.Show("Please select at least one subtracting element (non-filled region).", "No Subtracting Elements");
        return Result.Failed;
      }

      // Group filled regions by their view (using OwnerViewId).
      Dictionary<ElementId, List<FilledRegion>> regionsByView = new Dictionary<ElementId, List<FilledRegion>>();
      foreach (FilledRegion fr in selectedRegions)
      {
        ElementId viewId = fr.OwnerViewId;
        if (!regionsByView.ContainsKey(viewId))
          regionsByView[viewId] = new List<FilledRegion>();
        regionsByView[viewId].Add(fr);
      }

      // Define a scale factor for Clipper (working in integer coordinates).
      double clipperScale = 1e6;

      using (Transaction trans = new Transaction(doc, "Boolean Difference Filled Regions"))
      {
        trans.Start();

        // Process each view separately.
        foreach (var kvp in regionsByView)
        {
          ElementId viewId = kvp.Key;
          RevitView view = doc.GetElement(viewId) as RevitView;
          if (view == null)
            continue;

          // 1. Compute union outline from subtracting elements in this view.
          List<List<IntPoint>> subtractPolygons = new List<List<IntPoint>>();
          foreach (ElementId subId in subtractElementIds)
          {
            Element subElem = doc.GetElement(subId);
            BoundingBoxXYZ bbox = subElem.get_BoundingBox(view) ?? subElem.get_BoundingBox(null);
            if (bbox == null)
              continue;

            // Get all eight corners.
            List<XYZ> corners = new List<XYZ>
            {
              new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
              new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
              new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
              new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
              new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
              new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
              new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
              new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z)
            };

            List<DrawingPointF> localPoints = new List<DrawingPointF>();
            foreach (XYZ p in corners)
            {
              XYZ projected = ProjectPointToViewPlane(p, view);
              DrawingPointF localPt = WorldPointToViewLocal(projected, view);
              localPoints.Add(localPt);
            }
            List<DrawingPointF> hull = Compute2DConvexHull(localPoints);
            if (hull.Count < 3)
              continue;
            List<IntPoint> clipperPoly = new List<IntPoint>();
            foreach (DrawingPointF pt in hull)
            {
              clipperPoly.Add(new IntPoint((long)(pt.X * clipperScale), (long)(pt.Y * clipperScale)));
            }
            subtractPolygons.Add(clipperPoly);
          }

          // Union all subtracting polygons.
          List<List<IntPoint>> subtractUnion = new List<List<IntPoint>>();
          if (subtractPolygons.Count > 0)
          {
            Clipper clipperUnion = new Clipper();
            foreach (List<IntPoint> poly in subtractPolygons)
            {
              clipperUnion.AddPolygon(poly, PolyType.ptSubject);
            }
            clipperUnion.Execute(ClipType.ctUnion, subtractUnion, PolyFillType.pftNonZero, PolyFillType.pftNonZero);
          }
          else
          {
            // No subtract geometry in this view; skip.
            continue;
          }

          // Apply offset using our parallel offset algorithm.
          if (Math.Abs(offsetFeet.Value) > 1e-9)
          {
            double offsetClipper = offsetFeet.Value * clipperScale;
            List<List<IntPoint>> offsetSubUnion = new List<List<IntPoint>>();
            foreach (var poly in subtractUnion)
            {
              offsetSubUnion.Add(OffsetPolygonParallelInt(poly, offsetClipper));
            }
            subtractUnion = offsetSubUnion;
          }

          // 2. Process each filled region in this view.
          foreach (FilledRegion fr in kvp.Value)
          {
            // Extract all boundary curve loops from the filled region.
            List<CurveLoop> frLoops = GetFilledRegionCurveLoops(fr, view);
            if (frLoops == null || frLoops.Count == 0)
            {
              WinForms.MessageBox.Show($"Could not extract boundary for filled region in view \"{view.Name}\".", "Info");
              continue;
            }

            // Convert each CurveLoop into a Clipper polygon.
            List<List<IntPoint>> frPolygons = new List<List<IntPoint>>();
            foreach (CurveLoop cl in frLoops)
            {
              List<IntPoint> poly = new List<IntPoint>();
              List<XYZ> pts = new List<XYZ>();
              foreach (Curve curve in cl)
              {
                pts.Add(curve.GetEndPoint(0));
              }
              List<DrawingPointF> localPts = pts.Select(p => WorldPointToViewLocal(p, view)).ToList();
              foreach (DrawingPointF pt in localPts)
              {
                poly.Add(new IntPoint((long)(pt.X * clipperScale), (long)(pt.Y * clipperScale)));
              }
              frPolygons.Add(poly);
            }

            // Perform boolean difference: subtractUnion is subtracted from the filled-region polygons.
            List<List<IntPoint>> resultPolys = new List<List<IntPoint>>();
            Clipper diffClipper = new Clipper();
            foreach (List<IntPoint> subj in frPolygons)
            {
              diffClipper.AddPolygon(subj, PolyType.ptSubject);
            }
            foreach (List<IntPoint> clipPoly in subtractUnion)
            {
              diffClipper.AddPolygon(clipPoly, PolyType.ptClip);
            }
            diffClipper.Execute(ClipType.ctDifference, resultPolys, PolyFillType.pftNonZero, PolyFillType.pftNonZero);

            if (resultPolys.Count == 0)
            {
              WinForms.MessageBox.Show($"Boolean difference produced no result for a filled region in view \"{view.Name}\".", "Info");
              continue;
            }

            // Store the filled region's type ID before deletion.
            ElementId frTypeId = fr.GetTypeId();
            doc.Delete(fr.Id);

            // Combine all resulting curve loops into one new filled region.
            List<CurveLoop> combinedLoops = new List<CurveLoop>();
            foreach (List<IntPoint> resPoly in resultPolys)
            {
              List<XYZ> worldPts = new List<XYZ>();
              foreach (IntPoint ip in resPoly)
              {
                double x = ip.X / clipperScale;
                double y = ip.Y / clipperScale;
                XYZ wp = ViewLocalToWorldPoint(new DrawingPointF((float)x, (float)y), view);
                worldPts.Add(wp);
              }
              if (worldPts.Count < 3)
                continue;
              if (!IsClockwise(worldPts))
                worldPts.Reverse();

              CurveLoop newLoop = new CurveLoop();
              for (int i = 0; i < worldPts.Count; i++)
              {
                XYZ p1 = worldPts[i];
                XYZ p2 = worldPts[(i + 1) % worldPts.Count];
                newLoop.Append(Line.CreateBound(p1, p2));
              }
              // Clean the loop: remove overlapping or collinear edges.
              CurveLoop cleanLoop = CleanCurveLoop(newLoop, 1e-6);
              if (cleanLoop != null)
                combinedLoops.Add(cleanLoop);
            }

            if (combinedLoops.Count > 0)
            {
              FilledRegion newFR = FilledRegion.Create(doc, frTypeId, view.Id, combinedLoops);
            }
          }
        }
        trans.Commit();
      }
      return Result.Succeeded;
    }

    #region Helper Methods

    /// <summary>
    /// Displays a WinForm to prompt the user for an offset (in millimeters).
    /// Returns the offset converted to feet (Revit's internal units), or null if cancelled.
    /// </summary>
    private double? PromptForOffset()
    {
      WinForms.Form form = new WinForms.Form();
      form.Text = "Enter Offset (mm)";
      form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
      form.StartPosition = WinForms.FormStartPosition.CenterScreen;
      form.ClientSize = new DrawingSize(220, 80);
      form.MaximizeBox = false;
      form.MinimizeBox = false;

      WinForms.Label label = new WinForms.Label();
      label.Text = "Offset (mm):";
      label.Location = new DrawingPoint(10, 10);
      label.AutoSize = true;
      form.Controls.Add(label);

      WinForms.TextBox textBox = new WinForms.TextBox();
      textBox.Location = new DrawingPoint(100, 10);
      textBox.Width = 100;
      form.Controls.Add(textBox);

      WinForms.Button okButton = new WinForms.Button();
      okButton.Text = "OK";
      okButton.DialogResult = WinForms.DialogResult.OK;
      okButton.Location = new DrawingPoint(40, 40);
      form.Controls.Add(okButton);

      WinForms.Button cancelButton = new WinForms.Button();
      cancelButton.Text = "Cancel";
      cancelButton.DialogResult = WinForms.DialogResult.Cancel;
      cancelButton.Location = new DrawingPoint(120, 40);
      form.Controls.Add(cancelButton);

      form.AcceptButton = okButton;
      form.CancelButton = cancelButton;

      WinForms.DialogResult dr = form.ShowDialog();
      if (dr == WinForms.DialogResult.OK)
      {
        if (double.TryParse(textBox.Text, out double mmValue))
        {
          // Convert millimeters to feet.
          return mmValue / 304.8;
        }
      }
      return null;
    }

    /// <summary>
    /// Offsets a polygon using a parallel offset algorithm.
    /// For each vertex, the adjacent edges are offset parallelly and their intersection is computed.
    /// This implementation uses the polygon's 2D orientation (via IsClockwise2D) so that a positive offset always expands
    /// the shape and a negative offset contracts it.
    /// Returns the offset polygon as a list of IntPoint (in Clipper units).
    /// </summary>
    private List<IntPoint> OffsetPolygonParallelInt(List<IntPoint> poly, double offset)
    {
      List<DrawingPointF> pts = poly.Select(pt => new DrawingPointF((float)pt.X, (float)pt.Y)).ToList();
      List<DrawingPointF> offsetPts = OffsetPolygonParallel(pts, offset);
      List<IntPoint> newPoly = offsetPts.Select(p => new IntPoint((long)Math.Round(p.X), (long)Math.Round(p.Y))).ToList();
      return newPoly;
    }

    /// <summary>
    /// Offsets a polygon (as a list of DrawingPointF) by a fixed distance using a parallel offset.
    /// Determines the polygon's orientation via IsClockwise2D.
    /// For each vertex, computes the offset intersection of its two adjacent edges.
    /// </summary>
    private List<DrawingPointF> OffsetPolygonParallel(List<DrawingPointF> poly, double offset)
    {
      int n = poly.Count;
      List<DrawingPointF> newPoly = new List<DrawingPointF>();
      bool isCW = IsClockwise2D(poly);
      for (int i = 0; i < n; i++)
      {
        DrawingPointF prev = poly[(i - 1 + n) % n];
        DrawingPointF curr = poly[i];
        DrawingPointF next = poly[(i + 1) % n];

        DrawingPointF d0 = Normalize(Subtract(curr, prev));
        DrawingPointF d1 = Normalize(Subtract(next, curr));

        // Determine outward normal based on orientation.
        DrawingPointF outward0 = isCW ? new DrawingPointF(-d0.Y, d0.X) : new DrawingPointF(d0.Y, -d0.X);
        DrawingPointF outward1 = isCW ? new DrawingPointF(-d1.Y, d1.X) : new DrawingPointF(d1.Y, -d1.X);
        float effectiveOffset = (float)Math.Abs(offset);
        DrawingPointF normal0 = offset >= 0 ? outward0 : Multiply(outward0, -1);
        DrawingPointF normal1 = offset >= 0 ? outward1 : Multiply(outward1, -1);

        DrawingPointF P0 = Add(curr, Multiply(normal0, effectiveOffset));
        DrawingPointF P1 = Add(curr, Multiply(normal1, effectiveOffset));

        bool success;
        DrawingPointF ip = IntersectLines(P0, d0, P1, d1, out success);
        if (!success)
        {
          ip = new DrawingPointF((P0.X + P1.X) / 2, (P0.Y + P1.Y) / 2);
        }
        newPoly.Add(ip);
      }
      return newPoly;
    }

    // Helper vector operations.
    private DrawingPointF Subtract(DrawingPointF a, DrawingPointF b) => new DrawingPointF(a.X - b.X, a.Y - b.Y);
    private DrawingPointF Add(DrawingPointF a, DrawingPointF b) => new DrawingPointF(a.X + b.X, a.Y + b.Y);
    private DrawingPointF Multiply(DrawingPointF a, float scalar) => new DrawingPointF(a.X * scalar, a.Y * scalar);
    private double Length(DrawingPointF a) => Math.Sqrt(a.X * a.X + a.Y * a.Y);
    private DrawingPointF Normalize(DrawingPointF a)
    {
      double len = Length(a);
      return len < 1e-9 ? new DrawingPointF(0, 0) : new DrawingPointF((float)(a.X / len), (float)(a.Y / len));
    }

    /// <summary>
    /// Computes the intersection of two lines in 2D.
    /// Line1: P0 + t*d0; Line2: P1 + s*d1.
    /// Returns the intersection point; success is false if lines are parallel.
    /// </summary>
    private DrawingPointF IntersectLines(DrawingPointF P0, DrawingPointF d0, DrawingPointF P1, DrawingPointF d1, out bool success)
    {
      float cross = Cross(d0, d1);
      if (Math.Abs(cross) < 1e-9)
      {
        success = false;
        return new DrawingPointF(0, 0);
      }
      DrawingPointF diff = Subtract(P1, P0);
      float t = (diff.X * d1.Y - diff.Y * d1.X) / cross;
      success = true;
      return Add(P0, Multiply(d0, t));
    }

    /// <summary>
    /// Overload: Computes the 2D cross product of two vectors.
    /// </summary>
    private float Cross(DrawingPointF a, DrawingPointF b) => a.X * b.Y - a.Y * b.X;

    /// <summary>
    /// Determines if a 2D polygon (as a list of DrawingPointF) is oriented clockwise.
    /// </summary>
    private bool IsClockwise2D(List<DrawingPointF> pts)
    {
      double sum = 0;
      int n = pts.Count;
      for (int i = 0; i < n; i++)
      {
        DrawingPointF current = pts[i];
        DrawingPointF next = pts[(i + 1) % n];
        sum += (next.X - current.X) * (next.Y + current.Y);
      }
      return sum > 0;
    }

    /// <summary>
    /// Cleans a CurveLoop by removing duplicate and nearly collinear vertices.
    /// Returns a new CurveLoop or null if the result is invalid.
    /// </summary>
    private CurveLoop CleanCurveLoop(CurveLoop loop, double tol)
    {
      List<XYZ> pts = new List<XYZ>();
      foreach (Curve c in loop)
      {
        pts.Add(c.GetEndPoint(0));
      }
      pts = RemoveDuplicates(pts, tol);
      List<XYZ> cleaned = new List<XYZ>();
      int n = pts.Count;
      if (n < 3)
        return null;
      for (int i = 0; i < n; i++)
      {
        XYZ prev = pts[(i + n - 1) % n];
        XYZ curr = pts[i];
        XYZ next = pts[(i + 1) % n];
        XYZ v1 = curr - prev;
        XYZ v2 = next - curr;
        if (v1.IsAlmostEqualTo(XYZ.Zero) || v2.IsAlmostEqualTo(XYZ.Zero))
          continue;
        XYZ n1 = v1.Normalize();
        XYZ n2 = v2.Normalize();
        double dot = n1.DotProduct(n2);
        if (Math.Abs(dot) > 0.9999) // nearly collinear
          continue;
        cleaned.Add(curr);
      }
      if (cleaned.Count < 3)
        return null;
      CurveLoop newLoop = new CurveLoop();
      for (int i = 0; i < cleaned.Count; i++)
      {
        XYZ p1 = cleaned[i];
        XYZ p2 = cleaned[(i + 1) % cleaned.Count];
        newLoop.Append(Line.CreateBound(p1, p2));
      }
      return newLoop;
    }

    /// <summary>
    /// Removes duplicate points (within tolerance) from a list of XYZ.
    /// </summary>
    private List<XYZ> RemoveDuplicates(List<XYZ> pts, double tol)
    {
      List<XYZ> unique = new List<XYZ>();
      foreach (var p in pts)
      {
        if (unique.Count == 0 || !p.IsAlmostEqualTo(unique.Last(), tol))
          unique.Add(p);
      }
      return unique;
    }

    /// <summary>
    /// Extracts boundary curve loops from a filled region by examining its geometry.
    /// Iterates over all solid faces and selects the face whose normal is most perpendicular
    /// to the view's direction (i.e. most parallel to the view’s sketch plane) and returns all its loops.
    /// </summary>
    private List<CurveLoop> GetFilledRegionCurveLoops(FilledRegion fr, RevitView view)
    {
      List<CurveLoop> loopsFound = new List<CurveLoop>();
      Options opt = new Options { ComputeReferences = true };
      GeometryElement geomElem = fr.get_Geometry(opt);
      if (geomElem == null)
        return loopsFound;

      double bestScore = double.MaxValue;
      IList<CurveLoop> bestLoops = null;
      foreach (GeometryObject geomObj in geomElem)
      {
        if (geomObj is Solid solid)
        {
          foreach (Face face in solid.Faces)
          {
            if (face is PlanarFace pf)
            {
              double score = Math.Abs(pf.FaceNormal.DotProduct(view.ViewDirection));
              if (score < bestScore)
              {
                IList<CurveLoop> faceLoops = pf.GetEdgesAsCurveLoops();
                if (faceLoops != null && faceLoops.Count > 0)
                {
                  bestScore = score;
                  bestLoops = faceLoops;
                }
              }
            }
          }
        }
      }
      if (bestLoops != null)
      {
        foreach (CurveLoop loop in bestLoops)
        {
          loopsFound.Add(loop);
        }
      }
      return loopsFound;
    }

    /// <summary>
    /// Projects a world point onto the view’s sketch plane (defined by view.Origin and view.ViewDirection).
    /// </summary>
    private XYZ ProjectPointToViewPlane(XYZ p, RevitView view)
    {
      XYZ origin = view.Origin;
      XYZ vDir = view.ViewDirection;
      double distance = (p - origin).DotProduct(vDir);
      return p - distance * vDir;
    }

    /// <summary>
    /// Converts a world point (assumed to lie on the view’s sketch plane) into view-local 2D coordinates.
    /// </summary>
    private DrawingPointF WorldPointToViewLocal(XYZ p, RevitView view)
    {
      XYZ origin = view.Origin;
      XYZ right = view.RightDirection;
      XYZ up = view.UpDirection;
      double x = (p - origin).DotProduct(right);
      double y = (p - origin).DotProduct(up);
      return new DrawingPointF((float)x, (float)y);
    }

    /// <summary>
    /// Converts a local 2D point in view coordinates back into world coordinates (on the view’s sketch plane).
    /// </summary>
    private XYZ ViewLocalToWorldPoint(DrawingPointF pt, RevitView view)
    {
      XYZ origin = view.Origin;
      XYZ right = view.RightDirection;
      XYZ up = view.UpDirection;
      return origin + right.Multiply(pt.X) + up.Multiply(pt.Y);
    }

    /// <summary>
    /// Computes the 2D convex hull of a set of DrawingPointF using the monotone chain algorithm.
    /// </summary>
    private List<DrawingPointF> Compute2DConvexHull(List<DrawingPointF> points)
    {
      List<DrawingPointF> sorted = points.OrderBy(p => p.X).ThenBy(p => p.Y).ToList();
      List<DrawingPointF> lower = new List<DrawingPointF>();
      foreach (DrawingPointF p in sorted)
      {
        while (lower.Count >= 2 && Cross(lower[lower.Count - 2], lower[lower.Count - 1], p) <= 0)
          lower.RemoveAt(lower.Count - 1);
        lower.Add(p);
      }
      List<DrawingPointF> upper = new List<DrawingPointF>();
      for (int i = sorted.Count - 1; i >= 0; i--)
      {
        DrawingPointF p = sorted[i];
        while (upper.Count >= 2 && Cross(upper[upper.Count - 2], upper[upper.Count - 1], p) <= 0)
          upper.RemoveAt(upper.Count - 1);
        upper.Add(p);
      }
      lower.RemoveAt(lower.Count - 1);
      upper.RemoveAt(upper.Count - 1);
      List<DrawingPointF> hull = new List<DrawingPointF>();
      hull.AddRange(lower);
      hull.AddRange(upper);
      return hull;
    }

    /// <summary>
    /// Returns the 2D cross product (z-component) of vectors defined by three DrawingPointF points.
    /// </summary>
    private float Cross(DrawingPointF O, DrawingPointF A, DrawingPointF B)
    {
      return (A.X - O.X) * (B.Y - O.Y) - (A.Y - O.Y) * (B.X - O.X);
    }

    /// <summary>
    /// Checks whether a polygon defined by a list of XYZ points is oriented clockwise.
    /// </summary>
    private bool IsClockwise(List<XYZ> pts)
    {
      double sum = 0;
      for (int i = 0; i < pts.Count; i++)
      {
        XYZ current = pts[i];
        XYZ next = pts[(i + 1) % pts.Count];
        sum += (next.X - current.X) * (next.Y + current.Y);
      }
      return sum > 0;
    }

    #endregion
  }
}
