using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
using System.Drawing; // For PointF
using ClipperLib;   // Ensure you reference the official Angus Johnson Clipper library.
using PolyType = ClipperLib.PolyType;
using PolyFillType = ClipperLib.PolyFillType;
using ClipType = ClipperLib.ClipType;

// Alias Revit's View and ViewSheet to avoid ambiguity with System.Windows.Forms.View.
using RevitView = Autodesk.Revit.DB.View;
using RevitViewSheet = Autodesk.Revit.DB.ViewSheet;

namespace RevitAddin
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CreateFilledRegionsOfSelectedElementsInViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get document and UIDocument.
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Require that at least one element is selected.
            ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
            if (selIds == null || selIds.Count == 0)
            {
                MessageBox.Show("Please select one or more elements first.", "No Selection");
                return Result.Failed;
            }

            // Build a list of view entries for user selection.
            List<Dictionary<string, object>> viewEntries = new List<Dictionary<string, object>>();
            FilteredElementCollector viewCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitView))
                .WhereElementIsNotElementType();
            foreach (RevitView view in viewCollector)
            {
                if (view.IsTemplate)
                    continue;

                Dictionary<string, object> entry = new Dictionary<string, object>();
                entry["ViewId"] = view.Id.IntegerValue;
                entry["View Name"] = view.Name;

                // If the view is placed on a sheet, add sheet number and name.
                Viewport vp = null;
                FilteredElementCollector vpCollector = new FilteredElementCollector(doc).OfClass(typeof(Viewport));
                foreach (Viewport vpt in vpCollector)
                {
                    if (vpt.ViewId == view.Id)
                    {
                        vp = vpt;
                        break;
                    }
                }
                if (vp != null)
                {
                    Element sheetElem = doc.GetElement(vp.OwnerViewId);
                    if (sheetElem is RevitViewSheet viewSheet)
                    {
                        entry["Sheet Number"] = viewSheet.SheetNumber;
                        entry["Sheet Name"] = viewSheet.Name;
                    }
                    else
                    {
                        entry["Sheet Number"] = "";
                        entry["Sheet Name"] = "";
                    }
                }
                else
                {
                    entry["Sheet Number"] = "";
                    entry["Sheet Name"] = "";
                }
                viewEntries.Add(entry);
            }

            // Let the user choose one or more views using the provided CustomGUIs.DataGrid.
            List<string> propertyNames = new List<string> { "ViewId", "View Name", "Sheet Number", "Sheet Name" };
            List<Dictionary<string, object>> selectedViewEntries = CustomGUIs.DataGrid(viewEntries, propertyNames, false);
            if (selectedViewEntries.Count == 0)
            {
                message = "No views selected.";
                return Result.Cancelled;
            }

            // Get a default FilledRegionType.
            FilledRegionType fillRegionType = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault();
            if (fillRegionType == null)
            {
                message = "No filled region type found in the document.";
                return Result.Failed;
            }

            // Process each selected view.
            using (Transaction trans = new Transaction(doc, "Create Filled Regions"))
            {
                trans.Start();
                foreach (var viewEntry in selectedViewEntries)
                {
                    try
                    {
                        int viewIdInt = Convert.ToInt32(viewEntry["ViewId"]);
                        ElementId viewId = new ElementId(viewIdInt);
                        RevitView view = doc.GetElement(viewId) as RevitView;
                        if (view == null)
                            continue;

                        // We'll work in the view’s local coordinate system.
                        // Define a scale factor for Clipper (which uses integers).
                        double clipperScale = 1e6;

                        // Build a list of polygons (in view-local 2D coordinates) for each selected element.
                        List<List<IntPoint>> subjectPolygons = new List<List<IntPoint>>();
                        foreach (ElementId id in selIds)
                        {
                            Element elem = doc.GetElement(id);
                            // Get the element’s bounding box in the context of the view.
                            BoundingBoxXYZ bbox = elem.get_BoundingBox(view) ?? elem.get_BoundingBox(null);
                            if (bbox == null)
                                continue;

                            // Get all eight corners of the bounding box.
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

                            // For each corner, project it onto the view’s sketch plane and convert to local 2D coordinates.
                            List<PointF> localPoints = new List<PointF>();
                            foreach (XYZ p in corners)
                            {
                                XYZ projected = ProjectPointToViewPlane(p, view);
                                PointF localPt = WorldPointToViewLocal(projected, view);
                                localPoints.Add(localPt);
                            }

                            // Compute the convex hull of these 2D points.
                            List<PointF> hull = Compute2DConvexHull(localPoints);
                            if (hull.Count < 3)
                                continue;

                            // Convert the 2D hull to a Clipper polygon.
                            List<IntPoint> clipperPoly = new List<IntPoint>();
                            foreach (PointF pt in hull)
                            {
                                clipperPoly.Add(new IntPoint((long)(pt.X * clipperScale), (long)(pt.Y * clipperScale)));
                            }
                            subjectPolygons.Add(clipperPoly);
                        }

                        if (subjectPolygons.Count == 0)
                        {
                            TaskDialog.Show("Error", $"No valid outline geometry was created for view \"{view.Name}\".");
                            continue;
                        }

                        // Use Clipper to union all the subject polygons.
                        Clipper clipper = new Clipper();
                        foreach (List<IntPoint> poly in subjectPolygons)
                        {
                            clipper.AddPolygon(poly, PolyType.ptSubject);
                        }
                        List<List<IntPoint>> solution = new List<List<IntPoint>>();
                        clipper.Execute(ClipType.ctUnion, solution, PolyFillType.pftNonZero, PolyFillType.pftNonZero);

                        // Convert the unioned polygons back to view-local 2D points then into world coordinates.
                        List<CurveLoop> curveLoops = new List<CurveLoop>();
                        foreach (var poly in solution)
                        {
                            List<XYZ> worldPoints = new List<XYZ>();
                            foreach (IntPoint ip in poly)
                            {
                                double x = ip.X / clipperScale;
                                double y = ip.Y / clipperScale;
                                XYZ wp = ViewLocalToWorldPoint(new PointF((float)x, (float)y), view);
                                worldPoints.Add(wp);
                            }
                            if (worldPoints.Count < 3)
                                continue;
                            // Filled regions typically require a clockwise boundary.
                            if (!IsClockwise(worldPoints))
                                worldPoints.Reverse();

                            // Build a closed CurveLoop.
                            CurveLoop loop = new CurveLoop();
                            for (int i = 0; i < worldPoints.Count; i++)
                            {
                                XYZ p1 = worldPoints[i];
                                XYZ p2 = worldPoints[(i + 1) % worldPoints.Count];
                                Line line = Line.CreateBound(p1, p2);
                                loop.Append(line);
                            }
                            curveLoops.Add(loop);
                        }

                        if (curveLoops.Count == 0)
                        {
                            TaskDialog.Show("Error", $"No valid boundary created for view \"{view.Name}\".");
                            continue;
                        }

                        // Create the filled region in the current view using the unioned curve loops.
                        FilledRegion filledRegion = FilledRegion.Create(doc, fillRegionType.Id, viewId, curveLoops);
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Error", $"Error creating filled region in view \"{viewEntry["View Name"]}\": {ex.Message}");
                    }
                }
                trans.Commit();
            }

            return Result.Succeeded;
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
        /// Converts a world point (assumed to lie on the view’s sketch plane) into the view’s local 2D coordinates.
        /// </summary>
        private PointF WorldPointToViewLocal(XYZ p, RevitView view)
        {
            XYZ origin = view.Origin;
            XYZ right = view.RightDirection;
            XYZ up = view.UpDirection;
            double x = (p - origin).DotProduct(right);
            double y = (p - origin).DotProduct(up);
            return new PointF((float)x, (float)y);
        }

        /// <summary>
        /// Converts a local 2D point in view coordinates back into world coordinates (on the view’s sketch plane).
        /// </summary>
        private XYZ ViewLocalToWorldPoint(PointF pt, RevitView view)
        {
            XYZ origin = view.Origin;
            XYZ right = view.RightDirection;
            XYZ up = view.UpDirection;
            return origin + right.Multiply(pt.X) + up.Multiply(pt.Y);
        }

        /// <summary>
        /// Computes the 2D convex hull of a set of PointF using the monotone chain algorithm.
        /// </summary>
        private List<PointF> Compute2DConvexHull(List<PointF> points)
        {
            List<PointF> sorted = points.OrderBy(p => p.X).ThenBy(p => p.Y).ToList();
            List<PointF> lower = new List<PointF>();
            foreach (PointF p in sorted)
            {
                while (lower.Count >= 2 && Cross(lower[lower.Count - 2], lower[lower.Count - 1], p) <= 0)
                    lower.RemoveAt(lower.Count - 1);
                lower.Add(p);
            }
            List<PointF> upper = new List<PointF>();
            for (int i = sorted.Count - 1; i >= 0; i--)
            {
                PointF p = sorted[i];
                while (upper.Count >= 2 && Cross(upper[upper.Count - 2], upper[upper.Count - 1], p) <= 0)
                    upper.RemoveAt(upper.Count - 1);
                upper.Add(p);
            }
            lower.RemoveAt(lower.Count - 1);
            upper.RemoveAt(upper.Count - 1);
            List<PointF> hull = new List<PointF>();
            hull.AddRange(lower);
            hull.AddRange(upper);
            return hull;
        }

        /// <summary>
        /// Returns the 2D cross product (z-component) of vectors OA and OB.
        /// </summary>
        private float Cross(PointF O, PointF A, PointF B)
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
    }
}
