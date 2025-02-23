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
    public class CreateFilledRegionsOfSelectedElementsInSelectedViews : IExternalCommand
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

            // --- Faster view collection for the DataGrid ---
            // Create a mapping for non-sheet views that are placed on sheets.
            Dictionary<ElementId, RevitViewSheet> viewToSheetMap = new Dictionary<ElementId, RevitViewSheet>();
            FilteredElementCollector sheetCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitViewSheet));
            foreach (RevitViewSheet sheet in sheetCollector)
            {
                foreach (ElementId viewportId in sheet.GetAllViewports())
                {
                    Viewport viewport = doc.GetElement(viewportId) as Viewport;
                    if (viewport != null)
                    {
                        viewToSheetMap[viewport.ViewId] = sheet;
                    }
                }
            }

            // Collect all views in the project.
            FilteredElementCollector viewCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitView));

            // Prepare data for the data grid.
            List<Dictionary<string, object>> viewData = new List<Dictionary<string, object>>();
            // Map view Title to view object.
            Dictionary<string, RevitView> titleToViewMap = new Dictionary<string, RevitView>();

            foreach (RevitView view in viewCollector.Cast<RevitView>())
            {
                // Skip view templates, legends, schedules, and sheets themselves.
                if (view.IsTemplate || view.ViewType == ViewType.Legend || view.ViewType == ViewType.Schedule || view is RevitViewSheet)
                    continue;

                string sheetInfo = string.Empty;
                if (viewToSheetMap.TryGetValue(view.Id, out RevitViewSheet sheet))
                {
                    sheetInfo = $"{sheet.SheetNumber} - {sheet.Name}";
                }
                else
                {
                    sheetInfo = "Not Placed";
                }

                // Use Title as key (assume uniqueness).
                titleToViewMap[view.Title] = view;
                Dictionary<string, object> viewInfo = new Dictionary<string, object>
                {
                    { "Title", view.Title },
                    { "Sheet", sheetInfo }
                };
                viewData.Add(viewInfo);
            }

            // Define the column headers.
            List<string> viewColumns = new List<string> { "Title", "Sheet" };

            // Show the selection dialog using your custom DataGrid.
            List<Dictionary<string, object>> selectedViews = CustomGUIs.DataGrid(
                viewData,
                viewColumns,
                false  // Don't span all screens.
            );

            // If no view is selected, cancel the command.
            if (selectedViews == null || selectedViews.Count == 0)
            {
                message = "No views selected.";
                return Result.Cancelled;
            }

            // --- Next: Prompt for Filled Region Type selection ---
            // Collect all available FilledRegionTypes.
            FilteredElementCollector regionTypeCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType));
            List<Dictionary<string, object>> regionTypeData = new List<Dictionary<string, object>>();
            // Map filled region type name to FilledRegionType object.
            Dictionary<string, FilledRegionType> nameToRegionTypeMap = new Dictionary<string, FilledRegionType>();
            foreach (FilledRegionType frt in regionTypeCollector.Cast<FilledRegionType>())
            {
                string name = frt.Name;
                // Add to mapping and data grid list.
                nameToRegionTypeMap[name] = frt;
                regionTypeData.Add(new Dictionary<string, object>
                {
                    { "Name", name }
                });
            }

            // Define the columns for the filled region type selection.
            List<string> regionTypeColumns = new List<string> { "Name" };
            List<Dictionary<string, object>> selectedRegionTypes = CustomGUIs.DataGrid(
                regionTypeData,
                regionTypeColumns,
                false  // Don't span all screens.
            );

            // Ensure a selection was made.
            if (selectedRegionTypes == null || selectedRegionTypes.Count == 0)
            {
                message = "No filled region type selected.";
                return Result.Cancelled;
            }

            // For simplicity, take the first selected filled region type.
            string selectedRegionTypeName = selectedRegionTypes[0]["Name"].ToString();
            if (!nameToRegionTypeMap.TryGetValue(selectedRegionTypeName, out FilledRegionType selectedRegionType))
            {
                message = "Unable to resolve the selected filled region type.";
                return Result.Failed;
            }

            // --- Continue with filled region creation using the selected filled region type ---
            using (Transaction trans = new Transaction(doc, "Create Filled Regions"))
            {
                trans.Start();
                foreach (Dictionary<string, object> viewEntry in selectedViews)
                {
                    try
                    {
                        // Look up the Revit view using its Title.
                        string title = viewEntry["Title"].ToString();
                        if (!titleToViewMap.TryGetValue(title, out RevitView view))
                            continue;
                        ElementId viewId = view.Id;

                        // Work in the view’s local coordinate system.
                        double clipperScale = 1e6; // Scale factor for Clipper.

                        // Build a list of polygons (in view-local 2D coordinates) for each selected element.
                        List<List<IntPoint>> subjectPolygons = new List<List<IntPoint>>();
                        foreach (ElementId id in selIds)
                        {
                            Element elem = doc.GetElement(id);
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
                            TaskDialog.Show("Error", $"No valid outline geometry was created for view \"{view.Title}\".");
                            continue;
                        }

                        // Use Clipper to union all subject polygons.
                        Clipper clipper = new Clipper();
                        foreach (List<IntPoint> poly in subjectPolygons)
                        {
                            clipper.AddPolygon(poly, PolyType.ptSubject);
                        }
                        List<List<IntPoint>> solution = new List<List<IntPoint>>();
                        clipper.Execute(ClipType.ctUnion, solution, PolyFillType.pftNonZero, PolyFillType.pftNonZero);

                        // Convert unioned polygons back to view-local 2D points then to world coordinates.
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
                            // Ensure the polygon is clockwise (filled regions require a clockwise boundary).
                            if (!IsClockwise(worldPoints))
                                worldPoints.Reverse();

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
                            TaskDialog.Show("Error", $"No valid boundary created for view \"{view.Title}\".");
                            continue;
                        }

                        // Create the filled region in the current view using the unioned curve loops and the selected filled region type.
                        FilledRegion filledRegion = FilledRegion.Create(doc, selectedRegionType.Id, viewId, curveLoops);
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Error", $"Error creating filled region in view \"{viewEntry["Title"]}\": {ex.Message}");
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
