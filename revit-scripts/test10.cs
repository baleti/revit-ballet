using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Linq;

namespace RevitGridDimensions
{
    [Transaction(TransactionMode.Manual)]
    public class GridDimensionCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            try
            {
                // Get pre-selected elements
                var selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count != 2)
                {
                    message = "Please select exactly two elements (one grid and one other element).";
                    return Result.Failed;
                }

                // Get selected elements
                var selectedElements = selectedIds.Select(id => doc.GetElement(id)).ToList();
                
                // Find grid among selected elements
                Grid grid = selectedElements.OfType<Grid>().FirstOrDefault();
                Element otherElement = selectedElements.FirstOrDefault(e => !(e is Grid));

                if (grid == null || otherElement == null)
                {
                    message = "Please select one grid and one other element.";
                    return Result.Failed;
                }

                // Get visible grid line
                var gridCurves = grid.GetCurvesInView(DatumExtentType.ViewSpecific, activeView);
                if (!gridCurves.Any())
                {
                    message = "Grid is not visible in current view.";
                    return Result.Failed;
                }

                Line gridLine = gridCurves.First() as Line;
                if (gridLine == null)
                {
                    message = "Grid must be a straight line.";
                    return Result.Failed;
                }

                // Get grid direction and perpendicular direction
                XYZ gridDir = gridLine.Direction.Normalize();
                XYZ perpDir = gridDir.CrossProduct(activeView.ViewDirection).Normalize();

                // Get element's bounding box
                BoundingBoxXYZ bbox = otherElement.get_BoundingBox(activeView);
                if (bbox == null)
                {
                    message = "Cannot get element's location.";
                    return Result.Failed;
                }

                // Use center of bounding box as reference point
                XYZ elementPoint = (bbox.Min + bbox.Max) * 0.5;
                XYZ gridPoint = gridLine.GetEndPoint(0);

                // Project element point onto grid line
                IntersectionResult projection = gridLine.Project(elementPoint);
                if (projection == null)
                {
                    message = "Cannot project element onto grid.";
                    return Result.Failed;
                }

                // Create dimension line
                XYZ dimStart = projection.XYZPoint;
                XYZ dimEnd = elementPoint;
                
                // Ensure dimension line is perpendicular to grid
                XYZ dimDir = perpDir;
                Line dimLine = Line.CreateBound(dimStart + dimDir, dimEnd + dimDir);

                // Create references
                ReferenceArray refs = new ReferenceArray();
                refs.Append(new Reference(grid));
                refs.Append(new Reference(otherElement));

                // Create dimension
                using (Transaction trans = new Transaction(doc, "Create Grid to Element Dimension"))
                {
                    trans.Start();
                    doc.Create.NewDimension(activeView, dimLine, refs);
                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error creating dimension: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}
