using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class SplitFloors : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Selection sel = uidoc.Selection;

            // Get selected floors
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            List<Floor> selectedFloors = new List<Floor>();

            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is Floor floor)
                {
                    selectedFloors.Add(floor);
                }
            }

            if (selectedFloors.Count == 0)
            {
                TaskDialog.Show("Error", "Please select at least one floor.");
                return Result.Failed;
            }

            using (Transaction trans = new Transaction(doc, "Split Floors"))
            {
                trans.Start();

                foreach (Floor floor in selectedFloors)
                {
                    // Get floor type
                    FloorType floorType = doc.GetElement(floor.FloorType.Id) as FloorType;
                    
                    // Get floor level
                    Level level = doc.GetElement(floor.LevelId) as Level;
                    
                    // Get the floor's structural properties
                    Parameter structuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                    bool isStructural = structuralParam != null && structuralParam.AsInteger() == 1;

                    // Get floor boundaries
                    IList<CurveLoop> floorBoundaries = GetFloorBoundaries(floor);

                    // Skip splitting if only one boundary
                    if (floorBoundaries.Count <= 1) continue;

                    // Set default slope arrow and slope
                    Line slopeArrow = null;
                    double slope = 0.0;

                    // Create new floor for each boundary
                    foreach (CurveLoop boundary in floorBoundaries)
                    {
                        try
                        {
                            // Create the new floor with slope arrow
                            Floor newFloor = Floor.Create(
                                doc,
                                new List<CurveLoop> { boundary },
                                floorType.Id,
                                level.Id,
                                isStructural,
                                null,  // slope arrow - not trying to preserve it as it might cause issues
                                slope  // use detected slope or default 0.0
                            );

                            if (newFloor != null)
                            {
                                // Copy parameters from original floor to new floor
                                CopyParameters(floor, newFloor);
                                
                                // Ensure the floor type is set correctly
                                newFloor.FloorType = floorType;
                            }
                        }
                        catch (Autodesk.Revit.Exceptions.ArgumentException ex)
                        {
                            // Log or handle invalid geometry
                            TaskDialog.Show("Warning", $"Failed to create floor from boundary: {ex.Message}");
                            continue;
                        }
                    }

                    // Delete the original floor
                    doc.Delete(floor.Id);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Gets all boundary loops from a floor element
        /// </summary>
        private IList<CurveLoop> GetFloorBoundaries(Floor floor)
        {
            List<CurveLoop> curveLoops = new List<CurveLoop>();
            
            Options geomOptions = new Options();
            GeometryElement geomElem = floor.get_Geometry(geomOptions);
            
            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid)
                {
                    // Get the bottom face of the floor
                    foreach (Face face in solid.Faces)
                    {
                        if (face.ComputeNormal(new UV(0.5, 0.5)).IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                        {
                            EdgeArrayArray edgeArrayArray = face.EdgeLoops;
                            foreach (EdgeArray edgeArray in edgeArrayArray)
                            {
                                CurveLoop curveLoop = new CurveLoop();
                                foreach (Edge edge in edgeArray)
                                {
                                    Curve curve = edge.AsCurve();
                                    curveLoop.Append(curve);
                                }
                                curveLoops.Add(curveLoop);
                            }
                            break;
                        }
                    }
                }
            }

            return curveLoops;
        }

        /// <summary>
        /// Copies parameters from source floor to target floor
        /// </summary>
        private void CopyParameters(Floor sourceFloor, Floor targetFloor)
        {
            foreach (Parameter param in sourceFloor.Parameters)
            {
                // Skip read-only and invalid parameters
                if (param.IsReadOnly || !param.HasValue)
                    continue;

                Parameter targetParam = targetFloor.get_Parameter(param.Definition);
                if (targetParam != null && !targetParam.IsReadOnly)
                {
                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            targetParam.Set(param.AsDouble());
                            break;
                        case StorageType.Integer:
                            targetParam.Set(param.AsInteger());
                            break;
                        case StorageType.String:
                            targetParam.Set(param.AsString());
                            break;
                        case StorageType.ElementId:
                            targetParam.Set(param.AsElementId());
                            break;
                    }
                }
            }
        }
    }
}
