using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class CreateAreaExtrusions : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
        if (selectedIds.Count == 0)
        {
            message = "Please select one or more areas.";
            return Result.Cancelled;
        }

        List<Area> selectedAreas = new List<Area>();
        foreach (ElementId id in selectedIds)
        {
            if (doc.GetElement(id) is Area area)
                selectedAreas.Add(area);
        }

        if (selectedAreas.Count == 0)
        {
            message = "No valid areas selected.";
            return Result.Cancelled;
        }

        // Get all levels ordered by elevation
        List<Level> levels = new FilteredElementCollector(doc)
            .OfClass(typeof(Level))
            .WhereElementIsNotElementType()
            .Cast<Level>()
            .OrderBy(l => l.Elevation)
            .ToList();

        using (Transaction trans = new Transaction(doc, "Create Area Extrusions"))
        {
            trans.Start();
            foreach (Area currentArea in selectedAreas)
            {
                try 
                {
                    // Get area's level
                    Parameter levelParam = currentArea.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                    if (levelParam == null) continue;

                    ElementId levelId = levelParam.AsElementId();
                    Level currentLevel = doc.GetElement(levelId) as Level;
                    if (currentLevel == null) continue;

                    // Find next level above
                    Level nextLevel = levels.FirstOrDefault(l => l.Elevation > currentLevel.Elevation);
                    if (nextLevel == null) continue;

                    // Calculate height as difference between levels
                    double height = nextLevel.Elevation - currentLevel.Elevation;

                    // Get area boundary
                    SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
                    IList<IList<BoundarySegment>> boundaries = currentArea.GetBoundarySegments(options);
                    if (boundaries == null || boundaries.Count == 0) continue;

                    // Create profile from first boundary
                    CurveLoop loop = new CurveLoop();
                    foreach (BoundarySegment segment in boundaries[0])
                    {
                        loop.Append(segment.GetCurve());
                    }

                    // Create extrusion
                    Solid extrusion = GeometryCreationUtilities.CreateExtrusionGeometry(
                        new List<CurveLoop> { loop },
                        XYZ.BasisZ,
                        height
                    );

                    DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                    ds.SetShape(new GeometryObject[] { extrusion });
                    ds.Name = $"Area_{currentArea.Id}_Extrusion";
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", $"Failed to create extrusion: {ex.Message}");
                }
            }
            trans.Commit();
        }

        return Result.Succeeded;
    }
}
