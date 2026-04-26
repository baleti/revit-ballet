using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    [CommandMeta("Floor, Filled Region")]
    [CommandOutput("Floor, Filled Region")]
    public class Split : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<FilledRegion> selectedRegions = new List<FilledRegion>();
#if REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
            List<Floor> selectedFloors = new List<Floor>();
#endif

            foreach (ElementId id in uidoc.GetSelectionIds())
            {
                Element elem = doc.GetElement(id);
                if (elem is FilledRegion fr) selectedRegions.Add(fr);
#if REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                else if (elem is Floor fl) selectedFloors.Add(fl);
#endif
            }

            bool anySelected = selectedRegions.Count > 0;
#if REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
            anySelected = anySelected || selectedFloors.Count > 0;
#endif

            if (!anySelected)
            {
                CustomGUIs.SetCurrentUIDocument(uidoc);
                var items = new List<Dictionary<string, object>>();

                foreach (FilledRegion r in new FilteredElementCollector(doc).OfClass(typeof(FilledRegion)).Cast<FilledRegion>())
                    items.Add(new Dictionary<string, object> { { "Name", r.Name }, { "Type", "Filled Region" }, { "ElementIdObject", r.Id } });

#if REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                foreach (Floor f in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().Cast<Floor>())
                    items.Add(new Dictionary<string, object> { { "Name", f.Name }, { "Type", "Floor" }, { "ElementIdObject", f.Id } });
#endif

                if (items.Count == 0) return Result.Succeeded;

                var chosen = CustomGUIs.DataGrid(items, new List<string> { "Name", "Type" }, false);
                if (chosen == null || chosen.Count == 0) return Result.Succeeded;

                foreach (var row in chosen)
                {
                    if (!row.TryGetValue("ElementIdObject", out object idObj) || !(idObj is ElementId eid)) continue;
                    Element elem = doc.GetElement(eid);
                    if (elem is FilledRegion fr2) selectedRegions.Add(fr2);
#if REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                    else if (elem is Floor fl2) selectedFloors.Add(fl2);
#endif
                }
            }

            List<ElementId> newIds = new List<ElementId>();

            if (selectedRegions.Count > 0)
            {
                using (Transaction trans = new Transaction(doc, "Split Filled Regions"))
                {
                    trans.Start();
                    foreach (FilledRegion region in selectedRegions)
                    {
                        IList<CurveLoop> loops = region.GetBoundaries();
                        GraphicsStyle boundaryStyle = GetBoundaryLineStyle(doc, region);

                        foreach (CurveLoop loop in loops)
                        {
                            FilledRegion newRegion = FilledRegion.Create(doc, region.GetTypeId(), region.OwnerViewId, new List<CurveLoop> { loop });
                            if (boundaryStyle != null && newRegion != null)
                                SetBoundaryLineStyle(doc, newRegion, boundaryStyle);
                            if (newRegion != null) newIds.Add(newRegion.Id);
                        }
                        doc.Delete(region.Id);
                    }
                    trans.Commit();
                }
            }

#if REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
            if (selectedFloors.Count > 0)
            {
                using (Transaction trans = new Transaction(doc, "Split Floors"))
                {
                    trans.Start();
                    foreach (Floor floor in selectedFloors)
                    {
                        FloorType floorType = doc.GetElement(floor.FloorType.Id) as FloorType;
                        Level level = doc.GetElement(floor.LevelId) as Level;
                        Parameter structuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                        bool isStructural = structuralParam != null && structuralParam.AsInteger() == 1;

                        IList<CurveLoop> boundaries = GetFloorBoundaries(floor);
                        if (boundaries.Count <= 1) continue;

                        foreach (CurveLoop boundary in boundaries)
                        {
                            try
                            {
                                Floor newFloor = Floor.Create(doc, new List<CurveLoop> { boundary }, floorType.Id, level.Id, isStructural, null, 0.0);
                                if (newFloor != null)
                                {
                                    CopyParameters(floor, newFloor);
                                    newFloor.FloorType = floorType;
                                    newIds.Add(newFloor.Id);
                                }
                            }
                            catch (Autodesk.Revit.Exceptions.ArgumentException ex)
                            {
                                TaskDialog.Show("Warning", $"Failed to create floor from boundary: {ex.Message}");
                            }
                        }
                        doc.Delete(floor.Id);
                    }
                    trans.Commit();
                }
            }
#endif

            if (newIds.Count > 0)
                uidoc.SetSelectionIds(newIds);

            return Result.Succeeded;
        }

        private GraphicsStyle GetBoundaryLineStyle(Document doc, FilledRegion region)
        {
#if REVIT2017
            return null;
#else
            foreach (ElementId depId in region.GetDependentElements(null))
            {
                if (doc.GetElement(depId) is CurveElement curveElem)
                    return curveElem.LineStyle as GraphicsStyle;
            }
            return null;
#endif
        }

        private void SetBoundaryLineStyle(Document doc, FilledRegion newRegion, GraphicsStyle style)
        {
#if REVIT2017
            return;
#else
            foreach (ElementId depId in newRegion.GetDependentElements(null))
            {
                if (doc.GetElement(depId) is CurveElement curveElem)
                    curveElem.LineStyle = style;
            }
#endif
        }

#if REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
        private IList<CurveLoop> GetFloorBoundaries(Floor floor)
        {
            var curveLoops = new List<CurveLoop>();
            foreach (GeometryObject geomObj in floor.get_Geometry(new Options()))
            {
                if (geomObj is Solid solid)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face.ComputeNormal(new UV(0.5, 0.5)).IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                        {
                            foreach (EdgeArray edgeArray in face.EdgeLoops)
                            {
                                CurveLoop loop = new CurveLoop();
                                foreach (Edge edge in edgeArray)
                                    loop.Append(edge.AsCurve());
                                curveLoops.Add(loop);
                            }
                            break;
                        }
                    }
                }
            }
            return curveLoops;
        }

        private void CopyParameters(Floor source, Floor target)
        {
            foreach (Parameter p in source.Parameters)
            {
                if (p.IsReadOnly || !p.HasValue) continue;
                Parameter tp = target.get_Parameter(p.Definition);
                if (tp == null || tp.IsReadOnly) continue;
                switch (p.StorageType)
                {
                    case StorageType.Double:  tp.Set(p.AsDouble()); break;
                    case StorageType.Integer: tp.Set(p.AsInteger()); break;
                    case StorageType.String:  tp.Set(p.AsString()); break;
                    case StorageType.ElementId: tp.Set(p.AsElementId()); break;
                }
            }
        }
#endif
    }
}
