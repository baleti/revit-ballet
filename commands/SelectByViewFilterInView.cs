// SelectByFilterInView.cs
//
// Revit 2024 API – C# 7.3
// Select all elements *visible in the active view* that pass one or more
// user-chosen Parameter Filters, in both host and linked models.

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace YourNamespace   // ← adjust
{
    public class FilterWrapper
    {
        public ElementId Id  { get; set; }
        public string    Name { get; set; }
        public FilterWrapper(ElementId id, string name)
        {
            Id = id; Name = name;
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SelectByViewFilterInView : IExternalCommand
    {
        #region IExternalCommand
        public Result Execute(ExternalCommandData commandData,
                              ref string          message,
                              ElementSet          elements)
        {
            // Handles ---------------------------------------------------------
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uiDoc.Document;
            View       view  = uiDoc.ActiveView;

            // 1. Pick filters --------------------------------------------------
            var allFilters =
                new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .Select(f => new FilterWrapper(f.Id, f.Name))
                    .OrderBy(f => f.Name)
                    .ToList();

            if (allFilters.Count == 0)
            {
                TaskDialog.Show("Select by Filter", "No filters exist in this project.");
                return Result.Cancelled;
            }

            var props   = new List<string> { "Name" };
            var chosen  = CustomGUIs.DataGrid<FilterWrapper>(allFilters, props);
            if (chosen == null || chosen.Count == 0) return Result.Cancelled;

            // 2. Build combined filter ----------------------------------------
            ElementFilter filter = null;
            foreach (FilterWrapper w in chosen)
            {
                var pfe = doc.GetElement(w.Id) as ParameterFilterElement;
                if (pfe == null) continue;

                filter = (filter == null)
                       ? pfe.GetElementFilter()
                       : new LogicalOrFilter(filter, pfe.GetElementFilter());
            }
            if (filter == null) return Result.Cancelled;

            // 3. Host elements already scoped to the view ---------------------
            var hostIds =
                new FilteredElementCollector(doc, view.Id)
                    .WhereElementIsNotElementType()
                    .WherePasses(filter)
                    .Select(e => e.Id)
                    .ToList();

            // 4. Linked elements – visibility tested by bounding box ----------
            var linkRefs = new List<Reference>();

            var linkInstances =
                new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .Where(li => li.GetLinkDocument() != null)   // loaded only
                    .ToList();

            foreach (var li in linkInstances)
            {
                Document linkDoc = li.GetLinkDocument();
                if (linkDoc == null) continue;

                var xf = li.GetTotalTransform();

                foreach (Element e in new FilteredElementCollector(linkDoc)
                                          .WhereElementIsNotElementType()
                                          .WherePasses(filter))
                {
                    if (!IsLinkedElementVisibleInView(e, li, xf, view))
                        continue;

                    Reference linkRef =
                        new Reference(e).CreateLinkReference(li);
                    if (linkRef != null) linkRefs.Add(linkRef);
                }
            }

            // 5. Merge and apply selection via SelectionModeManager -----------
            var idSet = new HashSet<ElementId>(uiDoc.GetSelectionIds());
            foreach (ElementId id in hostIds) idSet.Add(id);

            var refMap = new Dictionary<string, Reference>();
            foreach (Reference r in uiDoc.GetReferences())
                refMap[r.ConvertToStableRepresentation(doc)] = r;
            foreach (Reference r in linkRefs)
                refMap[r.ConvertToStableRepresentation(doc)] = r;

            uiDoc.SetSelectionIds(idSet.ToList());
            uiDoc.SetReferences(refMap.Values.ToList());

            return Result.Succeeded;
        }
        #endregion

        #region Visibility helpers (same logic as your other command)
        private static bool IsLinkedElementVisibleInView(
            Element linkedElement,
            RevitLinkInstance linkInstance,
            Transform linkXf,
            View view)
        {
            BoundingBoxXYZ bb = linkedElement.get_BoundingBox(null);
            if (bb == null) return false;

            XYZ min = linkXf.OfPoint(bb.Min);
            XYZ max = linkXf.OfPoint(bb.Max);

            switch (view.ViewType)
            {
                case ViewType.FloorPlan:
                case ViewType.CeilingPlan:
                case ViewType.AreaPlan:
                case ViewType.EngineeringPlan:
                    return IsInPlanView(min, max, (ViewPlan)view);

                case ViewType.ThreeD:
                    return IsIn3DView(min, max, (View3D)view);

                case ViewType.Section:
                case ViewType.Elevation:
                    return IsInSectionElevationView(min, max, view);

                default:
                    return false;   // other view types not handled
            }
        }

        private static bool IsInPlanView(XYZ min, XYZ max, ViewPlan plan)
        {
            try
            {
                PlanViewRange vr = plan.GetViewRange();
                if (vr == null) return false;

                Document doc = plan.Document;
                double elev = plan.GenLevel?.ProjectElevation ?? 0;

                // Top
                double top = (doc.GetElement(vr.GetLevelId(PlanViewPlane.TopClipPlane)) as Level)?.ProjectElevation ?? elev;
                top += vr.GetOffset(PlanViewPlane.TopClipPlane);

                // Bottom
                double bottom = (doc.GetElement(vr.GetLevelId(PlanViewPlane.BottomClipPlane)) as Level)?.ProjectElevation ?? elev;
                bottom += vr.GetOffset(PlanViewPlane.BottomClipPlane);

                // Cut & depth – used as extra-show layers
                double cut = (doc.GetElement(vr.GetLevelId(PlanViewPlane.CutPlane)) as Level)?.ProjectElevation ?? elev;
                cut += vr.GetOffset(PlanViewPlane.CutPlane);

                double depth = (doc.GetElement(vr.GetLevelId(PlanViewPlane.ViewDepthPlane)) as Level)?.ProjectElevation ?? elev;
                depth += vr.GetOffset(PlanViewPlane.ViewDepthPlane);

                bool zOK = max.Z >= bottom && min.Z <= top &&
                           (max.Z >= cut || max.Z >= depth);

                if (!zOK) return false;

                if (!plan.CropBoxActive) return true;

                BoundingBoxXYZ crop = plan.CropBox;
                if (crop == null) return true;

                Transform t = crop.Transform;
                XYZ cMin = t.OfPoint(crop.Min);
                XYZ cMax = t.OfPoint(crop.Max);

                return !(max.X < cMin.X || min.X > cMax.X ||
                         max.Y < cMin.Y || min.Y > cMax.Y);
            }
            catch { return false; }
        }

        private static bool IsIn3DView(XYZ min, XYZ max, View3D v3)
        {
            try
            {
                BoundingBoxXYZ box = v3.IsSectionBoxActive
                                   ? v3.GetSectionBox()
                                   : v3.CropBoxActive ? v3.CropBox : null;

                if (box == null) return true;

                Transform t = box.Transform;
                XYZ bMin = t.OfPoint(box.Min);
                XYZ bMax = t.OfPoint(box.Max);

                return !(max.X < System.Math.Min(bMin.X, bMax.X) ||
                         min.X > System.Math.Max(bMin.X, bMax.X) ||
                         max.Y < System.Math.Min(bMin.Y, bMax.Y) ||
                         min.Y > System.Math.Max(bMin.Y, bMax.Y) ||
                         max.Z < System.Math.Min(bMin.Z, bMax.Z) ||
                         min.Z > System.Math.Max(bMin.Z, bMax.Z));
            }
            catch { return false; }
        }

        private static bool IsInSectionElevationView(XYZ min, XYZ max, View view)
        {
            try
            {
                if (!view.CropBoxActive) return true;

                BoundingBoxXYZ crop = view.CropBox;
                if (crop == null) return true;

                Transform inv = crop.Transform.Inverse;
                XYZ a = inv.OfPoint(min);
                XYZ b = inv.OfPoint(max);

                double minX = System.Math.Min(a.X, b.X);
                double maxX = System.Math.Max(a.X, b.X);
                double minY = System.Math.Min(a.Y, b.Y);
                double maxY = System.Math.Max(a.Y, b.Y);
                double minZ = System.Math.Min(a.Z, b.Z);
                double maxZ = System.Math.Max(a.Z, b.Z);

                bool inXY = !(maxX < crop.Min.X || minX > crop.Max.X ||
                              maxY < crop.Min.Y || minY > crop.Max.Y);

                if (!inXY) return false;

                // For sections: respect far-clip
                if (view.ViewType == ViewType.Section)
                {
                    Parameter far = view.get_Parameter(BuiltInParameter.VIEWER_BOUND_OFFSET_FAR);
                    if (far != null)
                    {
                        double farClip = far.AsDouble();   // host units
                        if (minZ > farClip) return false;
                    }
                }
                return true;
            }
            catch { return false; }
        }
        #endregion
    }
}
