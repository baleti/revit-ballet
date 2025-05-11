// LabelAreasOfSelectedViews.cs
// Revit 2024 – External Command (C# 7.3)
// -----------------------------------------------------------------------------
// Workflow
//   1. Determine the target Area Plan views:
//        • If the user has views / viewports selected, use those (Area Plans only).
//        • Otherwise, use the currently‑active view (if it is an Area Plan).
//   2. Scan every Area (host + visible links) in those views and gather the full
//      set of parameter names that appear.
//   3. Show the list of parameter names in a CustomGUIs.DataGrid dialog so the
//      user can pick which parameters to label with.
//   4. Place centered TextNotes inside the crop region of each target view. Each
//      label concatenates the chosen parameters' values for that Area, separated
//      by " – ", and finishes with the metric area in m².
//   5. Report how many labels were placed.
//
// NOTE: Works with rectangular crop; irregular crop falls back to bounding box.
// -----------------------------------------------------------------------------
// References required: RevitAPI.dll, RevitAPIUI.dll, CustomGUIs.dll

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

namespace YourCompany.RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class LabelAreasOfSelectedViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null) return Result.Failed;
            Document doc = uidoc.Document;

            // -----------------------------------------------------------------
            // 1. Resolve target views (Area Plans) from selection or active view
            // -----------------------------------------------------------------
            List<ViewPlan> targetViews = ResolveTargetAreaPlans(uidoc);
            if (targetViews.Count == 0)
            {
                TaskDialog.Show("Label Areas", "No Area Plan views were selected or active.");
                return Result.Cancelled;
            }

            // -----------------------------------------------------------------
            // 2. Gather available Area parameter names & sample values
            // -----------------------------------------------------------------
            Dictionary<string, string> sampleValueByParam = new Dictionary<string, string>();

            foreach (ViewPlan v in targetViews)
            {
                AreaDataCollector.CollectParameterSamples(doc, v, sampleValueByParam);
            }

            if (sampleValueByParam.Count == 0)
            {
                TaskDialog.Show("Label Areas", "No Areas found in the chosen views.");
                return Result.Cancelled;
            }

            // Build data for CustomGUIs.DataGrid
            List<Dictionary<string, object>> gridData = new List<Dictionary<string, object>>();
            foreach (var kvp in sampleValueByParam.OrderBy(k => k.Key))
            {
                gridData.Add(new Dictionary<string, object>
                {
                    { "Parameter", kvp.Key },
                    { "Sample", kvp.Value }
                });
            }

            List<string> columns = new List<string> { "Parameter", "Sample" };

            // -----------------------------------------------------------------
            // 3. Show dialog & get selected parameters
            // -----------------------------------------------------------------
            List<Dictionary<string, object>> selectedRows = CustomGUIs.DataGrid(gridData, columns, false);
            if (selectedRows == null || selectedRows.Count == 0) return Result.Cancelled;

            List<string> selectedParams = selectedRows.Select(r => r["Parameter"].ToString()).ToList();

            // -----------------------------------------------------------------
            // 4. Place TextNotes for each Area in each view
            // -----------------------------------------------------------------
            int totalLabels = 0;

            using (TransactionGroup tg = new TransactionGroup(doc, "Label Areas in Views"))
            {
                tg.Start();

                foreach (ViewPlan view in targetViews)
                {
                    totalLabels += Labeler.PlaceAreaLabelsInView(doc, view, selectedParams);
                }

                tg.Assimilate();
            }

            return Result.Succeeded;
        }

        // ---------------------------------------------------------
        // Helper – resolve target Area Plan views
        // ---------------------------------------------------------
        private static List<ViewPlan> ResolveTargetAreaPlans(UIDocument uidoc)
        {
            Document doc = uidoc.Document;
            ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
            List<ViewPlan> views = new List<ViewPlan>();

            if (selIds != null && selIds.Count > 0)
            {
                foreach (ElementId id in selIds)
                {
                    Element e = doc.GetElement(id);
                    if (e is ViewPlan vp && vp.ViewType == ViewType.AreaPlan)
                    {
                        views.Add(vp);
                    }
                    else if (e is Viewport vpPort)
                    {
                        View v = doc.GetElement(vpPort.ViewId) as View;
                        if (v is ViewPlan vp2 && vp2.ViewType == ViewType.AreaPlan)
                            views.Add(vp2);
                    }
                }
            }

            if (views.Count == 0)
            {
                // Fallback to active view if it's an Area Plan
                View active = uidoc.ActiveView;
                if (active is ViewPlan actVp && actVp.ViewType == ViewType.AreaPlan)
                    views.Add(actVp);
            }

            // Ensure uniqueness
            return views.Distinct().ToList();
        }

        // ---------------------------------------------------------
        // Helper – build parameter string value safely
        // ---------------------------------------------------------
        internal static string GetParamString(Element elem, string paramName)
        {
            Parameter p = elem.LookupParameter(paramName);
            if (p == null) return string.Empty;
            return p.StorageType == StorageType.String ? p.AsString() : p.AsValueString();
        }

        // ---------------------------------------------------------
        // INTERNAL CLASSES
        // ---------------------------------------------------------
        private static class AreaDataCollector
        {
            public static void CollectParameterSamples(Document hostDoc, ViewPlan view, IDictionary<string, string> samples)
            {
                CropRegionTester crop = new CropRegionTester(view);

                // Host Areas
                foreach (Area ar in new FilteredElementCollector(hostDoc, view.Id)
                                        .OfCategory(BuiltInCategory.OST_Areas)
                                        .WhereElementIsNotElementType()
                                        .Cast<Area>())
                {
                    XYZ pt = ar.Location is LocationPoint lp ? lp.Point : null;
                    if (pt == null || !crop.ContainsPoint(pt)) continue;
                    RegisterParams(ar, samples);
                }

                // Linked Areas
                foreach (RevitLinkInstance linkInst in new FilteredElementCollector(hostDoc, view.Id)
                                                    .OfClass(typeof(RevitLinkInstance))
                                                    .Cast<RevitLinkInstance>())
                {
                    Document linkDoc = linkInst.GetLinkDocument();
                    if (linkDoc == null) continue;
                    Transform trf = linkInst.GetTransform();

                    foreach (Area ar in new FilteredElementCollector(linkDoc)
                                            .OfCategory(BuiltInCategory.OST_Areas)
                                            .WhereElementIsNotElementType()
                                            .Cast<Area>())
                    {
                        XYZ linkPt = ar.Location is LocationPoint lp ? lp.Point : null;
                        if (linkPt == null) continue;
                        XYZ hostPt = trf.OfPoint(linkPt);
                        if (!crop.ContainsPoint(hostPt)) continue;
                        RegisterParams(ar, samples);
                    }
                }
            }

            private static void RegisterParams(Area ar, IDictionary<string, string> samples)
            {
                foreach (Parameter p in ar.Parameters)
                {
                    string name = p.Definition.Name;
                    if (!samples.ContainsKey(name))
                    {
                        string val = p.StorageType == StorageType.String ? p.AsString() : p.AsValueString();
                        samples[name] = val ?? string.Empty;
                    }
                }
            }
        }

        private static class Labeler
        {
            public static int PlaceAreaLabelsInView(Document doc, ViewPlan view, IList<string> paramNames)
            {
                CropRegionTester crop = new CropRegionTester(view);
                ElementId textTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);
                TextNoteOptions noteOpts = new TextNoteOptions(textTypeId)
                {
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    VerticalAlignment = VerticalTextAlignment.Middle
                };

                int labelCount = 0;
                using (Transaction tx = new Transaction(doc, $"Label Areas – {view.Name}"))
                {
                    tx.Start();

                    // HOST
                    labelCount += LabelAreasInDoc(doc, doc, view.Id, null, crop, paramNames, noteOpts);

                    // LINKS
                    foreach (RevitLinkInstance linkInst in new FilteredElementCollector(doc, view.Id)
                                                        .OfClass(typeof(RevitLinkInstance))
                                                        .Cast<RevitLinkInstance>())
                    {
                        Document linkDoc = linkInst.GetLinkDocument();
                        if (linkDoc == null) continue;
                        labelCount += LabelAreasInDoc(linkDoc, doc, view.Id, linkInst, crop, paramNames, noteOpts);
                    }

                    tx.Commit();
                }

                return labelCount;
            }

            private static int LabelAreasInDoc(Document srcDoc, Document hostDoc, ElementId hostViewId, RevitLinkInstance linkInst,
                                                CropRegionTester crop, IList<string> paramNames, TextNoteOptions noteOpts)
            {
                int count = 0;
                Transform trf = linkInst != null ? linkInst.GetTransform() : Transform.Identity;
                string linkPrefix = linkInst != null ? linkInst.Name + ": " : string.Empty;

                foreach (Area ar in new FilteredElementCollector(srcDoc)
                                        .OfCategory(BuiltInCategory.OST_Areas)
                                        .WhereElementIsNotElementType()
                                        .Cast<Area>())
                {
                    XYZ pt = ar.Location is LocationPoint lp ? lp.Point : null;
                    if (pt == null) continue;
                    XYZ hostPt = trf.OfPoint(pt);
                    if (!crop.ContainsPoint(hostPt)) continue;

                    List<string> parts = new List<string>();
                    foreach (string paramName in paramNames)
                    {
                        string val = LabelAreasOfSelectedViews.GetParamString(ar, paramName);
                        if (string.IsNullOrWhiteSpace(val)) val = "?";
                        parts.Add(val);
                    }
                    string labelText = string.Join(" – ", parts);

                    // Always create text notes in the host document
                    TextNote.Create(hostDoc, hostViewId, hostPt, labelText, noteOpts);
                    count++;
                }
                return count;
            }
        }

        // ---------------------------------------------------------
        // Crop-region containment tester (rectangular)
        // ---------------------------------------------------------
        private class CropRegionTester
        {
            private readonly bool _active;
            private readonly Transform _trf;
            private readonly XYZ _min, _max;

            public CropRegionTester(View view)
            {
                _active = view.CropBoxActive && view.CropBox != null;
                if (!_active) return;
                BoundingBoxXYZ bb = view.CropBox;
                _trf = bb.Transform ?? Transform.Identity;
                _min = bb.Min; _max = bb.Max;
            }

            public bool ContainsPoint(XYZ modelPt)
            {
                if (!_active) return true;
                XYZ local = _trf.Inverse.OfPoint(modelPt);
                return local.X >= _min.X && local.X <= _max.X &&
                       local.Y >= _min.Y && local.Y <= _max.Y &&
                       local.Z >= _min.Z && local.Z <= _max.Z;
            }
        }
    }
}
