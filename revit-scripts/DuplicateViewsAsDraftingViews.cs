using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace YourNamespace
{
    /// <summary>
    /// Contains helper methods for duplicating a view as a drafting view.
    /// </summary>
    public static class DraftingViewDuplicator
    {
        /// <summary>
        /// Validates and fixes a CurveLoop by filtering out segments that are too short, 
        /// forcing all endpoints to lie on Z = 0, and ensuring the loop is properly closed.
        /// </summary>
        /// <param name="loop">Input CurveLoop.</param>
        /// <param name="tolerance">Minimum distance tolerance.</param>
        /// <returns>A new, validated CurveLoop.</returns>
        private static CurveLoop ValidateAndFixCurveLoop(CurveLoop loop, double tolerance = 1e-6)
        {
            // Gather vertices from the original loop.
            List<XYZ> vertices = new List<XYZ>();

            foreach (Curve curve in loop)
            {
                // Force endpoints onto Z = 0.
                XYZ start = new XYZ(curve.GetEndPoint(0).X, curve.GetEndPoint(0).Y, 0);
                XYZ end = new XYZ(curve.GetEndPoint(1).X, curve.GetEndPoint(1).Y, 0);

                // Add endpoints only if they differ by more than the tolerance.
                if (vertices.Count == 0)
                {
                    vertices.Add(start);
                    if (start.DistanceTo(end) > tolerance)
                        vertices.Add(end);
                }
                else
                {
                    XYZ last = vertices.Last();
                    if (last.DistanceTo(start) > tolerance)
                        vertices.Add(start);
                    if (vertices.Last().DistanceTo(end) > tolerance)
                        vertices.Add(end);
                }
            }

            // Remove any adjacent duplicates.
            List<XYZ> filtered = new List<XYZ>();
            foreach (XYZ pt in vertices)
            {
                if (filtered.Count == 0 || filtered.Last().DistanceTo(pt) > tolerance)
                    filtered.Add(pt);
            }
            
            // Ensure the loop is closed.
            if (filtered.Count > 0 && filtered.First().DistanceTo(filtered.Last()) > tolerance)
                filtered.Add(filtered.First());

            // Rebuild the CurveLoop only using segments that are sufficiently long.
            CurveLoop newLoop = new CurveLoop();
            for (int i = 0; i < filtered.Count - 1; i++)
            {
                XYZ p1 = filtered[i];
                XYZ p2 = filtered[i + 1];
                if (p1.DistanceTo(p2) > tolerance)
                {
                    Line seg = Line.CreateBound(p1, p2);
                    newLoop.Append(seg);
                }
            }
            return newLoop;
        }

        /// <summary>
        /// Duplicates the provided view as a drafting view, including annotations and crop regions.
        /// </summary>
        public static ViewDrafting DuplicateView(Document doc, View originalView, ViewFamilyType draftingViewType)
        {
            // Compute the view title.
            string viewTitle = $"{originalView.Name} ({originalView.ViewType.ToString()})";
            // Create a new drafting view.
            ViewDrafting newDraftingView = ViewDrafting.Create(doc, draftingViewType.Id);
            
            // Define a base name.
            string baseName = viewTitle + " - Drafting View";

            // For legend views, ensure a unique name.
            if (originalView.ViewType == ViewType.Legend)
            {
                var existingNames = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewDrafting))
                    .Cast<ViewDrafting>()
                    .Select(v => v.Name)
                    .ToList();
                string newName = baseName;
                int counter = 2;
                while (existingNames.Contains(newName))
                {
                    newName = baseName + " " + counter.ToString();
                    counter++;
                }
                newDraftingView.Name = newName;
            }
            else
            {
                newDraftingView.Name = baseName;
            }

            // Collect annotation elements and detail components from the original view.
            FilteredElementCollector collector = new FilteredElementCollector(doc, originalView.Id)
                .WhereElementIsNotElementType();
            ICollection<ElementId> elementIds = new List<ElementId>();

            foreach (Element e in collector)
            {
                bool canCopy = false;
                if (e.Category != null && e.Category.CategoryType == CategoryType.Annotation)
                {
                    canCopy = true;
                }
                else if (e.Category != null && (
                    e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents ||
                    e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_IOSDetailGroups ||
                    e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Lines ||
                    e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_RasterImages ||
                    e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_InsulationLines))
                {
                    canCopy = true;
                }
                else if (e is ImportInstance importInstance && importInstance.ViewSpecific)
                {
                    canCopy = true;
                }

                if (canCopy)
                    elementIds.Add(e.Id);
            }
            
            // Copy acceptable elements to the new drafting view.
            if (elementIds.Count > 0)
                ElementTransformUtils.CopyElements(originalView, elementIds, newDraftingView, Transform.Identity, null);

            // Process the crop region.
            ViewCropRegionShapeManager cropManager = originalView.GetCropRegionShapeManager();
            IList<CurveLoop> cropLoops = cropManager.GetCropShape();
            if (cropLoops != null && cropLoops.Count > 0)
            {
                // Get the transform from the original view to the new drafting view.
                Transform viewTransform = ElementTransformUtils.GetTransformFromViewToView(originalView, newDraftingView);
                XYZ projectedOrigin = viewTransform.OfPoint(new XYZ(0, 0, 0));
                XYZ translation2D = new XYZ(projectedOrigin.X, projectedOrigin.Y, 0);

                // Process each crop region loop.
                List<CurveLoop> shiftedLoops = new List<CurveLoop>();
                foreach (CurveLoop loop in cropLoops)
                {
                    // Translate each curve in the loop.
                    CurveLoop shiftedLoop = new CurveLoop();
                    foreach (Curve c in loop)
                    {
                        Curve cShifted = c.CreateTransformed(Transform.CreateTranslation(translation2D));
                        shiftedLoop.Append(cShifted);
                    }

                    // Validate and fix the shifted loop.
                    CurveLoop validLoop = ValidateAndFixCurveLoop(shiftedLoop);
                    shiftedLoops.Add(validLoop);
                }

                // Validate and filter out degenerate loops (i.e. those with fewer than three distinct vertices).
                List<CurveLoop> validLoops = new List<CurveLoop>();
                foreach (CurveLoop loop in shiftedLoops)
                {
                    HashSet<string> distinctPoints = new HashSet<string>();
                    foreach (Curve c in loop)
                    {
                        XYZ pt = c.GetEndPoint(0);
                        string key = $"{Math.Round(pt.X, 6)},{Math.Round(pt.Y, 6)},{Math.Round(pt.Z, 6)}";
                        distinctPoints.Add(key);
                    }
                    if (distinctPoints.Count >= 3)
                    {
                        validLoops.Add(loop);
                    }
                }

                // If no valid loops remain, skip filled region creation.
                if (validLoops.Count > 0)
                {
                    // Look up the filled region type "White - Solid".
                    FilledRegionType filledRegionType = new FilteredElementCollector(doc)
                        .OfClass(typeof(FilledRegionType))
                        .Cast<FilledRegionType>()
                        .FirstOrDefault(frt =>
                            frt.Name.Equals("White - Solid", StringComparison.InvariantCultureIgnoreCase)
                            || (frt.Name.Contains("White") && frt.Name.Contains("Solid")));

                    if (filledRegionType == null)
                    {
                        filledRegionType = new FilteredElementCollector(doc)
                            .OfClass(typeof(FilledRegionType))
                            .Cast<FilledRegionType>()
                            .FirstOrDefault();
                    }

                    // Create the filled region using the valid loops.
                    FilledRegion filledRegion = null;
                    try
                    {
                        filledRegion = FilledRegion.Create(doc, filledRegionType.Id, newDraftingView.Id, validLoops);
                    }
                    catch (Exception)
                    {
                        // If creation fails, skip without notifying the user.
                    }

                    if (filledRegion != null)
                    {
                        Parameter commentsParam = filledRegion.LookupParameter("Comments");
                        if (commentsParam != null && !commentsParam.IsReadOnly)
                            commentsParam.Set("crop region");

                        // Adjust the boundary line style using a "White Line" style.
                        GraphicsStyle whiteLineStyle = new FilteredElementCollector(doc)
                            .OfClass(typeof(GraphicsStyle))
                            .Cast<GraphicsStyle>()
                            .FirstOrDefault(gs => gs.Name.Equals("White Line", StringComparison.InvariantCultureIgnoreCase));

                        if (whiteLineStyle == null)
                        {
                            Category lineCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                            if (lineCategory != null)
                            {
                                try
                                {
                                    Category newLineStyle = doc.Settings.Categories.NewSubcategory(lineCategory, "White Line");
                                    newLineStyle.LineColor = new Color(255, 255, 255);
                                    newLineStyle.SetLineWeight(1, GraphicsStyleType.Projection);
                                    whiteLineStyle = new FilteredElementCollector(doc)
                                        .OfClass(typeof(GraphicsStyle))
                                        .Cast<GraphicsStyle>()
                                        .FirstOrDefault(gs => gs.Name.Equals("White Line", StringComparison.InvariantCultureIgnoreCase));
                                }
                                catch (Exception)
                                {
                                    // Skip silently if the White Line style cannot be created.
                                }
                            }
                        }

                        if (whiteLineStyle != null)
                        {
                            try
                            {
                                var prop = typeof(FilledRegion).GetProperty("LineStyleId");
                                if (prop != null)
                                    prop.SetValue(filledRegion, whiteLineStyle.Id);
                            }
                            catch { }

                            // Recreate the filled region and add detail curves.
                            doc.Delete(filledRegion.Id);
                            filledRegion = FilledRegion.Create(doc, filledRegionType.Id, newDraftingView.Id, validLoops);
                            commentsParam = filledRegion.LookupParameter("Comments");
                            if (commentsParam != null && !commentsParam.IsReadOnly)
                                commentsParam.Set("crop region");
                            foreach (CurveLoop loop in validLoops)
                            {
                                foreach (Curve curve in loop)
                                {
                                    DetailCurve detailCurve = doc.Create.NewDetailCurve(newDraftingView, curve);
                                    if (detailCurve != null)
                                        detailCurve.LineStyle = whiteLineStyle;
                                }
                            }
                        }
                    }
                }
            }
            
            return newDraftingView;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class DuplicateViewsAsDraftingViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Collect eligible views (non-template and non-drafting).
            List<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.ViewType != ViewType.DraftingView)
                .ToList();

            // Prepare entries for a custom UI (assumed implemented elsewhere).
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            foreach (View v in views)
            {
                string title = $"{v.Name} ({v.ViewType.ToString()})";
                ViewSheet sheetForView = sheets.FirstOrDefault(s => s.GetAllPlacedViews().Contains(v.Id));
                string sheetName = "";
                string sheetFolder = "";
                if (sheetForView != null)
                {
                    sheetName = sheetForView.Name;
                    Parameter sheetFolderParam = sheetForView.LookupParameter("Sheet Folder");
                    if (sheetFolderParam != null)
                        sheetFolder = sheetFolderParam.AsString() ?? "";
                }
                entries.Add(new Dictionary<string, object>
                {
                    { "Id", v.Id.IntegerValue },
                    { "Title", title },
                    { "Sheet", sheetName },
                    { "SheetFolder", sheetFolder }
                });
            }

            List<string> propertyNames = new List<string> { "Id", "Title", "Sheet", "SheetFolder" };
            // CustomGUIs.DataGrid is assumed to provide a selection UI.
            List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);
            if (selectedEntries == null || selectedEntries.Count == 0)
            {
                TaskDialog.Show("Duplicate Views As Drafting Views", "No views were selected.");
                return Result.Cancelled;
            }

            // Retrieve the drafting view family type.
            ViewFamilyType draftingViewType = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Drafting);
            if (draftingViewType == null)
            {
                message = "No Drafting ViewFamilyType found.";
                return Result.Failed;
            }

            int count = 0;
            using (Transaction trans = new Transaction(doc, "Duplicate Views As Drafting Views"))
            {
                trans.Start();
                foreach (var entry in selectedEntries)
                {
                    if (entry.ContainsKey("Id") && int.TryParse(entry["Id"].ToString(), out int viewIdValue))
                    {
                        ElementId viewId = new ElementId(viewIdValue);
                        View originalView = doc.GetElement(viewId) as View;
                        if (originalView == null)
                            continue;

                        // Duplicate the view using our helper.
                        ViewDrafting newDraftingView = DraftingViewDuplicator.DuplicateView(doc, originalView, draftingViewType);
                        if (newDraftingView != null)
                            count++;
                    }
                }
                trans.Commit();
            }

            TaskDialog.Show("Duplicate Views As Drafting Views", $"{count} view(s) were duplicated.");
            return Result.Succeeded;
        }
    }
}
