using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace YourNamespace
{
    /// <summary>
    /// Contains a helper method for duplicating a view as a drafting view.
    /// </summary>
    public static class DraftingViewDuplicator
    {
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

            // If the view is a Legend, check for existing drafting views with the same name.
            if (originalView.ViewType == ViewType.Legend)
            {
                // Collect all existing drafting view names in the document.
                var existingNames = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewDrafting))
                    .Cast<ViewDrafting>()
                    .Select(v => v.Name)
                    .ToList();
                
                string newName = baseName;
                int counter = 2;
                // Append a counter until a unique name is found.
                while (existingNames.Contains(newName))
                {
                    newName = baseName + " " + counter.ToString();
                    counter++;
                }
                newDraftingView.Name = newName;
            }
            else
            {
                // For non-legend views, simply use the base name.
                newDraftingView.Name = baseName;
            }

            // Collect annotation elements and other detail components from the original view.
            FilteredElementCollector collector = new FilteredElementCollector(doc, originalView.Id)
                .WhereElementIsNotElementType();

            // Create a list to store valid element IDs.
            ICollection<ElementId> elementIds = new List<ElementId>();

            // Process each element to determine if it can be copied to a drafting view.
            foreach (Element e in collector)
            {
                bool canCopy = false;
                
                // Check for annotation elements.
                if (e.Category != null && e.Category.CategoryType == CategoryType.Annotation)
                {
                    canCopy = true;
                }
                // Check for specific built-in categories allowed in drafting views.
                else if (e.Category != null && (
                    e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents ||
                    e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_IOSDetailGroups ||
                    e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Lines ||
                    e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_RasterImages ||
                    e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_InsulationLines
                ))
                {
                    canCopy = true;
                }
                // Handle ImportInstance elements specifically - only allow detail item imports.
                else if (e is ImportInstance importInstance)
                {
                    // Check if it's a 2D or detail import (not a 3D model import).
                    if (importInstance.ViewSpecific)
                    {
                        canCopy = true;
                    }
                }
                
                if (canCopy)
                {
                    elementIds.Add(e.Id);
                }
            }

            // Copy the filtered elements to the new drafting view.
            if (elementIds.Count > 0)
            {
                ElementTransformUtils.CopyElements(originalView, elementIds, newDraftingView, Transform.Identity, null);
            }

            // Duplicate the crop region and annotations (filled region, detail curves, etc.).
            ViewCropRegionShapeManager cropManager = originalView.GetCropRegionShapeManager();
            IList<CurveLoop> cropLoops = cropManager.GetCropShape();

            if (cropLoops != null && cropLoops.Count > 0)
            {
                // Compute the transform between the original view and the new drafting view.
                Transform viewTransform = ElementTransformUtils.GetTransformFromViewToView(originalView, newDraftingView);
                XYZ projectedOrigin = viewTransform.OfPoint(new XYZ(0, 0, 0));
                XYZ translation2D = new XYZ(projectedOrigin.X, projectedOrigin.Y, 0);

                List<CurveLoop> shiftedLoops = new List<CurveLoop>();

                foreach (CurveLoop loop in cropLoops)
                {
                    CurveLoop shiftedLoop = new CurveLoop();
                    foreach (Curve curve in loop)
                    {
                        Curve shiftedCurve = curve.CreateTransformed(Transform.CreateTranslation(translation2D));
                        shiftedLoop.Append(shiftedCurve);
                    }
                    shiftedLoops.Add(shiftedLoop);

                    // Create an outer loop (e.g., offset rectangle) if possible.
                    try
                    {
                        double offsetDistanceFeet = 500 / 304.8; // Convert 500 mm to feet.
                        XYZ min = null;
                        XYZ max = null;
                        foreach (Curve curve in shiftedLoop)
                        {
                            for (int i = 0; i < 2; i++)
                            {
                                XYZ point = curve.GetEndPoint(i);
                                if (min == null)
                                {
                                    min = point;
                                    max = point;
                                }
                                else
                                {
                                    min = new XYZ(Math.Min(min.X, point.X), Math.Min(min.Y, point.Y), Math.Min(min.Z, point.Z));
                                    max = new XYZ(Math.Max(max.X, point.X), Math.Max(max.Y, point.Y), Math.Max(max.Z, point.Z));
                                }
                            }
                        }
                        min = new XYZ(min.X - offsetDistanceFeet, min.Y - offsetDistanceFeet, min.Z);
                        max = new XYZ(max.X + offsetDistanceFeet, max.Y + offsetDistanceFeet, max.Z);
                        CurveLoop outerLoop = new CurveLoop();
                        XYZ pt1 = new XYZ(min.X, min.Y, min.Z);
                        XYZ pt2 = new XYZ(max.X, min.Y, min.Z);
                        XYZ pt3 = new XYZ(max.X, max.Y, min.Z);
                        XYZ pt4 = new XYZ(min.X, max.Y, min.Z);
                        outerLoop.Append(Line.CreateBound(pt1, pt2));
                        outerLoop.Append(Line.CreateBound(pt2, pt3));
                        outerLoop.Append(Line.CreateBound(pt3, pt4));
                        outerLoop.Append(Line.CreateBound(pt4, pt1));
                        shiftedLoops.Add(outerLoop);
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Offset Error", $"Failed to create outer loop: {ex.Message}");
                    }
                }

                // Look up the filled region type "White - Solid".
                FilledRegionType filledRegionType = new FilteredElementCollector(doc)
                    .OfClass(typeof(FilledRegionType))
                    .Cast<FilledRegionType>()
                    .FirstOrDefault(frt => frt.Name.Equals("White - Solid", StringComparison.InvariantCultureIgnoreCase)
                        || (frt.Name.Contains("White") && frt.Name.Contains("Solid")));

                if (filledRegionType == null)
                {
                    TaskDialog.Show("Warning", "Filled region type 'White - Solid' was not found. A random filled region type will be used.");
                    filledRegionType = new FilteredElementCollector(doc)
                        .OfClass(typeof(FilledRegionType))
                        .Cast<FilledRegionType>()
                        .FirstOrDefault();
                }

                // Create the filled region in the new drafting view.
                FilledRegion filledRegion = FilledRegion.Create(doc, filledRegionType.Id, newDraftingView.Id, shiftedLoops);
                Parameter commentsParam = filledRegion.LookupParameter("Comments");
                if (commentsParam != null && !commentsParam.IsReadOnly)
                {
                    commentsParam.Set("crop region");
                }

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
                            newLineStyle.LineColor = new Color(255, 255, 255); // White.
                            newLineStyle.SetLineWeight(1, GraphicsStyleType.Projection);
                            whiteLineStyle = new FilteredElementCollector(doc)
                                .OfClass(typeof(GraphicsStyle))
                                .Cast<GraphicsStyle>()
                                .FirstOrDefault(gs => gs.Name.Equals("White Line", StringComparison.InvariantCultureIgnoreCase));
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Error", $"Failed to create White Line style: {ex.Message}");
                        }
                    }
                }

                if (whiteLineStyle != null)
                {
                    // Option 1: Try setting the filled region's LineStyleId via reflection.
                    try
                    {
                        var propertyInfo = typeof(FilledRegion).GetProperty("LineStyleId");
                        if (propertyInfo != null)
                        {
                            propertyInfo.SetValue(filledRegion, whiteLineStyle.Id);
                        }
                    }
                    catch { }

                    // Option 2: Recreate the filled region and add detail curves.
                    doc.Delete(filledRegion.Id);
                    filledRegion = FilledRegion.Create(doc, filledRegionType.Id, newDraftingView.Id, shiftedLoops);
                    commentsParam = filledRegion.LookupParameter("Comments");
                    if (commentsParam != null && !commentsParam.IsReadOnly)
                    {
                        commentsParam.Set("crop region");
                    }
                    foreach (CurveLoop loop in shiftedLoops)
                    {
                        foreach (Curve curve in loop)
                        {
                            DetailCurve detailCurve = doc.Create.NewDetailCurve(newDraftingView, curve);
                            if (detailCurve != null)
                            {
                                detailCurve.LineStyle = whiteLineStyle;
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

            // Collect eligible views.
            List<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.ViewType != ViewType.DraftingView)
                .ToList();

            // Prepare data for the custom GUI.
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
            // CustomGUIs.DataGrid is assumed to be defined elsewhere.
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

                        // Use the helper to duplicate the view.
                        ViewDrafting newDraftingView = DraftingViewDuplicator.DuplicateView(doc, originalView, draftingViewType);
                        if (newDraftingView != null)
                        {
                            count++;
                        }
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("Duplicate Views As Drafting Views", $"{count} view(s) were duplicated.");
            return Result.Succeeded;
        }
    }
}
