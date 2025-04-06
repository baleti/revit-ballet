using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace YourNamespace
{
    [Transaction(TransactionMode.Manual)]
    public class DuplicateViewsAsDraftingViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the active UIDocument and Document.
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Collect all eligible views: exclude templates and drafting views.
            List<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.ViewType != ViewType.DraftingView)
                .ToList();

            // Prepare dictionary entries for the custom GUI.
            // Each entry includes: Id, Title ("Name (ViewType)"), Sheet, and SheetFolder.
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            // Collect all view sheets.
            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            foreach (View v in views)
            {
                // Compute Title as "Name (ViewType)".
                string title = $"{v.Name} ({v.ViewType.ToString()})";

                // Check if the view is placed on a sheet (take the first if multiple).
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
                Dictionary<string, object> entry = new Dictionary<string, object>
                {
                    { "Id", v.Id.IntegerValue },
                    { "Title", title },
                    { "Sheet", sheetName },
                    { "SheetFolder", sheetFolder }
                };
                entries.Add(entry);
            }

            // Define the column names to display in the custom GUI.
            List<string> propertyNames = new List<string> { "Id", "Title", "Sheet", "SheetFolder" };

            // Prompt the user to select views using the custom GUI.
            List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);

            if (selectedEntries == null || selectedEntries.Count == 0)
            {
                TaskDialog.Show("Duplicate Views As Drafting Views", "No views were selected.");
                return Result.Cancelled;
            }

            // Retrieve a Drafting ViewFamilyType.
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
                    // Retrieve the view Id from the dictionary.
                    if (entry.ContainsKey("Id") && int.TryParse(entry["Id"].ToString(), out int viewIdValue))
                    {
                        ElementId viewId = new ElementId(viewIdValue);
                        View originalView = doc.GetElement(viewId) as View;
                        if (originalView == null)
                            continue;

                        // Compute the view title.
                        string viewTitle = $"{originalView.Name} ({originalView.ViewType.ToString()})";

                        // Create a new drafting view.
                        ViewDrafting newDraftingView = ViewDrafting.Create(doc, draftingViewType.Id);
                        newDraftingView.Name = viewTitle + " - Drafting View";

                        // Collect elements to copy:
                        // • All annotation elements (Text Notes, Raster Images, Detail Items, Lines, Dimensions, etc.)
                        // • All elements in the OST_DetailComponents category
                        ICollection<ElementId> elementIds = new FilteredElementCollector(doc, originalView.Id)
                            .WhereElementIsNotElementType()
                            .Where(e =>
                                (e.Category != null && e.Category.CategoryType == CategoryType.Annotation) ||
                                (e.Category != null && e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents)
                            )
                            .Select(e => e.Id)
                            .ToList();

                        // Copy the filtered elements to the new drafting view using an identity transform.
                        ElementTransformUtils.CopyElements(originalView, elementIds, newDraftingView, Transform.Identity, null);

                        // --- Create a filled region using the original view's crop region curves ---
                        // Get the crop region curves.
                        ViewCropRegionShapeManager cropManager = originalView.GetCropRegionShapeManager();
                        IList<CurveLoop> cropLoops = cropManager.GetCropShape();

                        if (cropLoops != null && cropLoops.Count > 0)
                        {
                            // Compute the transform between the original view and the new drafting view.
                            Transform viewTransform = ElementTransformUtils.GetTransformFromViewToView(originalView, newDraftingView);
                            // Determine where the world (internal) origin lands in the drafting view.
                            XYZ projectedOrigin = viewTransform.OfPoint(new XYZ(0, 0, 0));
                            // Use only X and Y components.
                            XYZ translation2D = new XYZ(projectedOrigin.X, projectedOrigin.Y, 0);

                            // Create new crop loops shifted by the 2D translation.
                            List<CurveLoop> shiftedLoops = new List<CurveLoop>();
                            
                            // Process each original crop loop
                            foreach (CurveLoop loop in cropLoops)
                            {
                                // 1. First create the shifted loop (from original crop region)
                                CurveLoop shiftedLoop = new CurveLoop();
                                foreach (Curve curve in loop)
                                {
                                    // Apply the translation to each curve.
                                    Curve shiftedCurve = curve.CreateTransformed(Transform.CreateTranslation(translation2D));
                                    shiftedLoop.Append(shiftedCurve);
                                }
                                shiftedLoops.Add(shiftedLoop);
                                
                                // 2. Create an outer loop by expanding the bounding box of the crop region
                                try
                                {
                                    // Convert 500mm to feet (Revit's internal unit)
                                    double offsetDistanceFeet = 500 / 304.8; // 1 foot = 304.8 mm
                                    
                                    // Find the bounding box of the loop
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
                                                min = new XYZ(
                                                    Math.Min(min.X, point.X),
                                                    Math.Min(min.Y, point.Y),
                                                    Math.Min(min.Z, point.Z)
                                                );
                                                
                                                max = new XYZ(
                                                    Math.Max(max.X, point.X),
                                                    Math.Max(max.Y, point.Y),
                                                    Math.Max(max.Z, point.Z)
                                                );
                                            }
                                        }
                                    }
                                    
                                    // Expand the bounding box by the offset distance
                                    min = new XYZ(min.X - offsetDistanceFeet, min.Y - offsetDistanceFeet, min.Z);
                                    max = new XYZ(max.X + offsetDistanceFeet, max.Y + offsetDistanceFeet, max.Z);
                                    
                                    // Create a rectangle from the expanded bounding box
                                    CurveLoop outerLoop = new CurveLoop();
                                    
                                    // Create the four lines of the rectangle (clockwise)
                                    XYZ pt1 = new XYZ(min.X, min.Y, min.Z);
                                    XYZ pt2 = new XYZ(max.X, min.Y, min.Z);
                                    XYZ pt3 = new XYZ(max.X, max.Y, min.Z);
                                    XYZ pt4 = new XYZ(min.X, max.Y, min.Z);
                                    
                                    outerLoop.Append(Line.CreateBound(pt1, pt2));
                                    outerLoop.Append(Line.CreateBound(pt2, pt3));
                                    outerLoop.Append(Line.CreateBound(pt3, pt4));
                                    outerLoop.Append(Line.CreateBound(pt4, pt1));
                                    
                                    // Add the outer loop to our collection
                                    shiftedLoops.Add(outerLoop);
                                }
                                catch (Exception ex)
                                {
                                    // If creation fails, log the error but continue with just the original loop
                                    TaskDialog.Show("Offset Error", $"Failed to create outer loop: {ex.Message}");
                                }
                            }

                            // Look for a filled region type named "Solid - White" (fallback to any type if not found).
                            FilledRegionType filledRegionType = new FilteredElementCollector(doc)
                                .OfClass(typeof(FilledRegionType))
                                .Cast<FilledRegionType>()
                                .FirstOrDefault(frt => frt.Name.Equals("Solid - White", StringComparison.InvariantCultureIgnoreCase));
                            if (filledRegionType == null)
                            {
                                filledRegionType = new FilteredElementCollector(doc)
                                    .OfClass(typeof(FilledRegionType))
                                    .Cast<FilledRegionType>()
                                    .FirstOrDefault();
                            }

                            // Create the filled region in the new drafting view using the shifted crop loops.
                            FilledRegion filledRegion = FilledRegion.Create(doc, filledRegionType.Id, newDraftingView.Id, shiftedLoops);

                            // --- Create a new approach to modify the filled region's boundary line style ---
                            if (filledRegion != null)
                            {
                                // Find the white line style
                                GraphicsStyle whiteLineStyle = new FilteredElementCollector(doc)
                                    .OfClass(typeof(GraphicsStyle))
                                    .Cast<GraphicsStyle>()
                                    .FirstOrDefault(gs => gs.Name.Equals("White Line", StringComparison.InvariantCultureIgnoreCase));

                                // If White Line style doesn't exist, create it
                                if (whiteLineStyle == null)
                                {
                                    // Get the line category
                                    Category lineCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                                    
                                    if (lineCategory != null)
                                    {
                                        // Create a new line style subcategory named "White Line"
                                        Category newLineStyle = null;
                                        try
                                        {
                                            newLineStyle = doc.Settings.Categories.NewSubcategory(lineCategory, "White Line");
                                            
                                            // Set the line color to white
                                            newLineStyle.LineColor = new Color(255, 255, 255); // RGB for white
                                            
                                            // Set line weight (1 = thin line)
                                            newLineStyle.SetLineWeight(1, GraphicsStyleType.Projection);
                                            
                                            // Get the graphics style for the new subcategory
                                            whiteLineStyle = new FilteredElementCollector(doc)
                                                .OfClass(typeof(GraphicsStyle))
                                                .Cast<GraphicsStyle>()
                                                .FirstOrDefault(gs => gs.Name.Equals("White Line", StringComparison.InvariantCultureIgnoreCase));
                                        }
                                        catch (Exception ex)
                                        {
                                            // If creation fails, log the error and try to use an existing style
                                            TaskDialog.Show("Error", $"Failed to create White Line style: {ex.Message}");
                                            
                                            // If we couldn't create a style, try to find any line style
                                            var allStyles = new FilteredElementCollector(doc)
                                                .OfClass(typeof(GraphicsStyle))
                                                .Cast<GraphicsStyle>()
                                                .Where(gs => gs.GraphicsStyleCategory != null && 
                                                            gs.GraphicsStyleCategory.Parent != null && 
                                                            gs.GraphicsStyleCategory.Parent.Name.Equals("Lines", StringComparison.InvariantCultureIgnoreCase))
                                                .ToList();
                                                
                                            if (allStyles.Count > 0)
                                            {
                                                whiteLineStyle = allStyles.FirstOrDefault();
                                                // Log available styles for debugging
                                                string availableStyles = string.Join(", ", allStyles.Select(s => s.Name));
                                                TaskDialog.Show("Available Line Styles", $"Using {whiteLineStyle.Name} instead. Available styles: {availableStyles}");
                                            }
                                        }
                                    }
                                }

                                if (whiteLineStyle != null)
                                {
                                    // Method 1: Try setting the FilledRegion's LineStyleId directly if it exists
                                    try
                                    {
                                        // Use reflection to access the property if it exists
                                        var propertyInfo = typeof(FilledRegion).GetProperty("LineStyleId");
                                        if (propertyInfo != null)
                                        {
                                            propertyInfo.SetValue(filledRegion, whiteLineStyle.Id);
                                        }
                                    }
                                    catch
                                    {
                                        // Property not available, continue to next method
                                    }
                                    
                                    // Method 2: Create DetailCurves with the desired line style around the filled region
                                    // First, delete the original filled region and recreate it
                                    doc.Delete(filledRegion.Id);
                                    
                                    // Recreate the filled region
                                    filledRegion = FilledRegion.Create(doc, filledRegionType.Id, newDraftingView.Id, shiftedLoops);
                                    
                                    // Now create DetailCurves with white line style around the boundary
                                    foreach (CurveLoop loop in shiftedLoops)
                                    {
                                        foreach (Curve curve in loop)
                                        {
                                            DetailCurve detailCurve = doc.Create.NewDetailCurve(newDraftingView, curve);
                                            if (detailCurve != null)
                                            {
                                                // Set the line style of the detail curve
                                                detailCurve.LineStyle = whiteLineStyle;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        // --- End of filled region creation ---

                        count++;
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("Duplicate Views As Drafting Views", $"{count} drafting view(s) created successfully.");
            return Result.Succeeded;
        }
    }
}
