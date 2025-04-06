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
                            // Get a FilledRegionType (assumes at least one exists in the document).
                            FilledRegionType filledRegionType = new FilteredElementCollector(doc)
                                .OfClass(typeof(FilledRegionType))
                                .Cast<FilledRegionType>()
                                .FirstOrDefault();

                            if (filledRegionType != null)
                            {
                                // Compute the transform between the original view and the new drafting view.
                                Transform viewTransform = ElementTransformUtils.GetTransformFromViewToView(originalView, newDraftingView);
                                // Determine where the world (internal) origin lands in the drafting view.
                                XYZ projectedOrigin = viewTransform.OfPoint(new XYZ(0, 0, 0));
                                // Use only X and Y components.
                                XYZ translation2D = new XYZ(projectedOrigin.X, projectedOrigin.Y, 0);

                                // Create new crop loops shifted by the 2D translation.
                                List<CurveLoop> shiftedLoops = new List<CurveLoop>();
                                foreach (CurveLoop loop in cropLoops)
                                {
                                    CurveLoop shiftedLoop = new CurveLoop();
                                    foreach (Curve curve in loop)
                                    {
                                        // Apply the translation to each curve.
                                        Curve shiftedCurve = curve.CreateTransformed(Transform.CreateTranslation(translation2D));
                                        shiftedLoop.Append(shiftedCurve);
                                    }
                                    shiftedLoops.Add(shiftedLoop);
                                }

                                // Create the filled region in the new drafting view using the shifted crop loops.
                                FilledRegion.Create(doc, filledRegionType.Id, newDraftingView.Id, shiftedLoops);
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
