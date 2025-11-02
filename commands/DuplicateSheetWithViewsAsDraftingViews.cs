using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace YourNamespace
{
    [Transaction(TransactionMode.Manual)]
    public class DuplicateSheetWithViewsAsDraftingViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // --- Collect only sheets (ViewSheet objects) ---
            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            foreach (ViewSheet sheet in sheets)
            {
                // Build a verbose title: "SheetNumber - SheetName"
                string title = $"{sheet.SheetNumber} - {sheet.Name}";
                string sheetFolder = "";
                Parameter folderParam = sheet.LookupParameter("Sheet Folder");
                if (folderParam != null)
                    sheetFolder = folderParam.AsString() ?? "";

                // The dictionary keys order is controlled by the propertyNames list in the GUI.
                // We'll put "Title" and "SheetFolder" first and "Id" as the last column.
                entries.Add(new Dictionary<string, object>
                {
                    { "Id", sheet.Id.IntegerValue },
                    { "Title", title },
                    { "SheetFolder", sheetFolder }
                });
            }

            // Change the order of columns: Title, SheetFolder, then Id.
            List<string> propertyNames = new List<string> { "Title", "SheetFolder", "Id" };
            // CustomGUIs.DataGrid is assumed to be defined elsewhere.
            List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);
            if (selectedEntries == null || selectedEntries.Count == 0)
            {
                TaskDialog.Show("Duplicate Sheets As Drafting Views", "No sheets were selected.");
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

            // Get existing sheet numbers to avoid duplicates.
            HashSet<string> existingSheetNumbers = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Select(s => s.SheetNumber)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            int sheetCount = 0;
            using (Transaction trans = new Transaction(doc, "Duplicate Sheets With Drafting Views"))
            {
                trans.Start();

                foreach (var entry in selectedEntries)
                {
                    if (entry.ContainsKey("Id") && int.TryParse(entry["Id"].ToString(), out int sheetIdValue))
                    {
                        ElementId sheetId = new ElementId(sheetIdValue);
                        ViewSheet originalSheet = doc.GetElement(sheetId) as ViewSheet;
                        if (originalSheet == null)
                            continue;

                        // --- Custom Sheet Duplication ---
                        // Get the title block type used on the original sheet.
                        ElementId titleBlockTypeId = GetTitleBlockTypeId(doc, originalSheet);
                        if (titleBlockTypeId == ElementId.InvalidElementId)
                        {
                            TaskDialog.Show("Error", $"Could not determine title block type for sheet {originalSheet.SheetNumber}.");
                            continue;
                        }

                        // Generate a unique sheet number.
                        string newSheetNumber = GenerateUniqueSheetNumber(originalSheet.SheetNumber, existingSheetNumbers);
                        existingSheetNumbers.Add(newSheetNumber);

                        // Create a new sheet with the same title block.
                        ViewSheet newSheet = ViewSheet.Create(doc, titleBlockTypeId);
                        newSheet.SheetNumber = newSheetNumber;
                        newSheet.Name = originalSheet.Name + " - Drafting Views";

                        // Copy parameters from the original sheet.
                        CopyParameters(originalSheet, newSheet);

                        // --- Improved Viewport Handling ---
                        // Gather viewports from the original sheet with enhanced information.
                        Dictionary<ElementId, ViewportInfo> viewportInfoDict = new Dictionary<ElementId, ViewportInfo>();
                        FilteredElementCollector viewportCollector = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .WhereElementIsNotElementType();

                        foreach (Viewport vp in viewportCollector)
                        {
                            if (vp.SheetId.Equals(originalSheet.Id))
                            {
                                View origView = doc.GetElement(vp.ViewId) as View;
                                if (origView == null) continue;

                                // Retrieve the label outline and convert it to a list of corner points.
                                List<XYZ> labelCorners = new List<XYZ>();
                                Outline labelOutline = vp.GetLabelOutline();
                                if (labelOutline != null)
                                {
                                    // Calculate the four corners of the rectangle.
                                    XYZ min = labelOutline.MinimumPoint;
                                    XYZ max = labelOutline.MaximumPoint;
                                    XYZ bottomRight = new XYZ(max.X, min.Y, min.Z);
                                    XYZ topLeft = new XYZ(min.X, max.Y, min.Z);
                                    labelCorners.Add(min);
                                    labelCorners.Add(bottomRight);
                                    labelCorners.Add(max);
                                    labelCorners.Add(topLeft);
                                }

                                viewportInfoDict[vp.ViewId] = new ViewportInfo
                                {
                                    Center = vp.GetBoxCenter(),
                                    ViewportId = vp.Id,
                                    ViewScale = origView.Scale,
                                    ViewName = origView.Name,
                                    LabelBoxXYZ = labelCorners
                                };

                                // Try to get viewport outline (assumed to represent the original crop region outline).
                                try
                                {
                                    Outline vpOutline = vp.GetBoxOutline();
                                    viewportInfoDict[vp.ViewId].OutlineMin = vpOutline.MinimumPoint;
                                    viewportInfoDict[vp.ViewId].OutlineMax = vpOutline.MaximumPoint;
                                }
                                catch (Exception)
                                {
                                    viewportInfoDict[vp.ViewId].HasValidOutline = false;
                                }
                            }
                        }

                        // For each viewport, duplicate its view as a drafting view, set its scale, and place it on the new sheet
                        // so that it is positioned based on the original crop region outline.
                        foreach (KeyValuePair<ElementId, ViewportInfo> vpInfo in viewportInfoDict)
                        {
                            View origView = doc.GetElement(vpInfo.Key) as View;
                            if (origView == null)
                                continue;

                            // Duplicate the view as a drafting view using your custom duplicator.
                            ViewDrafting newDraftingView = DraftingViewDuplicator.DuplicateView(doc, origView, draftingViewType);
                            if (newDraftingView == null)
                                continue;

                            // Account for the original view's scale.
                            newDraftingView.Scale = vpInfo.Value.ViewScale;

                            // Determine the target center based on the original crop region outline.
                            // If a valid outline is available then use its center; otherwise, fall back on the original viewport center.
                            XYZ targetCenter = vpInfo.Value.Center;
                            if (vpInfo.Value.HasValidOutline && vpInfo.Value.OutlineMin != null && vpInfo.Value.OutlineMax != null)
                            {
                                targetCenter = (vpInfo.Value.OutlineMin + vpInfo.Value.OutlineMax) * 0.5;
                            }

                            // Place the duplicated drafting view on the new sheet.
                            if (Viewport.CanAddViewToSheet(doc, newSheet.Id, newDraftingView.Id))
                            {
                                Viewport newViewport = Viewport.Create(doc, newSheet.Id, newDraftingView.Id, targetCenter);
                                try
                                {
                                    // Adjust the viewport to match the original crop region placement.
                                    newViewport.SetBoxCenter(targetCenter);
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Error setting viewport center: {ex.Message}");
                                }

                                // Optionally refine the viewport by calling the helper,
                                // which may further adjust placement and size if API methods are available.
                                AdjustViewportToMatchOriginal(newViewport, vpInfo.Value);
                            }
                            else
                            {
                                TaskDialog.Show("Warning", $"Cannot add view {newDraftingView.Name} to sheet {newSheet.SheetNumber}.");
                            }
                        }

                        // --- Duplicate Detailing Elements on the Sheet ---
                        // Collect all elements placed directly on the original sheet that are not viewports or title blocks.
                        ICollection<ElementId> detailElementIds = new FilteredElementCollector(doc, originalSheet.Id)
                            .WhereElementIsNotElementType()
                            .Where(e =>
                                !(e is Viewport) &&
                                !(e.Category != null && e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_TitleBlocks))
                            .Select(e => e.Id)
                            .ToList();

                        // Copy these detailing elements to the new sheet.
                        ElementTransformUtils.CopyElements(originalSheet, detailElementIds, newSheet, Transform.Identity, null);

                        sheetCount++;
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("Duplicate Sheets With Drafting Views",
                $"{sheetCount} sheet(s) were duplicated with replaced drafting view placements and detailing elements.");
            return Result.Succeeded;
        }

        /// <summary>
        /// Represents enhanced information about a viewport, including outline information (from the original crop region/fill region).
        /// </summary>
        private class ViewportInfo
        {
            public XYZ Center { get; set; }
            public ElementId ViewportId { get; set; }
            public int ViewScale { get; set; }
            public string ViewName { get; set; }
            public XYZ OutlineMin { get; set; }
            public XYZ OutlineMax { get; set; }
            public bool HasValidOutline { get; set; } = true;
            public IList<XYZ> LabelBoxXYZ { get; set; }
        }

        /// <summary>
        /// Adjusts the new viewport to match the original crop region placement and outline.
        /// </summary>
        private void AdjustViewportToMatchOriginal(Viewport newViewport, ViewportInfo originalInfo)
        {
            // Compute the target center based on the crop region outline.
            XYZ targetCenter = originalInfo.Center;
            if (originalInfo.HasValidOutline && originalInfo.OutlineMin != null && originalInfo.OutlineMax != null)
            {
                targetCenter = (originalInfo.OutlineMin + originalInfo.OutlineMax) * 0.5;
            }
            // Set the viewport's center to the computed target center.
            newViewport.SetBoxCenter(targetCenter);

            // If additional adjustments are needed (for example, matching width/height if supported by the API),
            // they can be added here. Some Revit versions support methods to set the box width and height.
            try
            {
                // Example (uncomment if supported):
                // double width = originalInfo.OutlineMax.X - originalInfo.OutlineMin.X;
                // double height = originalInfo.OutlineMax.Y - originalInfo.OutlineMin.Y;
                // newViewport.SetBoxWidth(width);
                // newViewport.SetBoxHeight(height);
            }
            catch (Exception)
            {
                // Fallback if these methods aren't available.
            }
        }

        /// <summary>
        /// Generates a unique sheet number based on the original sheet number and existing numbers.
        /// </summary>
        private string GenerateUniqueSheetNumber(string originalNumber, HashSet<string> existingNumbers)
        {
            string baseNumber = originalNumber;
            string suffix = "";
            int dotIndex = originalNumber.LastIndexOf('.');
            if (dotIndex > 0)
            {
                baseNumber = originalNumber.Substring(0, dotIndex + 1);
                suffix = originalNumber.Substring(dotIndex + 1);
                if (int.TryParse(suffix, out int suffixNumber))
                {
                    for (int i = 1; i <= 999; i++)
                    {
                        string candidateNumber = baseNumber + (suffixNumber + i).ToString(new string('0', suffix.Length));
                        if (!existingNumbers.Contains(candidateNumber))
                        {
                            return candidateNumber;
                        }
                    }
                }
            }
            else
            {
                for (int i = 1; i <= 999; i++)
                {
                    string candidateNumber = baseNumber + "." + i.ToString("000");
                    if (!existingNumbers.Contains(candidateNumber))
                    {
                        return candidateNumber;
                    }
                }
            }
            return originalNumber + "-" + DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        /// <summary>
        /// Retrieves the title block type ID used on a given sheet.
        /// </summary>
        private ElementId GetTitleBlockTypeId(Document doc, ViewSheet sheet)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType();

            foreach (FamilyInstance instance in collector)
            {
                if (instance.OwnerViewId.Equals(sheet.Id))
                {
                    return instance.GetTypeId();
                }
            }
            FilteredElementCollector typeCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .OfClass(typeof(FamilySymbol))
                .WhereElementIsElementType();
            return typeCollector.FirstElementId();
        }

        /// <summary>
        /// Copies modifiable parameters from the source sheet to the target sheet.
        /// </summary>
        private void CopyParameters(Element source, Element target)
        {
            foreach (Parameter param in source.Parameters)
            {
                if (!param.IsReadOnly && param.StorageType != StorageType.None && param.Definition != null)
                {
                    // Skip sheet number and name as these are already set.
                    if (param.Definition.Name.Equals("Sheet Number", StringComparison.OrdinalIgnoreCase) ||
                        param.Definition.Name.Equals("Sheet Name", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    Parameter targetParam = target.get_Parameter(param.Definition);
                    if (targetParam != null && !targetParam.IsReadOnly && param.HasValue)
                    {
                        try
                        {
                            switch (param.StorageType)
                            {
                                case StorageType.String:
                                    targetParam.Set(param.AsString());
                                    break;
                                case StorageType.Integer:
                                    targetParam.Set(param.AsInteger());
                                    break;
                                case StorageType.Double:
                                    targetParam.Set(param.AsDouble());
                                    break;
                                case StorageType.ElementId:
                                    targetParam.Set(param.AsElementId());
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Could not copy parameter {param.Definition.Name}: {ex.Message}");
                        }
                    }
                }
            }
        }
    }
}
