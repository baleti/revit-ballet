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

                        // --- Duplicate and Replace Viewports ---
                        // Gather viewports from the original sheet.
                        Dictionary<ElementId, XYZ> viewportInfo = new Dictionary<ElementId, XYZ>();
                        FilteredElementCollector viewportCollector = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .WhereElementIsNotElementType();
                        foreach (Viewport vp in viewportCollector)
                        {
                            if (vp.SheetId.Equals(originalSheet.Id))
                            {
                                viewportInfo[vp.ViewId] = vp.GetBoxCenter();
                            }
                        }

                        // For each viewport, duplicate its view as a drafting view and place it on the new sheet.
                        foreach (KeyValuePair<ElementId, XYZ> vpInfo in viewportInfo)
                        {
                            View origView = doc.GetElement(vpInfo.Key) as View;
                            if (origView == null)
                                continue;

                            // Use the helper to duplicate the view.
                            ViewDrafting newDraftingView = DraftingViewDuplicator.DuplicateView(doc, origView, draftingViewType);
                            if (newDraftingView == null)
                                continue;

                            // Place the duplicated drafting view on the new sheet at the same location.
                            if (Viewport.CanAddViewToSheet(doc, newSheet.Id, newDraftingView.Id))
                            {
                                Viewport.Create(doc, newSheet.Id, newDraftingView.Id, vpInfo.Value);
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
