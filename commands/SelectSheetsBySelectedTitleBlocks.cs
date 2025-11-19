using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class SelectSheetsBySelectedTitleBlocks : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get currently selected elements
                ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();

                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("Error", "Please select at least one titleblock first.");
                    return Result.Failed;
                }

                // Create a quick filter for titleblocks
                ElementCategoryFilter titleBlockFilter = new ElementCategoryFilter(BuiltInCategory.OST_TitleBlocks);
                
                // Filter selected elements efficiently using LINQ and the category filter
                List<ElementId> titleblockIds = selectedIds
                    .Select(id => doc.GetElement(id))
                    .Where(elem => elem != null && 
                           elem is FamilyInstance && 
                           titleBlockFilter.PassesFilter(doc, elem.Id))
                    .Select(elem => elem.Id)
                    .ToList();

                if (titleblockIds.Count == 0)
                {
                    TaskDialog.Show("Error", "No titleblocks found in selection. Please select titleblocks.");
                    return Result.Failed;
                }

                // Get all sheets in one batch query
                Dictionary<string, ElementId> allSheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .ToDictionary(sheet => sheet.SheetNumber, sheet => sheet.Id);

                // Create a multi-category filter for faster element parameter access
                var parameterFilter = new ElementMulticategoryFilter(new BuiltInCategory[] { BuiltInCategory.OST_TitleBlocks });
                
                // Get sheet numbers for selected titleblocks in batch
                HashSet<ElementId> sheetsToSelect = new HashSet<ElementId>();
                
                // Process titleblocks in batches for better performance
                const int batchSize = 50;
                for (int i = 0; i < titleblockIds.Count; i += batchSize)
                {
                    var batch = titleblockIds
                        .Skip(i)
                        .Take(batchSize)
                        .Select(id => doc.GetElement(id))
                        .Where(elem => elem != null);

                    foreach (Element titleblock in batch)
                    {
                        Parameter sheetParam = titleblock.get_Parameter(BuiltInParameter.SHEET_NUMBER);
                        if (sheetParam != null)
                        {
                            string sheetNumber = sheetParam.AsString();
                            if (allSheets.TryGetValue(sheetNumber, out ElementId sheetId))
                            {
                                sheetsToSelect.Add(sheetId);
                            }
                        }
                    }
                }

                if (sheetsToSelect.Count == 0)
                {
                    TaskDialog.Show("Warning", "No sheets found for the selected titleblocks.");
                    return Result.Failed;
                }

                // Select the sheets
                uidoc.SetSelectionIds(sheetsToSelect);

                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
