using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;
using System;

[Transaction(TransactionMode.ReadOnly)]
public class ListSheetsWithRevisions : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        try
        {
            // Step 1: Display sheets for selection
            var allSheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .OrderBy(s => s.SheetNumber);

            var sheetEntries = new List<Dictionary<string, object>>();
            
            foreach (var sheet in allSheets)
            {
                sheetEntries.Add(new Dictionary<string, object>
                {
                    { "Sheet Number", sheet.SheetNumber },
                    { "Sheet Name", sheet.Name },
                    { "Id", sheet.Id }
                });
            }

            if (sheetEntries.Count == 0)
            {
                TaskDialog.Show("Warning", "No sheets found in the project.");
                return Result.Succeeded;
            }

            // First dialog - sheet selection
            var sheetProps = new List<string> { "Sheet Number", "Sheet Name" };
            var selectedSheets = CustomGUIs.DataGrid(sheetEntries, sheetProps, false);

            if (selectedSheets == null || selectedSheets.Count == 0)
                return Result.Succeeded;

            // Step 2: Create revision data for selected sheets
            var selectedSheetNumbers = selectedSheets.Select(s => s["Sheet Number"].ToString()).ToHashSet();
            var revisionsData = new List<Dictionary<string, object>>();

            foreach (var sheetEntry in sheetEntries.Where(s => selectedSheetNumbers.Contains(s["Sheet Number"].ToString())))
            {
                var sheet = doc.GetElement(sheetEntry["Id"] as ElementId) as ViewSheet;
                if (sheet == null) continue;

                // Get all revisions for this sheet
                var sheetRevisions = sheet.GetAdditionalRevisionIds()
                    .Select(id => doc.GetElement(id) as Revision)
                    .Where(r => r != null)
                    .OrderBy(r => r.SequenceNumber)
                    .ToList();

                var entry = new Dictionary<string, object>
                {
                    { "Sheet Number", sheet.SheetNumber },
                    { "Sheet Name", sheet.Name }
                };

                // Create revision list as one string
                if (sheetRevisions.Any())
                {
                    var revisionList = sheetRevisions
                        .Select(r => $"Rev{r.SequenceNumber}: {r.IssuedTo}")
                        .ToList();
                    entry["Revisions"] = string.Join(", ", revisionList);
                }
                else
                {
                    entry["Revisions"] = "-";
                }

                revisionsData.Add(entry);
            }

            // Second dialog - show revisions
            var revisionProps = new List<string> { "Sheet Number", "Sheet Name", "Revisions" };
            CustomGUIs.DataGrid(revisionsData, revisionProps, false);

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}
