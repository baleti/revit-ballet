using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class ListDetailItemsByKeynote : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Step 1: Get all unique, non-empty keynote values in all detail items
        var detailItems = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_DetailComponents)
            .WhereElementIsNotElementType()
            .ToElements();

        var keynoteDict = new Dictionary<string, List<Element>>();
        foreach (var item in detailItems)
        {
            var keynote = item.LookupParameter("Keynote")?.AsString();
            if (!string.IsNullOrEmpty(keynote))
            {
                if (!keynoteDict.ContainsKey(keynote))
                {
                    keynoteDict[keynote] = new List<Element>();
                }
                keynoteDict[keynote].Add(item);
            }
        }

        var uniqueKeynotes = keynoteDict.Keys.ToList();
        if (uniqueKeynotes.Count == 0)
        {
            TaskDialog.Show("Info", "No detail items with keynote values found.");
            return Result.Succeeded;
        }

        // Prepare data for DataGrid
        var keynoteEntries = uniqueKeynotes.Select(k => new Dictionary<string, object> { { "Keynote", k } }).ToList();
        var selectedKeynoteEntry = CustomGUIs.DataGrid(keynoteEntries, new List<string> { "Keynote" }, false).FirstOrDefault();

        if (selectedKeynoteEntry == null)
        {
            return Result.Cancelled;
        }

        string selectedKeynote = selectedKeynoteEntry["Keynote"].ToString();

        // Step 2: List all detail items with the selected keynote parameter
        var detailItemEntries = keynoteDict[selectedKeynote]
            .Select(di => new Dictionary<string, object>
            {
                { "Id", di.Id.ToString() },
                { "Name", di.Name },
                { "Category", di.Category.Name }
            }).ToList();

        var selectedDetailItemEntry = CustomGUIs.DataGrid(detailItemEntries, new List<string> { "Id", "Name", "Category" }, false).FirstOrDefault();

        if (selectedDetailItemEntry == null)
        {
            return Result.Cancelled;
        }

        ElementId selectedDetailItemId = new ElementId(int.Parse(selectedDetailItemEntry["Id"].ToString()));
        Element selectedDetailItem = doc.GetElement(selectedDetailItemId);

        // Step 3: List all sheets on which the selected detail item is placed
        var sheets = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Sheets)
            .WhereElementIsNotElementType()
            .ToElements();

        var sheetEntries = new List<Dictionary<string, object>>();
        foreach (var sheet in sheets)
        {
            var viewSheet = sheet as ViewSheet;
            if (viewSheet == null) continue;

            var instances = new FilteredElementCollector(doc, viewSheet.Id)
                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DetailComponents))
                .Where(x => x.Id == selectedDetailItem.Id)
                .ToList();

            if (instances.Count > 0)
            {
                sheetEntries.Add(new Dictionary<string, object>
                {
                    { "Sheet Number", viewSheet.SheetNumber },
                    { "Sheet Name", viewSheet.Name }
                });
            }
        }

        if (sheetEntries.Count == 0)
        {
            TaskDialog.Show("Info", "The selected detail item is not placed on any sheets.");
            return Result.Succeeded;
        }

        CustomGUIs.DataGrid(sheetEntries, new List<string> { "Sheet Number", "Sheet Name" }, false);
        return Result.Succeeded;
    }
}