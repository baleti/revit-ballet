using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ListSheetsByDetailItemSelected : IExternalCommand
{
    private HashSet<string> uniqueSheetNames;

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
        uniqueSheetNames = new HashSet<string>();
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        ElementId selectedElementId = uidoc.Selection.GetElementIds().FirstOrDefault();
        string currentSheetInfo = string.Empty;
        ElementId activeSheetId = null;

        // Get the currently active sheet if any
        var activeView = uidoc.ActiveView;
        if (activeView is ViewSheet activeSheet)
        {
            currentSheetInfo = activeSheet.SheetNumber + ": " + activeSheet.Name;
            activeSheetId = activeSheet.Id;
        }
        else
        {
            // Check if the active view is placed on a sheet
            FilteredElementCollector sheetCollector = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet));
            foreach (ViewSheet sheet in sheetCollector)
            {
                if (sheet.GetAllPlacedViews().Contains(activeView.Id))
                {
                    currentSheetInfo = sheet.SheetNumber + ": " + sheet.Name;
                    activeSheetId = sheet.Id;
                    break;
                }
            }
        }

        if (selectedElementId != null)
        {
            Element selectedElement = doc.GetElement(selectedElementId);
            ElementId familyId = selectedElement.GetTypeId();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            var instances = collector
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Where(x => x.GetTypeId() == familyId)
                .ToList();

            Dictionary<ElementId, ViewSheet> sheetDict = new Dictionary<ElementId, ViewSheet>();
            FilteredElementCollector allSheetCollector = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet));
            foreach (ViewSheet sheet in allSheetCollector)
            {
                foreach (ElementId viewId in sheet.GetAllPlacedViews())
                {
                    if (!sheetDict.ContainsKey(viewId))
                    {
                        sheetDict.Add(viewId, sheet);
                    }
                }
            }

            List<ViewSheet> sheetsWithInstances = new List<ViewSheet>();
            foreach (var instance in instances)
            {
                ElementId ownerViewId = instance.OwnerViewId;
                if (ownerViewId != ElementId.InvalidElementId && sheetDict.ContainsKey(ownerViewId))
                {
                    ViewSheet sheet = sheetDict[ownerViewId];
                    if (!uniqueSheetNames.Contains(sheet.SheetNumber))
                    {
                        sheetsWithInstances.Add(sheet);
                        uniqueSheetNames.Add(sheet.SheetNumber);
                    }
                }
            }

            List<SheetEntry> sheetEntries = sheetsWithInstances
                .Select(sheet => new SheetEntry { SheetNumber = sheet.SheetNumber, SheetName = sheet.Name, SheetId = sheet.Id })
                .ToList();

            List<string> propertyNames = new List<string> { "SheetNumber", "SheetName" };

            // Determine the index of the initially highlighted entry
            List<int> initialSelectionIndices = null;
            if (activeSheetId != null)
            {
                initialSelectionIndices = new List<int>();
                int index = sheetEntries.FindIndex(s => s.SheetId == activeSheetId);
                if (index != -1)
                {
                    initialSelectionIndices.Add(index);
                }
            }

            List<SheetEntry> selectedSheets = CustomGUIs.DataGrid(sheetEntries, propertyNames, initialSelectionIndices, "Select Sheet");

            if (selectedSheets != null && selectedSheets.Count > 0)
            {
                foreach (var selectedSheet in selectedSheets)
                {
                    OpenSelectedSheet(uidoc, doc, selectedSheet.SheetNumber + ": " + selectedSheet.SheetName);
                }
            }
            else
            {
                TaskDialog.Show("Info", "No sheet selected.");
            }
        }
        else
        {
            TaskDialog.Show("Error", "No element selected.");
        }

        return Result.Succeeded;
    }

    private void OpenSelectedSheet(UIDocument uidoc, Document doc, string selectedSheetInfo)
    {
        string sheetNumber = selectedSheetInfo.Split(':')[0];

        FilteredElementCollector collector = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet));
        foreach (ViewSheet sheet in collector)
        {
            if (sheet.SheetNumber == sheetNumber)
            {
                try
                {
                    uidoc.ActiveView = sheet;
                    return;
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    TaskDialog.Show("Error", "Failed to open sheet. Sheet may be invalid.");
                    return;
                }
            }
        }
        TaskDialog.Show("Error", "Sheet '" + sheetNumber + "' not found.");
    }
}

public class SheetEntry
{
    public string SheetNumber { get; set; }
    public string SheetName { get; set; }
    public ElementId SheetId { get; set; }
}
