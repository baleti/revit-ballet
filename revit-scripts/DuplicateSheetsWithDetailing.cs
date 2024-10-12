// couldn't get this to work - finish later
using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

[Transaction(TransactionMode.Manual)]
public class DuplicateSheets : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Start a transaction
        using (Transaction tx = new Transaction(doc, "Duplicate Sheets"))
        {
            tx.Start();

            // Collect the sheets to duplicate
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewSheet));

            var sheetsToDuplicate = collector
                .Cast<ViewSheet>()
                .Where(vs => vs.Name == "TYPICAL BUILD-UPS EXTERNAL WALL - LOCATION PLAN - LEVEL 101" || vs.Name == "TYPICAL BUILD-UPS EXTERNAL WALL - LOCATION PLAN - LEVEL 100")
                .ToList();

            foreach (var sheet in sheetsToDuplicate)
            {
                try
                {
                    // Duplicate the sheet
                    ElementId newSheetId = doc.GetElement(sheet.Duplicate(ViewDuplicateOption.Duplicate))?.Id;

                    if (newSheetId != null && newSheetId != ElementId.InvalidElementId)
                    {
                        ViewSheet newSheet = doc.GetElement(newSheetId) as ViewSheet;
                        if (newSheet != null)
                        {
                            // Increment the sheet number inline
                            if (int.TryParse(sheet.SheetNumber, out int number))
                            {
                                number++; // Increment the number
                                newSheet.SheetNumber = number.ToString();
                            }

                            // Optionally, you can also rename the sheet
                            newSheet.Name = sheet.Name + " (Copy)";
                        }
                    }
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    tx.RollBack();
                    return Result.Failed;
                }
            }

            tx.Commit();
        }

        return Result.Succeeded;
    }
}
