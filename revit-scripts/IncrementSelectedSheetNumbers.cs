using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class IncrementSheetNumbers : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Collect all the sheets in the document
        var sheets = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Sheets)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .ToList();

        // Prepare data for the custom GUI
        List<Dictionary<string, object>> sheetData = sheets.Select(sheet => new Dictionary<string, object>
        {
            { "Sheet Number", sheet.SheetNumber },
            { "Sheet Name", sheet.Name }
        }).ToList();

        List<string> propertyNames = new List<string> { "Sheet Number", "Sheet Name" };

        // Call the custom GUI to allow the user to select sheets
        var selectedSheets = CustomGUIs.DataGrid(sheetData, propertyNames, false);

        if (selectedSheets.Count == 0)
        {
            TaskDialog.Show("Info", "No sheets selected.");
            return Result.Cancelled;
        }

        // Prompt user for integer input
        int incrementValue;
        if (!PromptForIncrementValue(out incrementValue))
        {
            TaskDialog.Show("Error", "Invalid increment value.");
            return Result.Cancelled;
        }

        // Sort selected sheets based on increment value
        var sortedSelectedSheets = incrementValue < 0
            ? selectedSheets.OrderBy(sheet => int.Parse(sheet["Sheet Number"].ToString())).ToList()
            : selectedSheets.OrderByDescending(sheet => int.Parse(sheet["Sheet Number"].ToString())).ToList();

        using (Transaction trans = new Transaction(doc, "Increment Sheet Numbers"))
        {
            trans.Start();

            foreach (var selectedSheetData in sortedSelectedSheets)
            {
                string sheetNumber = selectedSheetData["Sheet Number"].ToString();
                var sheet = sheets.FirstOrDefault(s => s.SheetNumber == sheetNumber);

                if (sheet != null)
                {
                    // Increment the sheet number by the user-provided value
                    string incrementedSheetNumber = IncrementSheetNumber(sheet.SheetNumber, incrementValue);
                    sheet.SheetNumber = incrementedSheetNumber;
                }
            }

            trans.Commit();
        }

        TaskDialog.Show("Info", "Sheet numbers incremented successfully.");
        return Result.Succeeded;
    }

    private bool PromptForIncrementValue(out int incrementValue)
    {
        incrementValue = 0;
        using (System.Windows.Forms.Form inputForm = new System.Windows.Forms.Form())
        {
            inputForm.Text = "Enter Increment Value";
            Label label = new Label() { Left = 10, Top = 20, Text = "Increment Value:" };
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox() { Left = 130, Top = 20, Width = 100 };
            Button confirmation = new Button() { Text = "OK", Left = 130, Width = 100, Top = 50, DialogResult = DialogResult.OK };

            confirmation.Click += (sender, e) => { inputForm.Close(); };

            inputForm.Controls.Add(label);
            inputForm.Controls.Add(textBox);
            inputForm.Controls.Add(confirmation);
            inputForm.AcceptButton = confirmation;

            if (inputForm.ShowDialog() == DialogResult.OK)
            {
                return int.TryParse(textBox.Text, out incrementValue);
            }
        }

        return false;
    }

    private string IncrementSheetNumber(string sheetNumber, int incrementValue)
    {
        // Try to parse the sheet number as an integer
        if (int.TryParse(sheetNumber, out int number))
        {
            return (number + incrementValue).ToString();
        }

        // If parsing fails, try incrementing by detecting numeric suffix
        string numericPart = new string(sheetNumber.Where(char.IsDigit).ToArray());
        if (int.TryParse(numericPart, out int numericValue))
        {
            string nonNumericPart = new string(sheetNumber.Where(char.IsLetter).ToArray());
            return $"{nonNumericPart}{numericValue + incrementValue}";
        }

        // If no numeric part, return the same string (or handle as needed)
        return sheetNumber;
    }
}
