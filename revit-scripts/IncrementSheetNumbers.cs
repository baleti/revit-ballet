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
            ? selectedSheets.OrderBy(sheet => ExtractNumericPart(sheet["Sheet Number"].ToString())).ToList()
            : selectedSheets.OrderByDescending(sheet => ExtractNumericPart(sheet["Sheet Number"].ToString())).ToList();

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
        // If the string is at least 4 characters, and the third character is a non-alphanumeric
        // (e.g., '.' or '-'), treat the first three characters as the prefix, and everything
        // after that as the numeric portion to be incremented.
        if (sheetNumber.Length > 3 && !char.IsLetterOrDigit(sheetNumber[2]))
        {
            // Example: "35.100" -> prefix = "35.", numericPart = "100"
            //          "35-100" -> prefix = "35-", numericPart = "100"
            string prefix = sheetNumber.Substring(0, 3);
            string numericPart = sheetNumber.Substring(3);

            if (int.TryParse(numericPart, out int numericValue))
            {
                int incrementedValue = numericValue + incrementValue;
                // Keep the same number of digits by using the length of the original numericPart
                string incrementedNumericPart = incrementedValue.ToString(
                    new string('0', numericPart.Length));
                return prefix + incrementedNumericPart;
            }
        }

        // Fallback to your original approach of:
        // 1) Extract everything up to the first digit as prefix
        // 2) The remainder is the numeric part
        // 3) Increment, then reassemble
        string oldPrefix = new string(sheetNumber.TakeWhile(ch => !char.IsDigit(ch)).ToArray());
        string oldNumericPart = sheetNumber.Substring(oldPrefix.Length);

        if (int.TryParse(oldNumericPart, out int oldNumericValue))
        {
            int incrementedOldValue = oldNumericValue + incrementValue;
            string incrementedOldNumericPart = incrementedOldValue.ToString(
                new string('0', oldNumericPart.Length));
            return oldPrefix + incrementedOldNumericPart;
        }

        // If we cannot parse, return the original value
        return sheetNumber;
    }

    private int ExtractNumericPart(string sheetNumber)
    {
        // Extract the numeric part from the sheet number
        string numericPart = new string(sheetNumber.Where(char.IsDigit).ToArray());
        return int.TryParse(numericPart, out int result) ? result : 0;
    }
}
