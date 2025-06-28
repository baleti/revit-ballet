using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.ReadOnly)]
public class DeselectRandomly : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        ICollection<ElementId> selectedElementIds = uiDoc.GetSelectionIds();

        if (selectedElementIds == null || selectedElementIds.Count == 0)
        {
            TaskDialog.Show("Error", "No elements are currently selected.");
            return Result.Cancelled;
        }

        // Prompt the user for a fraction (0.0 to 1.0)
        double fractionToKeep = PromptForFraction();
        if (fractionToKeep < 0.0 || fractionToKeep > 1.0)
        {
            TaskDialog.Show("Error", "The entered value must be between 0.0 and 1.0.");
            return Result.Cancelled;
        }

        // Randomly decide which elements to keep
        Random random = new Random();
        List<ElementId> elementsToKeep = selectedElementIds
            .OrderBy(_ => random.NextDouble()) // Shuffle the list randomly
            .Take((int)(fractionToKeep * selectedElementIds.Count)) // Take the desired fraction
            .ToList();

        // Update the selection
        uiDoc.SetSelectionIds(elementsToKeep);

        return Result.Succeeded;
    }

    private double PromptForFraction()
    {
        using (System.Windows.Forms.Form promptForm = new System.Windows.Forms.Form())
        {
            promptForm.Width = 300;
            promptForm.Height = 150;
            promptForm.Text = "Fraction to Keep";
            promptForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            promptForm.MaximizeBox = false;
            promptForm.MinimizeBox = false;
            promptForm.StartPosition = FormStartPosition.CenterScreen;

            Label textLabel = new Label() { Left = 10, Top = 20, Text = "Enter a fraction (0.0 to 1.0):" };
            System.Windows.Forms.TextBox inputBox = new System.Windows.Forms.TextBox() { Left = 10, Top = 50, Width = 260 };

            Button confirmation = new Button()
            {
                Text = "OK",
                Left = 190,
                Width = 80,
                Top = 80,
                DialogResult = DialogResult.OK
            };
            confirmation.Click += (sender, e) => { promptForm.Close(); };

            promptForm.Controls.Add(textLabel);
            promptForm.Controls.Add(inputBox);
            promptForm.Controls.Add(confirmation);
            promptForm.AcceptButton = confirmation;

            if (promptForm.ShowDialog() == DialogResult.OK)
            {
                if (double.TryParse(inputBox.Text, out double fraction))
                {
                    return fraction;
                }
            }

            return -1; // Return invalid value if input is not valid
        }
    }
}
