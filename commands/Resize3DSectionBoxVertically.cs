using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Windows.Forms; // For the input box

[Transaction(TransactionMode.Manual)]
public class Resize3DSectionBoxVertically : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;
        var activeView = doc.ActiveView;

        // Check if the current view is a 3D view
        if (!(activeView is View3D view3D))
        {
            TaskDialog.Show("Error", "This command can only be run in a 3D view.");
            return Result.Failed;
        }

        // Check if the section box is enabled
        if (!view3D.IsSectionBoxActive)
        {
            message = "The section box is not active in this view.";
            return Result.Failed;
        }

        // Prompt user for input using Windows Forms
        string input = PromptForScalingFactor();
        if (string.IsNullOrWhiteSpace(input))
        {
            return Result.Cancelled;
        }

        // Parse the user input
        if (!double.TryParse(input, out double scaleFactor))
        {
            TaskDialog.Show("Error", "Invalid input. Please enter a valid decimal number.");
            return Result.Failed;
        }

        // Get the current section box
        BoundingBoxXYZ sectionBox = view3D.GetSectionBox();

        // Calculate the center of the section box
        XYZ center = (sectionBox.Min + sectionBox.Max) / 2;

        // Calculate the new vertical extents by scaling the section box
        double newHeight = (sectionBox.Max.Z - sectionBox.Min.Z) * scaleFactor / 2;

        // Set the new min and max points for the section box vertically
        sectionBox.Min = new XYZ(sectionBox.Min.X, sectionBox.Min.Y, center.Z - newHeight);
        sectionBox.Max = new XYZ(sectionBox.Max.X, sectionBox.Max.Y, center.Z + newHeight);

        using (Transaction tx = new Transaction(doc, "Resize 3D Section Box Vertically"))
        {
            tx.Start();
            view3D.SetSectionBox(sectionBox);
            tx.Commit();
        }

        return Result.Succeeded;
    }

    // Simple input dialog to prompt the user for the scaling factor using Windows Forms
    public static string PromptForScalingFactor()
    {
        using (System.Windows.Forms.Form inputForm = new System.Windows.Forms.Form())
        {
            inputForm.Text = "Enter Scaling Factor";
            inputForm.Width = 260;
            inputForm.Height = 135;
            inputForm.StartPosition = FormStartPosition.CenterScreen; // Center the form

            // Label for prompt (on one line)
            Label label = new Label() { Left = 10, Top = 10, Width = 230, Text = "Scaling factor (e.g. 1.5):" };
            
            // Input box
            System.Windows.Forms.TextBox inputBox = new System.Windows.Forms.TextBox() { Left = 10, Top = 35, Width = 230, Text = "1.0" };

            // OK button
            Button confirmation = new Button() { Text = "OK", Left = 170, Width = 70, Top = 65 };

            // Assign OK button as default and listen for Escape key
            confirmation.DialogResult = DialogResult.OK;
            inputForm.Controls.Add(label);
            inputForm.Controls.Add(inputBox);
            inputForm.Controls.Add(confirmation);

            inputForm.AcceptButton = confirmation;

            // Handle 'Escape' key to cancel
            inputForm.KeyPreview = true; // Allows the form to catch key presses
            inputForm.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    inputForm.DialogResult = DialogResult.Cancel; // Set DialogResult to Cancel
                    inputForm.Close(); // Close the form
                }
            };

            // Return input if confirmed, otherwise null if cancelled
            return inputForm.ShowDialog() == DialogResult.OK ? inputBox.Text : null;
        }
    }
}
