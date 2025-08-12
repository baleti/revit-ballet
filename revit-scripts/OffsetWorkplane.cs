using System;
using System.Windows.Forms; // For the simple input dialog
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

[Transaction(TransactionMode.Manual)]
public class OffsetWorkplane : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Prompt user for offset distance in mm
        string input = OffsetWorkplaneInputDialog("Enter offset (in mm):", "Offset Current Workplane");
        if (string.IsNullOrEmpty(input))
        {
            // User canceled or empty input
            return Result.Cancelled;
        }

        // Attempt to parse the floating-point value
        if (!double.TryParse(input, out double offsetMm))
        {
            message = "Invalid numeric input.";
            return Result.Failed;
        }

        // Convert millimeters to feet:
        // 1 mm = 0.0032808399 ft  (approximately)
        double offsetFeet = offsetMm * 0.0032808399;

        // Get the active view's SketchPlane
        SketchPlane activeSketchPlane = doc.ActiveView.SketchPlane;
        if (activeSketchPlane == null)
        {
            message = "No active work plane found in the current view.";
            return Result.Failed;
        }

        // Get the underlying geometric Plane
        Plane currentPlane = activeSketchPlane.GetPlane();
        XYZ normal = currentPlane.Normal;
        XYZ origin = currentPlane.Origin;

        // Create a new Plane offset by 'offsetFeet' along the plane's normal
        XYZ newOrigin = origin + normal * offsetFeet;
        Plane newPlane = Plane.CreateByNormalAndOrigin(normal, newOrigin);

        // Create a new SketchPlane from the new Plane
        using (Transaction tx = new Transaction(doc, "Offset Workplane"))
        {
            tx.Start();
            SketchPlane offsetSketchPlane = SketchPlane.Create(doc, newPlane);

            // Set the active view's SketchPlane to the new offset plane
            doc.ActiveView.SketchPlane = offsetSketchPlane;

            tx.Commit();
        }

        return Result.Succeeded;
    }

    /// <summary>
    /// Shows a simple WinForm dialog box to get a user-entered string (e.g. offset distance in mm).
    /// Returns the user input as a string, or null/empty if canceled.
    /// </summary>
    /// <param name="prompt">Message or label to display.</param>
    /// <param name="title">Title of the dialog box.</param>
    /// <returns>User input string.</returns>
    private string OffsetWorkplaneInputDialog(string prompt, string title)
    {
        System.Windows.Forms.Form form = new System.Windows.Forms.Form();
        Label label = new Label();
        System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
        Button buttonOk = new Button();
        Button buttonCancel = new Button();

        form.Text = title;
        label.Text = prompt;
        textBox.Text = "";

        buttonOk.Text = "OK";
        buttonCancel.Text = "Cancel";

        // Set dialog controls layout
        label.SetBounds(9, 10, 372, 13);
        textBox.SetBounds(12, 36, 372, 20);
        buttonOk.SetBounds(228, 72, 75, 23);
        buttonCancel.SetBounds(309, 72, 75, 23);

        label.AutoSize = true;
        textBox.Anchor |= AnchorStyles.Right;
        buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

        buttonOk.DialogResult = DialogResult.OK;
        buttonCancel.DialogResult = DialogResult.Cancel;

        form.ClientSize = new System.Drawing.Size(396, 107);
        form.Controls.AddRange(new System.Windows.Forms.Control[] { label, textBox, buttonOk, buttonCancel });
        form.ClientSize = new System.Drawing.Size(
            Math.Max(300, label.Right + 10), 
            form.ClientSize.Height);
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;
        form.AcceptButton = buttonOk;
        form.CancelButton = buttonCancel;

        DialogResult dialogResult = form.ShowDialog();
        if (dialogResult == DialogResult.OK)
        {
            return textBox.Text;
        }
        else
        {
            return null;
        }
    }
}
