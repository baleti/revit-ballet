using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

[Transaction(TransactionMode.Manual)]
public class OffsetLinesFromTheirWorkplane : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;

        try
        {
            // Get selected elements
            ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                message = "Please select at least one model line.";
                return Result.Failed;
            }

            // Prompt user for offset distance in mm
            string input = OffsetWorkplaneInputDialog("Enter offset (in mm):", "Offset Lines");
            if (string.IsNullOrEmpty(input))
            {
                return Result.Cancelled;
            }

            // Attempt to parse the floating-point value
            if (!double.TryParse(input, out double offsetMm))
            {
                message = "Invalid numeric input.";
                return Result.Failed;
            }

            // Convert millimeters to feet
            double offsetFeet = offsetMm * 0.0032808399;

            using (Transaction tx = new Transaction(doc, "Offset Lines"))
            {
                tx.Start();

                foreach (ElementId elementId in selectedIds)
                {
                    Element element = doc.GetElement(elementId);
                    
                    // Check if element is a ModelCurve
                    ModelCurve modelCurve = element as ModelCurve;
                    if (modelCurve == null)
                    {
                        continue;
                    }

                    // Get the current SketchPlane
                    SketchPlane currentSketchPlane = modelCurve.SketchPlane;
                    if (currentSketchPlane == null)
                    {
                        continue;
                    }

                    // Get the current curve geometry
                    Curve curve = modelCurve.GeometryCurve;
                    if (curve == null)
                    {
                        continue;
                    }

                    // Get the plane's normal vector for offset direction
                    Plane currentPlane = currentSketchPlane.GetPlane();
                    XYZ normal = currentPlane.Normal;

                    // Create a new plane offset from the current one
                    Plane newPlane = Plane.CreateByNormalAndOrigin(
                        normal,
                        currentPlane.Origin + normal.Multiply(offsetFeet));

                    // Create a new SketchPlane at the offset location
                    SketchPlane newSketchPlane = SketchPlane.Create(doc, newPlane);

                    // Project the curve onto the new plane
                    XYZ translation = normal.Multiply(offsetFeet);
                    Curve offsetCurve = curve.CreateTransformed(Transform.CreateTranslation(translation));

                    // Create a new model curve with the same line style
                    ModelCurve newCurve = doc.Create.NewModelCurve(offsetCurve, newSketchPlane);
                    newCurve.LineStyle = modelCurve.LineStyle;
                    
                    // Delete the original curve
                    doc.Delete(elementId);
                }

                tx.Commit();
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }

    /// <summary>
    /// Shows a simple WinForm dialog box to get a user-entered string (e.g. offset distance in mm).
    /// Returns the user input as a string, or null/empty if canceled.
    /// </summary>
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
        return dialogResult == DialogResult.OK ? textBox.Text : null;
    }
}
