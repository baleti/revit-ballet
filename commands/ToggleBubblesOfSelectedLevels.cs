using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms; // Alias for Windows Forms

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace HideLevelBubbles
{
    // Enum to choose the operation.
    public enum BubbleOperation
    {
        Hide,
        Show
    }

    // Enum to choose which bubble end(s) to process.
    public enum BubbleOption
    {
        End0,
        End1,
        Both
    }

    // A combined dialog that allows the user to choose whether to Hide or Show bubbles
    // and which bubble end(s) to process.
    public class BubbleOperationDialog : WinForms.Form
    {
        // Properties that will hold the user selections.
        public BubbleOperation SelectedOperation { get; private set; }
        public BubbleOption SelectedBubbleOption { get; private set; }

        // Controls for the "Operation" group.
        private WinForms.GroupBox grpOperation;
        private WinForms.RadioButton rbHide;
        private WinForms.RadioButton rbShow;

        // Controls for the "Bubble Ends" group.
        private WinForms.GroupBox grpBubbleEnds;
        private WinForms.RadioButton rbEnd0;
        private WinForms.RadioButton rbEnd1;
        private WinForms.RadioButton rbBoth;

        // OK and Cancel buttons.
        private WinForms.Button btnOK;
        private WinForms.Button btnCancel;

        public BubbleOperationDialog()
        {
            // Set form properties.
            this.Text = "Hide or Show Level Bubbles";
            this.Width = 350;
            this.Height = 300;
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            this.StartPosition = WinForms.FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Operation group box.
            grpOperation = new WinForms.GroupBox
            {
                Text = "Operation",
                Left = 20,
                Top = 20,
                Width = 290,
                Height = 70
            };
            rbHide = new WinForms.RadioButton
            {
                Text = "Hide Bubbles",
                Left = 20,
                Top = 30,
                Width = 120,
                Checked = true  // Default selection.
            };
            rbShow = new WinForms.RadioButton
            {
                Text = "Show Bubbles",
                Left = 150,
                Top = 30,
                Width = 120
            };
            grpOperation.Controls.Add(rbHide);
            grpOperation.Controls.Add(rbShow);

            // Bubble Ends group box.
            grpBubbleEnds = new WinForms.GroupBox
            {
                Text = "Bubble End(s)",
                Left = 20,
                Top = 110,
                Width = 290,
                Height = 100
            };
            rbEnd0 = new WinForms.RadioButton
            {
                Text = "End0",
                Left = 20,
                Top = 30,
                Width = 80
            };
            rbEnd1 = new WinForms.RadioButton
            {
                Text = "End1",
                Left = 110,
                Top = 30,
                Width = 80
            };
            rbBoth = new WinForms.RadioButton
            {
                Text = "Both",
                Left = 200,
                Top = 30,
                Width = 80,
                Checked = true  // Default selection.
            };
            grpBubbleEnds.Controls.Add(rbEnd0);
            grpBubbleEnds.Controls.Add(rbEnd1);
            grpBubbleEnds.Controls.Add(rbBoth);

            // OK and Cancel buttons.
            btnOK = new WinForms.Button
            {
                Text = "OK",
                Left = 70,
                Width = 80,
                Top = 230,
                DialogResult = WinForms.DialogResult.OK
            };
            btnCancel = new WinForms.Button
            {
                Text = "Cancel",
                Left = 180,
                Width = 80,
                Top = 230,
                DialogResult = WinForms.DialogResult.Cancel
            };

            // Add controls to the form.
            this.Controls.Add(grpOperation);
            this.Controls.Add(grpBubbleEnds);
            this.Controls.Add(btnOK);
            this.Controls.Add(btnCancel);

            // Set Accept and Cancel buttons.
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            // Wire up the OK button click event.
            btnOK.Click += (s, e) =>
            {
                // Set the operation.
                this.SelectedOperation = rbHide.Checked ? BubbleOperation.Hide : BubbleOperation.Show;

                // Set the bubble option.
                if (rbEnd0.Checked)
                    this.SelectedBubbleOption = BubbleOption.End0;
                else if (rbEnd1.Checked)
                    this.SelectedBubbleOption = BubbleOption.End1;
                else
                    this.SelectedBubbleOption = BubbleOption.Both;

                this.DialogResult = WinForms.DialogResult.OK;
                this.Close();
            };

            btnCancel.Click += (s, e) =>
            {
                this.DialogResult = WinForms.DialogResult.Cancel;
                this.Close();
            };
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ToggleBubblesOfSelectedLevels : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the active document and view.
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            if (uiDoc == null)
            {
                message = "No active document.";
                return Result.Failed;
            }
            Document doc = uiDoc.Document;
            Autodesk.Revit.DB.View activeView = doc.ActiveView;

            // Retrieve the currently selected elements (levels).
            ICollection<ElementId> selIds = uiDoc.GetSelectionIds();
            List<Level> selectedLevels = new List<Level>();

            foreach (ElementId id in selIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is Level level)
                {
                    selectedLevels.Add(level);
                }
            }

            if (selectedLevels.Count == 0)
            {
                message = "Please select one or more level elements.";
                return Result.Failed;
            }

            // Display the dialog to capture user choices.
            BubbleOperation chosenOperation;
            BubbleOption chosenBubbleOption;
            using (BubbleOperationDialog dialog = new BubbleOperationDialog())
            {
                if (dialog.ShowDialog() != WinForms.DialogResult.OK)
                {
                    message = "Operation cancelled by the user.";
                    return Result.Cancelled;
                }
                chosenOperation = dialog.SelectedOperation;
                chosenBubbleOption = dialog.SelectedBubbleOption;
            }

            // Start a transaction.
            using (Transaction trans = new Transaction(doc, "Hide/Show Level Bubbles"))
            {
                trans.Start();

                foreach (Level level in selectedLevels)
                {
                    DatumPlane dp = level as DatumPlane;
                    if (dp != null)
                    {
                        try
                        {
                            // Process based on the chosen bubble option and operation.
                            switch (chosenBubbleOption)
                            {
                                case BubbleOption.End0:
                                    if (chosenOperation == BubbleOperation.Hide)
                                        dp.HideBubbleInView(DatumEnds.End0, activeView);
                                    else
                                        dp.ShowBubbleInView(DatumEnds.End0, activeView);
                                    break;
                                case BubbleOption.End1:
                                    if (chosenOperation == BubbleOperation.Hide)
                                        dp.HideBubbleInView(DatumEnds.End1, activeView);
                                    else
                                        dp.ShowBubbleInView(DatumEnds.End1, activeView);
                                    break;
                                case BubbleOption.Both:
                                    if (chosenOperation == BubbleOperation.Hide)
                                    {
                                        dp.HideBubbleInView(DatumEnds.End0, activeView);
                                        dp.HideBubbleInView(DatumEnds.End1, activeView);
                                    }
                                    else
                                    {
                                        dp.ShowBubbleInView(DatumEnds.End0, activeView);
                                        dp.ShowBubbleInView(DatumEnds.End1, activeView);
                                    }
                                    break;
                            }
                        }
                        catch (Autodesk.Revit.Exceptions.ArgumentException)
                        {
                            // If the datum plane is not visible in this view, ignore the error.
                        }
                    }
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
}
