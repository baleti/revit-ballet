using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms; // Alias for Windows Forms

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
    public class ToggleLevelBubblesInSelectedViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
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

            // Get the active document.
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            if (uiDoc == null)
            {
                message = "No active document.";
                return Result.Failed;
            }
            Document doc = uiDoc.Document;

            // Retrieve the currently selected elements (views or viewports).
            ICollection<ElementId> selIds = uiDoc.GetSelectionIds();
            List<Autodesk.Revit.DB.View> selectedViews = new List<Autodesk.Revit.DB.View>();

            foreach (ElementId id in selIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is Autodesk.Revit.DB.View view)
                {
                    if (!selectedViews.Contains(view))
                        selectedViews.Add(view);
                }
                else if (elem is Viewport vp)
                {
                    Element viewElem = doc.GetElement(vp.ViewId);
                    if (viewElem is Autodesk.Revit.DB.View vpView && !selectedViews.Contains(vpView))
                        selectedViews.Add(vpView);
                }
            }

            if (selectedViews.Count == 0)
            {
                message = "Please select one or more views or viewports.";
                return Result.Failed;
            }

            // Collect all Level elements (levels derive from DatumPlane).
            IList<Element> levels = new FilteredElementCollector(doc)
                                        .OfClass(typeof(Level))
                                        .ToElements();

            // Start a transaction.
            using (Transaction trans = new Transaction(doc, "Hide/Show Level Bubbles"))
            {
                trans.Start();

                foreach (Autodesk.Revit.DB.View view in selectedViews)
                {
                    foreach (Element levelElem in levels)
                    {
                        DatumPlane dp = levelElem as DatumPlane;
                        if (dp != null)
                        {
                            try
                            {
                                // Process based on the chosen bubble option and operation.
                                switch (chosenBubbleOption)
                                {
                                    case BubbleOption.End0:
                                        if (chosenOperation == BubbleOperation.Hide)
                                            dp.HideBubbleInView(DatumEnds.End0, view);
                                        else
                                            dp.ShowBubbleInView(DatumEnds.End0, view);
                                        break;
                                    case BubbleOption.End1:
                                        if (chosenOperation == BubbleOperation.Hide)
                                            dp.HideBubbleInView(DatumEnds.End1, view);
                                        else
                                            dp.ShowBubbleInView(DatumEnds.End1, view);
                                        break;
                                    case BubbleOption.Both:
                                        if (chosenOperation == BubbleOperation.Hide)
                                        {
                                            dp.HideBubbleInView(DatumEnds.End0, view);
                                            dp.HideBubbleInView(DatumEnds.End1, view);
                                        }
                                        else
                                        {
                                            dp.ShowBubbleInView(DatumEnds.End0, view);
                                            dp.ShowBubbleInView(DatumEnds.End1, view);
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
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
}
