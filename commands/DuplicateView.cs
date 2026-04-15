// DuplicateView.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;       // For Windows Forms
using WinFormsTextBox = System.Windows.Forms.TextBox;
using WinFormsControl = System.Windows.Forms.Control;
using System.Drawing;             // For form sizing
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

// Aliases to disambiguate Revit types from Windows Forms types.
using RevitView = Autodesk.Revit.DB.View;
using RevitViewport = Autodesk.Revit.DB.Viewport;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace RevitAddin
{
    /// <summary>
    /// Represents the three duplication modes.
    /// </summary>
    public enum DuplicationMode
    {
        WithoutDetailing,
        WithDetailing,
        Dependent
    }

    /// <summary>
    /// Contains the common duplication logic.
    /// </summary>
    internal static class ViewDuplicator
    {
        /// <summary>
        /// Loops through the provided views and duplicates each.
        /// The new view is renamed using the optional prefix/suffix around the original name, plus "Copy" and timestamp.
        /// If a view does not support the chosen duplication option, it is skipped.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="views">The list of views to duplicate.</param>
        /// <param name="mode">The chosen duplication mode.</param>
        /// <param name="duplicateCount">The number of duplicates to create.</param>
        /// <param name="prefix">Optional prefix prepended to the view name.</param>
        /// <param name="suffix">Optional suffix appended to the view name.</param>
        public static void DuplicateViews(Document doc, IEnumerable<RevitView> views, DuplicationMode mode, int duplicateCount, string prefix = "", string suffix = "")
        {
            // Prepare a timestamp for naming.
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");

            foreach (var view in views)
            {
                // Process all views (ViewType restriction removed)
                if (view != null)
                {
                    ViewDuplicateOption duplicateOption;

                    // Determine the duplication option based on the selected mode.
                    switch (mode)
                    {
                        case DuplicationMode.Dependent:
                            duplicateOption = ViewDuplicateOption.AsDependent;
                            break;
                        case DuplicationMode.WithDetailing:
                            duplicateOption = ViewDuplicateOption.WithDetailing;
                            break;
                        default:
                            duplicateOption = ViewDuplicateOption.Duplicate;
                            break;
                    }

                    // Check if the view supports duplication with the selected option.
                    if (!view.CanViewBeDuplicated(duplicateOption))
                    {
                        // Option not supported; skip this view.
                        continue;
                    }

                    // Create the specified number of duplicates
                    for (int i = 0; i < duplicateCount; i++)
                    {
                        // Duplicate the view using the chosen option.
                        ElementId newViewId = view.Duplicate(duplicateOption);

                        // Rename the new view
                        RevitView dupView = doc.GetElement(newViewId) as RevitView;
                        if (dupView != null)
                        {
                            // Sanitize the original name by trimming curly braces
                            string baseName = view.Name.Trim(new char[] { '{', '}' });

                            // If only one duplicate is requested, don't append the _N suffix
                            if (duplicateCount == 1)
                            {
                                dupView.Name = $"{prefix}{baseName} - Copy {timestamp}{suffix}";
                            }
                            else
                            {
                                // For multiple duplicates, append _N suffix starting from 1
                                dupView.Name = $"{prefix}{baseName} - Copy {timestamp}_{i + 1}{suffix}";
                            }
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// A Windows Form that prompts the user to select a duplication option and number of duplicates.
    /// </summary>
    internal class DuplicationOptionsForm : System.Windows.Forms.Form
    {
        public DuplicationMode SelectedMode { get; private set; } = DuplicationMode.WithoutDetailing;
        public int DuplicateCount { get; private set; } = 1;
        public string Prefix { get; private set; } = "";
        public string Suffix { get; private set; } = "";

        private RadioButton radioWithoutDetailing;
        private RadioButton radioWithDetailing;
        private RadioButton radioDependent;
        private WinFormsTextBox txtPrefix;
        private WinFormsTextBox txtSuffix;
        private NumericUpDown numDuplicates;
        private Label lblDuplicates;
        private Button okButton;
        private Button cancelButton;

        public DuplicationOptionsForm()
        {
            this.Text = "Duplication Options";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(280, 265);
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            radioWithoutDetailing = new RadioButton()
            {
                Text = "Without Detailing",
                Left = 20, Top = 20, Width = 230, Checked = true
            };
            radioWithDetailing = new RadioButton()
            {
                Text = "With Detailing",
                Left = 20, Top = 50, Width = 230
            };
            radioDependent = new RadioButton()
            {
                Text = "Dependent",
                Left = 20, Top = 80, Width = 230
            };

            var lblPrefix = new Label()
            {
                Text = "Prefix:", Left = 20, Top = 120, Width = 55,
                TextAlign = ContentAlignment.MiddleLeft
            };
            txtPrefix = new WinFormsTextBox() { Left = 80, Top = 117, Width = 180 };

            var lblSuffix = new Label()
            {
                Text = "Suffix:", Left = 20, Top = 152, Width = 55,
                TextAlign = ContentAlignment.MiddleLeft
            };
            txtSuffix = new WinFormsTextBox() { Left = 80, Top = 149, Width = 180 };

            lblDuplicates = new Label()
            {
                Text = "Number of duplicates:",
                Left = 20, Top = 190, Width = 140
            };
            numDuplicates = new NumericUpDown()
            {
                Left = 165, Top = 188, Width = 70,
                Minimum = 1, Maximum = 100, Value = 1
            };

            okButton = new Button()
            {
                Text = "OK", Left = 60, Width = 70, Top = 225,
                DialogResult = DialogResult.OK
            };
            cancelButton = new Button()
            {
                Text = "Cancel", Left = 145, Width = 70, Top = 225,
                DialogResult = DialogResult.Cancel
            };

            this.Controls.AddRange(new WinFormsControl[]
            {
                radioWithoutDetailing, radioWithDetailing, radioDependent,
                lblPrefix, txtPrefix, lblSuffix, txtSuffix,
                lblDuplicates, numDuplicates, okButton, cancelButton
            });

            this.AcceptButton = okButton;
            this.CancelButton = cancelButton;
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);

            if (this.DialogResult == DialogResult.OK)
            {
                if (radioWithoutDetailing.Checked)
                    SelectedMode = DuplicationMode.WithoutDetailing;
                else if (radioWithDetailing.Checked)
                    SelectedMode = DuplicationMode.WithDetailing;
                else if (radioDependent.Checked)
                    SelectedMode = DuplicationMode.Dependent;

                DuplicateCount = (int)numDuplicates.Value;
                Prefix = txtPrefix.Text;
                Suffix = txtSuffix.Text;
            }
        }
    }

    /// <summary>
    /// Command that duplicates views from selection or active view.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class DuplicateView : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Use InputResolver to get views from selection or active view
            List<RevitView> selectedViews = InputResolver.ResolveViews(uiDoc)
                .Cast<RevitView>()
                .ToList();

            if (selectedViews.Count == 0)
            {
                TaskDialog.Show("Error", "No valid views available.");
                return Result.Cancelled;
            }

            // Prompt the user for duplication options.
            using (var form = new DuplicationOptionsForm())
            {
                if (form.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;

                DuplicationMode mode = form.SelectedMode;
                int duplicateCount = form.DuplicateCount;
                string prefix = form.Prefix;
                string suffix = form.Suffix;

                // Start a transaction.
                using (Transaction trans = new Transaction(doc, "Duplicate View"))
                {
                    trans.Start();
                    try
                    {
                        // Duplicate the selected views using the chosen mode and count.
                        ViewDuplicator.DuplicateViews(doc, selectedViews, mode, duplicateCount, prefix, suffix);
                        trans.Commit();
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        message = ex.Message;
                        return Result.Failed;
                    }
                }
            }

            return Result.Succeeded;
        }
    }
}
