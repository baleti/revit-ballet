// DuplicateViews.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;       // For Windows Forms
using System.Drawing;             // For form sizing
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

// Aliases to disambiguate Revit types from Windows Forms types.
using RevitView = Autodesk.Revit.DB.View;
using RevitViewport = Autodesk.Revit.DB.Viewport;

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
        /// The new view is renamed using the original name, "Copy", timestamp, and a sequential number suffix.
        /// If a view does not support the chosen duplication option, it is skipped.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="views">The list of views to duplicate.</param>
        /// <param name="mode">The chosen duplication mode.</param>
        /// <param name="duplicateCount">The number of duplicates to create.</param>
        public static void DuplicateViews(Document doc, IEnumerable<RevitView> views, DuplicationMode mode, int duplicateCount)
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
                                dupView.Name = $"{baseName} - Copy {timestamp}";
                            }
                            else
                            {
                                // For multiple duplicates, append _N suffix starting from 1
                                dupView.Name = $"{baseName} - Copy {timestamp}_{i + 1}";
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

        private RadioButton radioWithoutDetailing;
        private RadioButton radioWithDetailing;
        private RadioButton radioDependent;
        private NumericUpDown numDuplicates;
        private Label lblDuplicates;
        private Button okButton;
        private Button cancelButton;

        public DuplicationOptionsForm()
        {
            // Set basic form properties.
            this.Text = "Duplication Options";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(250, 190);
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Create radio buttons.
            radioWithoutDetailing = new RadioButton()
            {
                Text = "Without Detailing",
                Left = 20,
                Top = 20,
                Width = 200,
                Checked = true
            };
            radioWithDetailing = new RadioButton()
            {
                Text = "With Detailing",
                Left = 20,
                Top = 50,
                Width = 200
            };
            radioDependent = new RadioButton()
            {
                Text = "Dependent",
                Left = 20,
                Top = 80,
                Width = 200
            };

            // Create number of duplicates control
            lblDuplicates = new Label()
            {
                Text = "Number of duplicates:",
                Left = 20,
                Top = 120,
                Width = 130
            };

            numDuplicates = new NumericUpDown()
            {
                Left = 155,
                Top = 118,
                Width = 70,
                Minimum = 1,
                Maximum = 100,
                Value = 1
            };

            // Create OK and Cancel buttons.
            okButton = new Button()
            {
                Text = "OK",
                Left = 50,
                Width = 70,
                Top = 150,
                DialogResult = DialogResult.OK
            };
            cancelButton = new Button()
            {
                Text = "Cancel",
                Left = 130,
                Width = 70,
                Top = 150,
                DialogResult = DialogResult.Cancel
            };

            // Add controls to the form.
            this.Controls.Add(radioWithoutDetailing);
            this.Controls.Add(radioWithDetailing);
            this.Controls.Add(radioDependent);
            this.Controls.Add(lblDuplicates);
            this.Controls.Add(numDuplicates);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);

            // Set the Accept and Cancel buttons.
            this.AcceptButton = okButton;
            this.CancelButton = cancelButton;
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);

            // When closing with OK, store the selected duplication mode and count.
            if (this.DialogResult == DialogResult.OK)
            {
                if (radioWithoutDetailing.Checked)
                    SelectedMode = DuplicationMode.WithoutDetailing;
                else if (radioWithDetailing.Checked)
                    SelectedMode = DuplicationMode.WithDetailing;
                else if (radioDependent.Checked)
                    SelectedMode = DuplicationMode.Dependent;

                DuplicateCount = (int)numDuplicates.Value;
            }
        }
    }

    /// <summary>
    /// Command that duplicates the views currently selected in the active view (or, if a viewport is selected, its associated view).
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class DuplicateSelectedViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();

            // Use the currently selected elements.
            List<RevitView> selectedViews = new List<RevitView>();

            foreach (ElementId id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element is RevitView view)
                {
                    selectedViews.Add(view);
                }
                else if (element is RevitViewport viewport)
                {
                    // If a viewport is selected, add its corresponding view.
                    RevitView viewFromViewport = doc.GetElement(viewport.ViewId) as RevitView;
                    if (viewFromViewport != null)
                    {
                        selectedViews.Add(viewFromViewport);
                    }
                }
            }

            if (selectedViews.Count == 0)
            {
                TaskDialog.Show("Error", "No valid views selected.");
                return Result.Cancelled;
            }

            // Prompt the user for duplication options.
            using (var form = new DuplicationOptionsForm())
            {
                if (form.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;

                DuplicationMode mode = form.SelectedMode;
                int duplicateCount = form.DuplicateCount;

                // Start a transaction.
                using (Transaction trans = new Transaction(doc, "Duplicate Selected Views"))
                {
                    trans.Start();
                    try
                    {
                        // Duplicate the selected views using the chosen mode and count.
                        ViewDuplicator.DuplicateViews(doc, selectedViews, mode, duplicateCount);
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
