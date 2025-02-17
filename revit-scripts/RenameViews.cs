using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using WinForms = System.Windows.Forms;
using RevitDB = Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace MyProject.Commands
{
    // Helper class used by the RenameViews command for the DataGrid selection.
    public class ViewInfo
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public RevitDB.View RevitView { get; set; }
    }

    /// <summary>
    /// Original command: prompts the user (via a custom DataGrid) to select views to rename.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RenameViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, RevitDB.ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            WinForms.DialogResult dr;

            UIDocument uidoc = uiApp.ActiveUIDocument;
            RevitDB.Document doc = uidoc.Document;

            // 1) Gather all non-template views.
            List<RevitDB.View> allViews = new RevitDB.FilteredElementCollector(doc)
                .OfClass(typeof(RevitDB.View))
                .Cast<RevitDB.View>()
                .Where(v => !v.IsTemplate)
                .OrderBy(v => v.Name)
                .ToList();

            // Convert views to ViewInfo objects.
            List<ViewInfo> viewInfoList = allViews.Select(v => new ViewInfo
            {
                Name = v.Name,
                Title = v.Title,
                RevitView = v
            }).ToList();

            // 2) Let the user pick which views to rename using a custom DataGrid.
            // (Ensure that CustomGUIs.DataGrid exists in your project.)
            List<ViewInfo> selectedViewInfos = CustomGUIs.DataGrid(
                viewInfoList,
                new List<string> { "Title" },
                null,
                "Select Views to Rename"
            );

            if (selectedViewInfos == null || selectedViewInfos.Count == 0)
            {
                // User didn't pick anything.
                return Result.Succeeded;
            }

            // Convert back to Revit view objects.
            List<RevitDB.View> selectedViews = selectedViewInfos.Select(vi => vi.RevitView).ToList();

            // 3) Show the interactive Find/Replace dialog.
            using (InteractiveFindReplaceForm renameForm = new InteractiveFindReplaceForm(selectedViews))
            {
                if (renameForm.ShowDialog() != WinForms.DialogResult.OK)
                {
                    // User cancelled.
                    return Result.Succeeded;
                }

                string findStr = renameForm.FindText;
                string replaceStr = renameForm.ReplaceText;

                // 4) Rename the selected views in a transaction.
                using (RevitDB.Transaction tx = new RevitDB.Transaction(doc, "Rename Views"))
                {
                    tx.Start();
                    foreach (RevitDB.View v in selectedViews)
                    {
                        try
                        {
                            string newName = v.Name.Replace(findStr, replaceStr);
                            if (!newName.Equals(v.Name, StringComparison.OrdinalIgnoreCase))
                            {
                                v.Name = newName;
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Error", $"Could not rename view \"{v.Name}\".\n{ex.Message}");
                        }
                    }
                    tx.Commit();
                }
            }

            return Result.Succeeded;
        }
    }

    /// <summary>
    /// New command: Uses the currently selected views (or viewports) to rename.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RenameSelectedViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, RevitDB.ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uidoc = uiApp.ActiveUIDocument;
            RevitDB.Document doc = uidoc.Document;

            ICollection<RevitDB.ElementId> selectedIds = uidoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("Error", "Please select at least one view or viewport.");
                return Result.Failed;
            }

            List<RevitDB.View> selectedViews = new List<RevitDB.View>();
            HashSet<RevitDB.ElementId> processedViewIds = new HashSet<RevitDB.ElementId>();

            foreach (RevitDB.ElementId id in selectedIds)
            {
                RevitDB.Element elem = doc.GetElement(id);
                if (elem is RevitDB.View view && !view.IsTemplate)
                {
                    if (!processedViewIds.Contains(view.Id))
                    {
                        selectedViews.Add(view);
                        processedViewIds.Add(view.Id);
                    }
                }
                else if (elem is RevitDB.Viewport viewport)
                {
                    RevitDB.ElementId viewId = viewport.ViewId;
                    if (viewId != null)
                    {
                        RevitDB.View associatedView = doc.GetElement(viewId) as RevitDB.View;
                        if (associatedView != null && !associatedView.IsTemplate && !processedViewIds.Contains(associatedView.Id))
                        {
                            selectedViews.Add(associatedView);
                            processedViewIds.Add(associatedView.Id);
                        }
                    }
                }
            }

            if (selectedViews.Count == 0)
            {
                TaskDialog.Show("Error", "No valid views were selected.");
                return Result.Failed;
            }

            using (InteractiveFindReplaceForm renameForm = new InteractiveFindReplaceForm(selectedViews))
            {
                if (renameForm.ShowDialog() != WinForms.DialogResult.OK)
                {
                    return Result.Succeeded;
                }

                string findStr = renameForm.FindText;
                string replaceStr = renameForm.ReplaceText;

                using (RevitDB.Transaction tx = new RevitDB.Transaction(doc, "Rename Selected Views"))
                {
                    tx.Start();
                    foreach (RevitDB.View v in selectedViews)
                    {
                        try
                        {
                            string newName = v.Name.Replace(findStr, replaceStr);
                            if (!newName.Equals(v.Name, StringComparison.OrdinalIgnoreCase))
                            {
                                v.Name = newName;
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Error", $"Could not rename view \"{v.Name}\".\n{ex.Message}");
                        }
                    }
                    tx.Commit();
                }
            }

            return Result.Succeeded;
        }
    }

    /// <summary>
    /// Shared interactive Find/Replace form that displays:
    ///   - "Find" and "Replace" text boxes.
    ///   - Two multi-line text boxes showing "Before" (original view names) and "After" (preview).
    ///   - An OK button.
    /// Pressing Escape cancels the form.
    /// </summary>
    public class InteractiveFindReplaceForm : WinForms.Form
    {
        private WinForms.Label _lblFind;
        private WinForms.TextBox _txtFind;
        private WinForms.Label _lblReplace;
        private WinForms.TextBox _txtReplace;
        private WinForms.Label _lblBefore;
        private WinForms.Label _lblAfter;
        private WinForms.RichTextBox _rtbBefore;
        private WinForms.RichTextBox _rtbAfter;
        private WinForms.Button _okButton;

        private readonly List<RevitDB.View> _selectedViews;

        public string FindText { get; private set; }
        public string ReplaceText { get; private set; }

        public InteractiveFindReplaceForm(List<RevitDB.View> selectedViews)
        {
            _selectedViews = selectedViews;

            // Setup form properties.
            this.FormBorderStyle = WinForms.FormBorderStyle.Sizable;
            this.StartPosition = WinForms.FormStartPosition.CenterScreen;
            this.Text = "Find / Replace (Preview)";
            this.Size = new Size(700, 600);
            this.MinimumSize = new Size(500, 400);
            this.KeyPreview = true;
            this.KeyDown += (s, e) =>
            {
                if (e.KeyCode == WinForms.Keys.Escape)
                {
                    this.DialogResult = WinForms.DialogResult.Cancel;
                    this.Close();
                }
            };

            // Create controls.
            _lblFind = new WinForms.Label { Text = "Find:" };
            _txtFind = new WinForms.TextBox();
            _lblReplace = new WinForms.Label { Text = "Replace:" };
            _txtReplace = new WinForms.TextBox();
            _lblBefore = new WinForms.Label { Text = "Before (Original):" };
            _lblAfter = new WinForms.Label { Text = "After (Preview):" };

            _rtbBefore = new WinForms.RichTextBox
            {
                ReadOnly = true,
                BackColor = SystemColors.Window,
                WordWrap = false,
                ScrollBars = WinForms.RichTextBoxScrollBars.Both
            };
            _rtbAfter = new WinForms.RichTextBox
            {
                ReadOnly = true,
                BackColor = SystemColors.Window,
                WordWrap = false,
                ScrollBars = WinForms.RichTextBoxScrollBars.Both
            };

            _okButton = new WinForms.Button
            {
                Text = "OK",
                DialogResult = WinForms.DialogResult.OK
            };
            this.AcceptButton = _okButton;
            _okButton.Click += (s, e) =>
            {
                this.FindText = _txtFind.Text;
                this.ReplaceText = _txtReplace.Text;
                this.DialogResult = WinForms.DialogResult.OK;
                this.Close();
            };

            // Add controls to the form.
            this.Controls.Add(_lblFind);
            this.Controls.Add(_txtFind);
            this.Controls.Add(_lblReplace);
            this.Controls.Add(_txtReplace);
            this.Controls.Add(_lblBefore);
            this.Controls.Add(_rtbBefore);
            this.Controls.Add(_lblAfter);
            this.Controls.Add(_rtbAfter);
            this.Controls.Add(_okButton);

            // Layout controls on resize.
            this.Resize += (s, e) => LayoutControls();

            // Populate initial text.
            foreach (var view in _selectedViews)
            {
                _rtbBefore.AppendText(view.Name + Environment.NewLine);
                _rtbAfter.AppendText(view.Name + Environment.NewLine);
            }

            // Update preview as the user types.
            _txtFind.TextChanged += (s, e) => UpdatePreview();
            _txtReplace.TextChanged += (s, e) => UpdatePreview();

            LayoutControls();
        }

        private void LayoutControls()
        {
            int margin = 8;
            int topY = margin;
            int labelWidth = 50;
            int controlHeight = 24;

            _lblFind.SetBounds(margin, topY + 4, labelWidth, controlHeight);
            _txtFind.SetBounds(_lblFind.Right + 5, topY, 150, controlHeight);
            _lblReplace.SetBounds(_txtFind.Right + 20, topY + 4, labelWidth, controlHeight);
            _txtReplace.SetBounds(_lblReplace.Right + 5, topY, 150, controlHeight);

            int buttonWidth = 60;
            _okButton.SetBounds(this.ClientSize.Width - buttonWidth - margin, topY, buttonWidth, controlHeight);

            topY += controlHeight + margin;
            _lblBefore.SetBounds(margin, topY, 200, controlHeight);
            topY += controlHeight + 2;

            int totalHeightRemaining = this.ClientSize.Height - topY - margin;
            int halfHeight = totalHeightRemaining / 2;

            _rtbBefore.SetBounds(margin, topY, this.ClientSize.Width - 2 * margin, halfHeight - controlHeight);
            topY += (halfHeight - controlHeight) + margin;

            _lblAfter.SetBounds(margin, topY, 200, controlHeight);
            topY += controlHeight + 2;

            _rtbAfter.SetBounds(margin, topY, this.ClientSize.Width - 2 * margin, halfHeight - controlHeight);
        }

        private void UpdatePreview()
        {
            string findStr = _txtFind.Text ?? string.Empty;
            string replaceStr = _txtReplace.Text ?? string.Empty;

            _rtbBefore.Clear();
            foreach (var view in _selectedViews)
            {
                _rtbBefore.AppendText(view.Name + Environment.NewLine);
            }

            _rtbAfter.Clear();
            foreach (var view in _selectedViews)
            {
                string newName = view.Name.Replace(findStr, replaceStr);
                _rtbAfter.AppendText(newName + Environment.NewLine);
            }
        }
    }
}
