using System;
using System.Collections.Generic;
using System.Drawing;
using WinForms = System.Windows.Forms;
using RevitDB = Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace MyProject.Commands
{
    /// <summary>
    /// Uses the currently selected views (or viewports) to rename.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RenameSelectedViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, RevitDB.ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uidoc = uiApp.ActiveUIDocument;
            RevitDB.Document doc = uidoc.Document;

            var selectedIds = uidoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("Error", "Please select at least one view or viewport.");
                return Result.Failed;
            }

            var selectedViews = new List<RevitDB.View>();
            var processedViewIds = new HashSet<RevitDB.ElementId>();

            foreach (var id in selectedIds)
            {
                var elem = doc.GetElement(id);
                if (elem is RevitDB.View view && !view.IsTemplate)
                {
                    if (processedViewIds.Add(view.Id))
                        selectedViews.Add(view);
                }
                else if (elem is RevitDB.Viewport viewport)
                {
                    var viewId = viewport.ViewId;
                    if (viewId != null && processedViewIds.Add(viewId))
                    {
                        var assocView = doc.GetElement(viewId) as RevitDB.View;
                        if (assocView != null && !assocView.IsTemplate)
                            selectedViews.Add(assocView);
                    }
                }
            }

            if (selectedViews.Count == 0)
            {
                TaskDialog.Show("Error", "No valid views were selected.");
                return Result.Failed;
            }

            using (var renameForm = new InteractiveFindReplaceForm(selectedViews))
            {
                if (renameForm.ShowDialog() != WinForms.DialogResult.OK)
                    return Result.Succeeded;

                string findStr    = renameForm.FindText;
                string replaceStr = renameForm.ReplaceText;

                using (var tx = new RevitDB.Transaction(doc, "Rename Selected Views"))
                {
                    tx.Start();
                    foreach (var v in selectedViews)
                    {
                        try
                        {
                            var newName = v.Name.Replace(findStr, replaceStr);
                            if (!newName.Equals(v.Name, StringComparison.OrdinalIgnoreCase))
                                v.Name = newName;
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
        private readonly List<RevitDB.View> _selectedViews;
        private WinForms.TextBox      _txtFind;
        private WinForms.TextBox      _txtReplace;
        private WinForms.RichTextBox  _rtbBefore;
        private WinForms.RichTextBox  _rtbAfter;
        private WinForms.Button       _okButton;

        public string FindText    { get; private set; }
        public string ReplaceText { get; private set; }

        public InteractiveFindReplaceForm(List<RevitDB.View> selectedViews)
        {
            _selectedViews = selectedViews;

            // Form setup
            this.FormBorderStyle = WinForms.FormBorderStyle.Sizable;
            this.StartPosition   = WinForms.FormStartPosition.CenterScreen;
            this.Text            = "Find / Replace (Preview)";
            this.Size            = new Size(700, 600);
            this.MinimumSize     = new Size(500, 400);
            this.KeyPreview      = true;
            this.KeyDown        += (s, e) =>
            {
                if (e.KeyCode == WinForms.Keys.Escape)
                {
                    this.DialogResult = WinForms.DialogResult.Cancel;
                    this.Close();
                }
            };

            // Labels & textboxes
            var lblFind    = new WinForms.Label { Text = "Find:" };
            _txtFind       = new WinForms.TextBox();
            var lblReplace = new WinForms.Label { Text = "Replace:" };
            _txtReplace    = new WinForms.TextBox();
            var lblBefore  = new WinForms.Label { Text = "Before (Original):" };
            var lblAfter   = new WinForms.Label { Text = "After (Preview):" };

            _rtbBefore = new WinForms.RichTextBox
            {
                ReadOnly   = true,
                BackColor  = System.Drawing.SystemColors.Window,
                WordWrap   = false,
                ScrollBars = WinForms.RichTextBoxScrollBars.Both
            };
            _rtbAfter = new WinForms.RichTextBox
            {
                ReadOnly   = true,
                BackColor  = System.Drawing.SystemColors.Window,
                WordWrap   = false,
                ScrollBars = WinForms.RichTextBoxScrollBars.Both
            };

            _okButton = new WinForms.Button
            {
                Text         = "OK",
                DialogResult = WinForms.DialogResult.OK
            };
            this.AcceptButton = _okButton;
            _okButton.Click += (s, e) =>
            {
                FindText    = _txtFind.Text;
                ReplaceText = _txtReplace.Text;
                this.DialogResult = WinForms.DialogResult.OK;
                this.Close();
            };

            // Add controls
            this.Controls.AddRange(new WinForms.Control[] {
                lblFind, _txtFind,
                lblReplace, _txtReplace,
                lblBefore, _rtbBefore,
                lblAfter, _rtbAfter,
                _okButton
            });

            // Populate both boxes on startup
            foreach (var v in _selectedViews)
            {
                _rtbBefore.AppendText(v.Name + Environment.NewLine);
                _rtbAfter .AppendText(v.Name + Environment.NewLine);
            }

            // Wire up preview updates
            _txtFind.TextChanged    += (s, e) => UpdatePreview();
            _txtReplace.TextChanged += (s, e) => UpdatePreview();

            this.Resize += (s, e) => LayoutControls();
            LayoutControls();
        }

        private void LayoutControls()
        {
            const int margin = 8;
            int y = margin, labelW = 60, ctrlH = 24;

            // Find / Replace row
            Controls[0].SetBounds(margin, y + 4, labelW, ctrlH);         // lblFind
            _txtFind.SetBounds(labelW + margin * 2, y, 150, ctrlH);
            Controls[2].SetBounds(_txtFind.Right + margin, y + 4, labelW, ctrlH); // lblReplace
            _txtReplace.SetBounds(Controls[2].Right + margin, y, 150, ctrlH);
            _okButton.SetBounds(ClientSize.Width - 80 - margin, y, 80, ctrlH);

            // Before box
            y += ctrlH + margin;
            Controls[4].SetBounds(margin, y, 200, ctrlH);
            y += ctrlH + 2;
            int half = (ClientSize.Height - y - margin) / 2;
            _rtbBefore.SetBounds(margin, y, ClientSize.Width - 2 * margin, half - ctrlH);

            // After box
            y += half;
            Controls[6].SetBounds(margin, y, 200, ctrlH);
            y += ctrlH + 2;
            _rtbAfter.SetBounds(margin, y, ClientSize.Width - 2 * margin, half - ctrlH);
        }

        private void UpdatePreview()
        {
            _rtbAfter.Clear();
            foreach (var v in _selectedViews)
            {
                _rtbAfter.AppendText(v.Name.Replace(_txtFind.Text, _txtReplace.Text) + Environment.NewLine);
            }
        }
    }
}
