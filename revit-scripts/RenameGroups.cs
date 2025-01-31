using System;
using System.Collections.Generic;
using System.Linq;
using WinForms = System.Windows.Forms;
using Draw = System.Drawing;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class RenameGroups : IExternalCommand
{
    private class GroupTypeInfo
    {
        public GroupType GroupType { get; set; }
        public string CurrentName { get; set; }
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        try
        {
            var selectedIds = uiDoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("Error", "Please select group instances first.");
                return Result.Cancelled;
            }

            // Collect unique group types from selected groups
            var groupTypeIds = new HashSet<ElementId>();
            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is Group group)
                {
                    groupTypeIds.Add(group.GroupType.Id);
                }
            }

            var groupTypes = groupTypeIds
                .Select(id => doc.GetElement(id) as GroupType)
                .Where(gt => gt != null)
                .ToList();

            if (groupTypes.Count == 0)
            {
                TaskDialog.Show("Error", "No model or detail groups selected.");
                return Result.Cancelled;
            }

            var groupTypeInfos = groupTypes
                .Select(gt => new GroupTypeInfo { GroupType = gt, CurrentName = gt.Name })
                .ToList();

            using (var form = new RenameGroupsForm(groupTypeInfos))
            {
                if (form.ShowDialog() != WinForms.DialogResult.OK)
                    return Result.Succeeded;

                using (var tx = new Transaction(doc, "Rename Groups"))
                {
                    tx.Start();
                    foreach (var gti in groupTypeInfos)
                    {
                        try
                        {
                            string currentName = gti.CurrentName;
                            string newName = currentName;

                            if (!string.IsNullOrEmpty(form.FindText) && currentName.Contains(form.FindText))
                            {
                                newName = newName.Replace(form.FindText, form.ReplaceText);
                            }

                            if (!string.IsNullOrEmpty(form.PatternText))
                            {
                                newName = form.PatternText.Replace("{}", newName);
                            }

                            if (newName != currentName)
                            {
                                gti.GroupType.Name = newName;
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Error", $"Failed to rename {gti.CurrentName}: {ex.Message}");
                        }
                    }
                    tx.Commit();
                }
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Critical Error", ex.ToString());
            return Result.Failed;
        }
    }

    private class RenameGroupsForm : WinForms.Form
    {
        private readonly List<GroupTypeInfo> _groupTypeInfos;
        private WinForms.TextBox _txtFind;
        private WinForms.TextBox _txtReplace;
        private WinForms.TextBox _txtPattern;
        private WinForms.RichTextBox _rtbBefore;
        private WinForms.RichTextBox _rtbAfter;
        private WinForms.Button _btnCancel;

        public string FindText => _txtFind.Text;
        public string ReplaceText => _txtReplace.Text;
        public string PatternText => _txtPattern.Text;

        public RenameGroupsForm(List<GroupTypeInfo> groupTypeInfos)
        {
            _groupTypeInfos = groupTypeInfos;
            InitializeComponent();
            InitializeData();
        }

        private void InitializeComponent()
        {
            this.Text = "Rename Groups";
            this.Size = new Draw.Size(600, 500);
            this.MinimumSize = new Draw.Size(500, 400);
            this.Font = new Draw.Font("Segoe UI", 9);
            this.KeyPreview = true;

            var mainLayout = new WinForms.TableLayoutPanel
            {
                Dock = WinForms.DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 6,
                Padding = new WinForms.Padding(10),
                ColumnStyles = 
                {
                    new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 80),
                    new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100)
                },
                RowStyles = 
                {
                    new WinForms.RowStyle(WinForms.SizeType.Absolute, 30),
                    new WinForms.RowStyle(WinForms.SizeType.Absolute, 30),
                    new WinForms.RowStyle(WinForms.SizeType.Absolute, 30),
                    new WinForms.RowStyle(WinForms.SizeType.Absolute, 20),
                    new WinForms.RowStyle(WinForms.SizeType.Percent, 45),
                    new WinForms.RowStyle(WinForms.SizeType.Percent, 45)
                }
            };

            mainLayout.Controls.Add(new WinForms.Label { Text = "Find:", TextAlign = Draw.ContentAlignment.MiddleRight }, 0, 0);
            _txtFind = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill };
            mainLayout.Controls.Add(_txtFind, 1, 0);

            mainLayout.Controls.Add(new WinForms.Label { Text = "Replace:", TextAlign = Draw.ContentAlignment.MiddleRight }, 0, 1);
            _txtReplace = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill };
            mainLayout.Controls.Add(_txtReplace, 1, 1);

            mainLayout.Controls.Add(new WinForms.Label { Text = "Pattern:", TextAlign = Draw.ContentAlignment.MiddleRight }, 0, 2);
            _txtPattern = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill, Text = "{}" };
            mainLayout.Controls.Add(_txtPattern, 1, 2);

            var lblHint = new WinForms.Label 
            { 
                Text = "Use {} to represent current value",
                Font = new Draw.Font(this.Font, Draw.FontStyle.Italic),
                ForeColor = Draw.Color.Gray,
                Dock = WinForms.DockStyle.Top
            };
            mainLayout.Controls.Add(lblHint, 1, 3);

            _rtbBefore = new WinForms.RichTextBox { ReadOnly = true, Dock = WinForms.DockStyle.Fill };
            var groupBefore = new WinForms.GroupBox { Text = "Current Names", Dock = WinForms.DockStyle.Fill };
            groupBefore.Controls.Add(_rtbBefore);
            mainLayout.Controls.Add(groupBefore, 0, 4);
            mainLayout.SetColumnSpan(groupBefore, 2);

            _rtbAfter = new WinForms.RichTextBox { ReadOnly = true, Dock = WinForms.DockStyle.Fill };
            var groupAfter = new WinForms.GroupBox { Text = "Preview", Dock = WinForms.DockStyle.Fill };
            groupAfter.Controls.Add(_rtbAfter);
            mainLayout.Controls.Add(groupAfter, 0, 5);
            mainLayout.SetColumnSpan(groupAfter, 2);

            var btnOk = new WinForms.Button 
            { 
                Text = "OK",
                Size = new Draw.Size(80, 30),
                DialogResult = WinForms.DialogResult.OK
            };

            _btnCancel = new WinForms.Button
            {
                Text = "Cancel",
                Size = new Draw.Size(80, 30),
                DialogResult = WinForms.DialogResult.Cancel
            };

            var buttonPanel = new WinForms.Panel { Dock = WinForms.DockStyle.Bottom, Height = 40 };
            buttonPanel.Controls.Add(btnOk);
            buttonPanel.Controls.Add(_btnCancel);
            btnOk.Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right;
            _btnCancel.Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right;
            
            btnOk.Location = new Draw.Point(buttonPanel.Width - btnOk.Width - 10, 5);
            _btnCancel.Location = new Draw.Point(buttonPanel.Width - btnOk.Width - _btnCancel.Width - 20, 5);

            _txtFind.TextChanged += UpdatePreview;
            _txtReplace.TextChanged += UpdatePreview;
            _txtPattern.TextChanged += UpdatePreview;
            this.AcceptButton = btnOk;
            this.CancelButton = _btnCancel;
            this.Controls.Add(mainLayout);
            this.Controls.Add(buttonPanel);
        }

        protected override bool ProcessDialogKey(WinForms.Keys keyData)
        {
            if (keyData == WinForms.Keys.Escape)
            {
                this.DialogResult = WinForms.DialogResult.Cancel;
                this.Close();
                return true;
            }
            return base.ProcessDialogKey(keyData);
        }

        private void InitializeData()
        {
            foreach (var gti in _groupTypeInfos)
            {
                _rtbBefore.AppendText($"{gti.CurrentName}{Environment.NewLine}");
            }
            UpdatePreview(this, EventArgs.Empty);
        }

        private void UpdatePreview(object sender, EventArgs e)
        {
            _rtbAfter.Clear();
            foreach (var gti in _groupTypeInfos)
            {
                string current = ApplyTransformations(gti.CurrentName);
                _rtbAfter.AppendText($"{current}{Environment.NewLine}");
            }
        }

        private string ApplyTransformations(string input)
        {
            if (!string.IsNullOrEmpty(FindText) && input.Contains(FindText))
            {
                input = input.Replace(FindText, ReplaceText);
            }

            if (!string.IsNullOrEmpty(PatternText))
            {
                input = PatternText.Replace("{}", input);
            }

            return input;
        }
    }
}
