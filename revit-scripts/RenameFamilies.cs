using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace MyRevitAddin
{
    /// <summary>
    /// Command that prompts the user to select families using a DataGrid,
    /// then shows a renaming dialog to update family names.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RenameFamilies : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Collect all Family elements in the document.
            List<Family> families = new FilteredElementCollector(doc)
                                        .OfClass(typeof(Family))
                                        .Cast<Family>()
                                        .ToList();

            if (families.Count == 0)
            {
                WinForms.MessageBox.Show("No families found in the document.", "Error");
                return Result.Cancelled;
            }

            // Build grid entries for the DataGrid.
            // Each row shows the Family Id and its current Name.
            List<Dictionary<string, object>> gridEntries = new List<Dictionary<string, object>>();
            foreach (Family fam in families)
            {
                var row = new Dictionary<string, object>
                {
                    { "Id", fam.Id.IntegerValue },
                    { "Name", fam.Name }
                };
                gridEntries.Add(row);
            }
            List<string> propertyNames = new List<string> { "Id", "Name" };

            // Prompt the user to select families.
            var selectedRows = CustomGUIs.DataGrid(gridEntries, propertyNames, spanAllScreens: false);
            if (selectedRows == null || selectedRows.Count == 0)
                return Result.Succeeded;

            // Build a list of FamilyRenameInfo objects from the selected rows.
            List<FamilyRenameInfo> renameInfos = new List<FamilyRenameInfo>();
            foreach (var row in selectedRows)
            {
                int idVal = Convert.ToInt32(row["Id"]);
                ElementId id = new ElementId(idVal);
                Family fam = doc.GetElement(id) as Family;
                if (fam != null)
                {
                    renameInfos.Add(new FamilyRenameInfo
                    {
                        Family = fam,
                        CurrentName = fam.Name
                    });
                }
            }

            // Show the rename dialog.
            using (var form = new RenameFamilyForm(renameInfos))
            {
                var result = form.ShowDialog();
                if (result != WinForms.DialogResult.OK)
                    return Result.Succeeded;

                string findText = form.FindText;
                string replaceText = form.ReplaceText;
                string patternText = form.PatternText;

                // Start a transaction to apply the renaming.
                using (Transaction tx = new Transaction(doc, "Rename Families"))
                {
                    tx.Start();
                    foreach (var info in renameInfos)
                    {
                        string newName = info.CurrentName;
                        if (!string.IsNullOrEmpty(findText) && newName.Contains(findText))
                            newName = newName.Replace(findText, replaceText);
                        if (!string.IsNullOrEmpty(patternText))
                            newName = patternText.Replace("{}", newName);

                        // Only update if the name has changed.
                        if (newName != info.CurrentName)
                        {
                            try
                            {
                                // Rename the family (works in both family and project documents).
                                info.Family.Name = newName;
                            }
                            catch (Exception ex)
                            {
                                WinForms.MessageBox.Show(
                                    $"Failed to rename family '{info.Family.Name}':\n{ex.Message}", 
                                    "Error");
                            }
                        }
                    }
                    tx.Commit();
                }
            }

            return Result.Succeeded;
        }
    }

    /// <summary>
    /// Helper class to hold a Family element and its current name.
    /// </summary>
    public class FamilyRenameInfo
    {
        public Family Family { get; set; }
        public string CurrentName { get; set; }
    }

    /// <summary>
    /// A WinForms dialog for renaming families.
    /// This dialog is modeled after the provided RenameParamForm.
    /// </summary>
    public class RenameFamilyForm : WinForms.Form
    {
        private readonly List<FamilyRenameInfo> _familyInfos;
        private WinForms.TextBox _txtFind;
        private WinForms.TextBox _txtReplace;
        private WinForms.TextBox _txtPattern;
        private WinForms.RichTextBox _rtbBefore;
        private WinForms.RichTextBox _rtbAfter;

        public string FindText => _txtFind.Text;
        public string ReplaceText => _txtReplace.Text;
        public string PatternText => _txtPattern.Text;

        public RenameFamilyForm(List<FamilyRenameInfo> familyInfos)
        {
            _familyInfos = familyInfos;
            InitializeComponent();
            InitializeData();
            this.KeyDown += (s, e) =>
            {
                if (e.KeyCode == WinForms.Keys.Escape)
                {
                    this.DialogResult = WinForms.DialogResult.Cancel;
                    this.Close();
                }
            };
        }

        private void InitializeComponent()
        {
            this.Text = "Rename Families";
            this.Size = new Drawing.Size(600, 500);
            this.MinimumSize = new Drawing.Size(500, 400);
            this.Font = new Drawing.Font("Segoe UI", 9);
            this.KeyPreview = true;

            var mainLayout = new WinForms.TableLayoutPanel
            {
                Dock = WinForms.DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 6,
                Padding = new WinForms.Padding(10)
            };
            mainLayout.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 80));
            mainLayout.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100));
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30));
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30));
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30));
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 20));
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));

            var lblFind = new WinForms.Label
            {
                Text = "Find:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock = WinForms.DockStyle.Fill
            };
            mainLayout.Controls.Add(lblFind, 0, 0);
            _txtFind = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill };
            mainLayout.Controls.Add(_txtFind, 1, 0);

            var lblReplace = new WinForms.Label
            {
                Text = "Replace:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock = WinForms.DockStyle.Fill
            };
            mainLayout.Controls.Add(lblReplace, 0, 1);
            _txtReplace = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill };
            mainLayout.Controls.Add(_txtReplace, 1, 1);

            var lblPattern = new WinForms.Label
            {
                Text = "Pattern:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock = WinForms.DockStyle.Fill
            };
            mainLayout.Controls.Add(lblPattern, 0, 2);
            _txtPattern = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill, Text = "{}" };
            mainLayout.Controls.Add(_txtPattern, 1, 2);

            var lblHint = new WinForms.Label
            {
                Text = "Use {} to represent current value",
                Font = new Drawing.Font(this.Font, Drawing.FontStyle.Italic),
                ForeColor = Drawing.Color.Gray,
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

            var buttonPanel = new WinForms.Panel { Dock = WinForms.DockStyle.Bottom, Height = 40 };
            var btnCancel = new WinForms.Button
            {
                Text = "Cancel",
                Size = new Drawing.Size(80, 30),
                DialogResult = WinForms.DialogResult.Cancel
            };
            buttonPanel.Controls.Add(btnCancel);
            btnCancel.Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right;
            var btnOk = new WinForms.Button
            {
                Text = "OK",
                Size = new Drawing.Size(80, 30),
                DialogResult = WinForms.DialogResult.OK
            };
            buttonPanel.Controls.Add(btnOk);
            btnOk.Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right;
            btnOk.Location = new Drawing.Point(buttonPanel.Width - btnOk.Width - 10, buttonPanel.Height - btnOk.Height - 5);
            btnCancel.Location = new Drawing.Point(btnOk.Left - btnCancel.Width - 10, btnOk.Top);

            // Update preview when any text changes.
            _txtFind.TextChanged += UpdatePreview;
            _txtReplace.TextChanged += UpdatePreview;
            _txtPattern.TextChanged += UpdatePreview;

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;

            this.Controls.Add(mainLayout);
            this.Controls.Add(buttonPanel);
        }

        private void InitializeData()
        {
            foreach (var info in _familyInfos)
                _rtbBefore.AppendText(info.CurrentName + Environment.NewLine);
            UpdatePreview(this, EventArgs.Empty);
        }

        private void UpdatePreview(object sender, EventArgs e)
        {
            _rtbAfter.Clear();
            foreach (var info in _familyInfos)
            {
                string current = ApplyTransformations(info.CurrentName);
                _rtbAfter.AppendText(current + Environment.NewLine);
            }
        }

        private string ApplyTransformations(string input)
        {
            string output = input;
            if (!string.IsNullOrEmpty(FindText) && output.Contains(FindText))
                output = output.Replace(FindText, ReplaceText);
            if (!string.IsNullOrEmpty(PatternText))
                output = PatternText.Replace("{}", output);
            return output;
        }
    }
}
