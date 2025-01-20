using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class RenameViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        Document doc = uiApp.ActiveUIDocument.Document;

        // 1) Gather all non-template views
        List<Autodesk.Revit.DB.View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(Autodesk.Revit.DB.View))
            .Cast<Autodesk.Revit.DB.View>()
            .Where(v => !v.IsTemplate)
            .OrderBy(v => v.Name)
            .ToList();

        // 2) Let user pick which views to rename (via your existing CustomGUIs.DataGrid)
        List<Autodesk.Revit.DB.View> selectedViews = CustomGUIs.DataGrid(
            allViews, 
            new List<string> { "Name" }, 
            null, 
            "Select Views to Rename"
        );

        if (selectedViews == null || selectedViews.Count == 0)
        {
            // user didn't pick anything
            return Result.Succeeded;
        }

        // 3) Show the second dialog that displays "before" & "after" lines
        using (var renameForm = new InteractiveFindReplaceForm(selectedViews))
        {
            // If user hits Esc or closes the dialog, we treat as canceled
            if (renameForm.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return Result.Succeeded; 
            }

            // Retrieve final "Find" / "Replace" strings
            string findStr = renameForm.FindText;
            string replaceStr = renameForm.ReplaceText;

            // 4) Rename the selected views in a transaction
            using (Transaction tx = new Transaction(doc, "Rename Views"))
            {
                tx.Start();
                foreach (Autodesk.Revit.DB.View v in selectedViews)
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
                        TaskDialog.Show("Error", 
                            $"Could not rename view \"{v.Name}\".\n{ex.Message}");
                    }
                }
                tx.Commit();
            }
        }

        return Result.Succeeded;
    }
}

/// <summary>
/// A WinForms form that shows:
///   - "Find" and "Replace" fields at the top
///   - Two multi-line text boxes, each line representing a selected view:
///        1) "Before" (original names)
///        2) "After" (live-updated preview of rename)
///   - OK button at the bottom
///   - Press Esc at any time to close/cancel
/// 
/// The two text boxes remain at equal height even when the user resizes the form.
/// If you wish to add partial substring highlighting, see the commented code in UpdatePreview().
/// </summary>
public class InteractiveFindReplaceForm : System.Windows.Forms.Form
{
    // Controls
    private System.Windows.Forms.Label _lblFind;
    private System.Windows.Forms.TextBox _txtFind;

    private System.Windows.Forms.Label _lblReplace;
    private System.Windows.Forms.TextBox _txtReplace;

    private System.Windows.Forms.Label _lblBefore;
    private System.Windows.Forms.Label _lblAfter;

    private System.Windows.Forms.RichTextBox _rtbBefore;
    private System.Windows.Forms.RichTextBox _rtbAfter;

    private System.Windows.Forms.Button _okButton;

    // Data
    private readonly List<Autodesk.Revit.DB.View> _selectedViews;

    public string FindText { get; private set; }
    public string ReplaceText { get; private set; }

    public InteractiveFindReplaceForm(List<Autodesk.Revit.DB.View> selectedViews)
    {
        _selectedViews = selectedViews;

        // Make form resizable
        this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        this.Text = "Find / Replace (Preview)";
        this.Size = new Size(700, 600); // default size
        this.MinimumSize = new Size(500, 400); // optional min size

        // We want ESC to close the form as Cancel
        this.KeyPreview = true;  // ensures form sees keystrokes even if controls have focus
        this.KeyDown += (s, e) =>
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Escape)
            {
                this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                this.Close();
            }
        };

        // Create top row controls: "Find" / "Replace"
        _lblFind = new System.Windows.Forms.Label { Text = "Find:" };
        _txtFind = new System.Windows.Forms.TextBox();

        _lblReplace = new System.Windows.Forms.Label { Text = "Replace:" };
        _txtReplace = new System.Windows.Forms.TextBox();

        // "Before" / "After" labels
        _lblBefore = new System.Windows.Forms.Label { Text = "Before (Original):" };
        _lblAfter = new System.Windows.Forms.Label { Text = "After (Preview):" };

        // Two RichTextBoxes for multi-line display
        _rtbBefore = new System.Windows.Forms.RichTextBox 
        {
            ReadOnly = true,
            BackColor = System.Drawing.SystemColors.Window,  // so it looks normal
            WordWrap = false,       // optional if names can be long
            ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Both
        };
        _rtbAfter = new System.Windows.Forms.RichTextBox
        {
            ReadOnly = true,
            BackColor = System.Drawing.SystemColors.Window,
            WordWrap = false,
            ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Both
        };

        // OK button at bottom
        _okButton = new System.Windows.Forms.Button
        {
            Text = "OK",
            DialogResult = System.Windows.Forms.DialogResult.OK
        };
        // Pressing Enter triggers the OK button
        this.AcceptButton = _okButton;
        _okButton.Click += (s, e) =>
        {
            this.FindText = _txtFind.Text;
            this.ReplaceText = _txtReplace.Text;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        };

        // Add them to the form
        this.Controls.Add(_lblFind);
        this.Controls.Add(_txtFind);
        this.Controls.Add(_lblReplace);
        this.Controls.Add(_txtReplace);

        this.Controls.Add(_lblBefore);
        this.Controls.Add(_rtbBefore);

        this.Controls.Add(_lblAfter);
        this.Controls.Add(_rtbAfter);

        this.Controls.Add(_okButton);

        // Layout on form resize
        this.Resize += (s, e) => LayoutControls();

        // Populate initial text
        foreach (var view in _selectedViews)
        {
            _rtbBefore.AppendText(view.Name + Environment.NewLine);
            _rtbAfter.AppendText(view.Name + Environment.NewLine);
        }

        // When user types, update preview
        _txtFind.TextChanged += (s, e) => UpdatePreview();
        _txtReplace.TextChanged += (s, e) => UpdatePreview();

        // Initial layout
        LayoutControls();
    }

    private void LayoutControls()
    {
        // Manual layout with some spacing
        int margin = 8;
        int topY = margin;

        // place "Find" label & text
        int labelWidth = 50;
        int controlHeight = 24;

        _lblFind.SetBounds(margin, topY + 4, labelWidth, controlHeight);
        _txtFind.SetBounds(_lblFind.Right + 5, topY, 150, controlHeight);

        // place "Replace" label & text
        _lblReplace.SetBounds(_txtFind.Right + 20, topY + 4, labelWidth, controlHeight);
        _txtReplace.SetBounds(_lblReplace.Right + 5, topY, 150, controlHeight);

        // OK button in top row (far right)
        int buttonWidth = 60;
        _okButton.SetBounds(this.ClientSize.Width - buttonWidth - margin, topY, buttonWidth, controlHeight);

        topY += controlHeight + margin;

        // Next row: "Before" label
        _lblBefore.SetBounds(margin, topY, 200, controlHeight);
        topY += controlHeight + 2;

        // Then "After" label will be placed halfway down
        // We'll split the vertical space for two RichTextBoxes equally
        int totalHeightRemaining = this.ClientSize.Height - topY - margin;
        // Each text box gets half
        int halfHeight = totalHeightRemaining / 2;

        // "Before" text box
        _rtbBefore.SetBounds(margin, topY, this.ClientSize.Width - 2 * margin, halfHeight - controlHeight);
        topY += (halfHeight - controlHeight) + margin;

        // "After" label
        _lblAfter.SetBounds(margin, topY, 200, controlHeight);
        topY += controlHeight + 2;

        // "After" text box
        _rtbAfter.SetBounds(margin, topY, this.ClientSize.Width - 2 * margin, halfHeight - controlHeight);
    }

    private void UpdatePreview()
    {
        // We'll rebuild the "Before" and "After" text from scratch each time the user types.
        // If you wanted partial substring highlighting in "Before", you could do so by
        // searching for matches and coloring them (as we've shown previously).
        // For now, we'll keep it simpler.

        string findStr = _txtFind.Text ?? string.Empty;
        string replaceStr = _txtReplace.Text ?? string.Empty;

        // Clear & rebuild "Before"
        _rtbBefore.Clear();
        foreach (var view in _selectedViews)
        {
            // If you want to highlight "findStr" in "Before", you could call
            // AppendTextWithHighlight(_rtbBefore, view.Name, findStr);
            // For simplicity, we'll just re-display the original name:
            _rtbBefore.AppendText(view.Name + Environment.NewLine);
        }

        // Clear & rebuild "After" with the replaced names
        _rtbAfter.Clear();
        foreach (var view in _selectedViews)
        {
            string newName = view.Name.Replace(findStr, replaceStr);
            _rtbAfter.AppendText(newName + Environment.NewLine);
        }
    }

    // Example helper if you want partial highlight in "Before" box:
    // private void AppendTextWithHighlight(System.Windows.Forms.RichTextBox rtb, string text, string findText)
    // {
    //     if (string.IsNullOrEmpty(findText))
    //     {
    //         // no highlight, just append
    //         rtb.AppendText(text + Environment.NewLine);
    //         return;
    //     }
    //
    //     int startIdx = 0;
    //     while (true)
    //     {
    //         int matchPos = text.IndexOf(findText, startIdx, StringComparison.OrdinalIgnoreCase);
    //         if (matchPos < 0)
    //         {
    //             // append the remainder normally
    //             rtb.SelectionColor = Color.Black;
    //             rtb.AppendText(text.Substring(startIdx) + Environment.NewLine);
    //             break;
    //         }
    //         // append text before match
    //         rtb.SelectionColor = Color.Black;
    //         rtb.AppendText(text.Substring(startIdx, matchPos - startIdx));
    //
    //         // highlight the match
    //         rtb.SelectionColor = Color.Red; // pick your highlight color
    //         rtb.AppendText(text.Substring(matchPos, findText.Length));
    //
    //         startIdx = matchPos + findText.Length;
    //     }
    // }
}
