using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ResetTitlesOfSelectedViewports : IExternalCommand
    {
        private const string SETTINGS_FOLDER = "revit-scripts";
        private const string SETTINGS_SUBFOLDER = "ResetTitlesOfSelectedViewports";
        private const string SETTINGS_FILE = "ResetTitlesOfSelectedViewports";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Get the current selection
                Selection selection = uiDoc.Selection;
                ICollection<ElementId> selectedIds = selection.GetElementIds();

                // Filter for viewports and views
                List<Viewport> selectedViewports = new List<Viewport>();
                List<View> selectedViews = new List<View>();
                
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem is Viewport viewport)
                    {
                        selectedViewports.Add(viewport);
                    }
                    else if (elem is View view && !(view is ViewSheet))
                    {
                        selectedViews.Add(view);
                    }
                }

                // For selected views, find their corresponding viewports on sheets
                if (selectedViews.Count > 0)
                {
                    // Get all viewports in the document
                    FilteredElementCollector viewportCollector = new FilteredElementCollector(doc)
                        .OfClass(typeof(Viewport));

                    foreach (View view in selectedViews)
                    {
                        // Find viewports that display this view
                        foreach (Viewport vp in viewportCollector)
                        {
                            if (vp.ViewId == view.Id)
                            {
                                selectedViewports.Add(vp);
                            }
                        }
                    }
                }

                // Remove duplicates in case same viewport was selected directly and through its view
                selectedViewports = selectedViewports.Distinct().ToList();

                if (selectedViewports.Count == 0)
                {
                    TaskDialog.Show("No Valid Selection", 
                        "Please select one or more viewports on sheets, or views that are placed on sheets.");
                    return Result.Cancelled;
                }

                // Get last used value or default
                double offsetMm = GetLastOffsetValue();

                // Prompt user for offset distance
                using (OffsetInputForm inputForm = new OffsetInputForm(offsetMm))
                {
                    if (inputForm.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        return Result.Cancelled;
                    }
                    offsetMm = inputForm.OffsetValue;
                }

                // Save the value for next time
                SaveOffsetValue(offsetMm);

                // Convert mm to feet (Revit's internal units)
                double offsetDistance = UnitUtils.ConvertToInternalUnits(offsetMm, UnitTypeId.Millimeters);

                using (Transaction trans = new Transaction(doc, "Reset Viewport Titles"))
                {
                    trans.Start();

                    int successCount = 0;
                    foreach (Viewport viewport in selectedViewports)
                    {
                        try
                        {
                            // Get the viewport outline on the sheet
                            Outline outline = viewport.GetBoxOutline();
                            
                            // Calculate the viewport height (full height from bottom to top)
                            double viewportHeight = outline.MaximumPoint.Y - outline.MinimumPoint.Y;
                            
                            // Reset the label line length to 0 (no leader line)
                            viewport.LabelLineLength = 0;
                            
                            // Get the current viewport view
                            ElementId viewId = viewport.ViewId;
                            View view = doc.GetElement(viewId) as View;
                            
                            // Get the title block height (approximate)
                            double titleHeight = 0;
                            if (view != null)
                            {
                                // Try to get the annotation size for a more accurate title height
                                Parameter annotationCrop = viewport.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL);
                                if (annotationCrop != null && annotationCrop.AsInteger() == 1)
                                {
                                    // Approximate title block height - typically around 0.1 feet (about 30mm)
                                    titleHeight = UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Millimeters);
                                }
                            }
                            
                            // Calculate the label offset
                            // The title label's top edge should be at the bottom of the viewport + offset
                            // So we need to account for the title height
                            double yOffset = -(viewportHeight / 2 + offsetDistance + titleHeight / 2);
                            
                            // Set the label offset
                            // This positions the title relative to the viewport center
                            viewport.LabelOffset = new XYZ(0, yOffset, 0);

                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            // Continue with next viewport if there's an error with this one
                            TaskDialog.Show("Error", $"Error processing viewport: {ex.Message}");
                            continue;
                        }
                    }

                    trans.Commit();

                    string selectionInfo = "";
                    if (selectedViews.Count > 0)
                    {
                        selectionInfo = $" ({selectedViews.Count} view(s) and {selectedIds.Count - selectedViews.Count} viewport(s) selected)";
                    }
                    
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private double GetLastOffsetValue()
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string settingsPath = Path.Combine(appDataPath, SETTINGS_FOLDER, SETTINGS_SUBFOLDER, SETTINGS_FILE);
                
                if (File.Exists(settingsPath))
                {
                    string content = File.ReadAllText(settingsPath);
                    if (double.TryParse(content, out double value))
                    {
                        return value;
                    }
                }
            }
            catch
            {
                // If any error occurs, just return default
            }
            
            return 0; // Default value
        }

        private void SaveOffsetValue(double value)
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string folderPath = Path.Combine(appDataPath, SETTINGS_FOLDER, SETTINGS_SUBFOLDER);
                
                // Create directory if it doesn't exist
                Directory.CreateDirectory(folderPath);
                
                string settingsPath = Path.Combine(folderPath, SETTINGS_FILE);
                File.WriteAllText(settingsPath, value.ToString());
            }
            catch
            {
                // Silently fail if we can't save the settings
            }
        }
    }

    // Simple input form for offset value
    public class OffsetInputForm : System.Windows.Forms.Form
    {
        private System.Windows.Forms.TextBox textBox;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label label;
        
        public double OffsetValue { get; private set; }

        public OffsetInputForm(double defaultValue)
        {
            InitializeComponents(defaultValue);
        }

        private void InitializeComponents(double defaultValue)
        {
            this.Text = "Viewport Title Offset";
            this.Size = new System.Drawing.Size(300, 150);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Label
            label = new System.Windows.Forms.Label();
            label.Text = "Enter offset from viewport (millimeters):";
            label.Location = new System.Drawing.Point(20, 20);
            label.Size = new System.Drawing.Size(250, 20);

            // TextBox
            textBox = new System.Windows.Forms.TextBox();
            textBox.Text = defaultValue.ToString();
            textBox.Location = new System.Drawing.Point(20, 45);
            textBox.Size = new System.Drawing.Size(250, 20);
            textBox.KeyPress += TextBox_KeyPress;
            textBox.SelectAll(); // Select all text for easy replacement

            // OK Button
            okButton = new System.Windows.Forms.Button();
            okButton.Text = "OK";
            okButton.Location = new System.Drawing.Point(110, 80);
            okButton.Size = new System.Drawing.Size(75, 23);
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Click += OkButton_Click;

            // Cancel Button
            cancelButton = new System.Windows.Forms.Button();
            cancelButton.Text = "Cancel";
            cancelButton.Location = new System.Drawing.Point(195, 80);
            cancelButton.Size = new System.Drawing.Size(75, 23);
            cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;

            // Add controls to form
            this.Controls.Add(label);
            this.Controls.Add(textBox);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);

            // Set accept and cancel buttons
            this.AcceptButton = okButton;
            this.CancelButton = cancelButton;
        }

        private void TextBox_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            System.Windows.Forms.TextBox tb = (System.Windows.Forms.TextBox)sender;
            
            // If minus sign is pressed and text is selected or cursor is at start
            if (e.KeyChar == '-')
            {
                // If all text is selected or cursor is at the beginning, allow it
                if (tb.SelectionLength == tb.Text.Length || 
                    (tb.SelectionStart == 0 && tb.SelectionLength == 0))
                {
                    if (tb.SelectionLength > 0)
                    {
                        // Replace selected text with minus sign
                        tb.Text = "-";
                        tb.SelectionStart = 1;
                        e.Handled = true;
                        return;
                    }
                    else if (tb.Text.IndexOf('-') == -1)
                    {
                        // Allow minus at beginning if not already present
                        return;
                    }
                }
                e.Handled = true;
                return;
            }
            
            // Allow only digits, decimal point, and control characters
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
                return;
            }

            // Allow only one decimal point
            if (e.KeyChar == '.' && tb.Text.IndexOf('.') > -1)
            {
                // If decimal exists but it's in selected text, allow replacement
                if (tb.SelectionLength > 0 && tb.SelectedText.Contains("."))
                {
                    return;
                }
                e.Handled = true;
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            // Validate and parse input - now accepts negative values
            if (double.TryParse(textBox.Text, out double value))
            {
                OffsetValue = value;
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please enter a valid number.", "Invalid Input", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                this.DialogResult = System.Windows.Forms.DialogResult.None;
            }
        }
    }
}
