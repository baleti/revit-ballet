using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Text.RegularExpressions;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ExportSelectedViewsToDWG : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = commandData.Application.ActiveUIDocument;
        var doc = uidoc.Document;
        
        // Get current selection using the custom method
        var selectedIds = uidoc.GetSelectionIds();
        if (!selectedIds.Any())
        {
            TaskDialog.Show("No Selection", "Please select sheets or views to export.");
            return Result.Cancelled;
        }
        
        // Filter for sheets and views
        var viewsAndSheets = new List<Autodesk.Revit.DB.View>();
        foreach (var id in selectedIds)
        {
            var elem = doc.GetElement(id);
            if (elem is Autodesk.Revit.DB.View view && view.CanBePrinted && 
                !(view.ViewType == ViewType.Internal || view.IsTemplate))
            {
                viewsAndSheets.Add(view);
            }
        }
        
        if (!viewsAndSheets.Any())
        {
            TaskDialog.Show("Invalid Selection", "No exportable views or sheets were selected.");
            return Result.Cancelled;
        }
        
        // Show naming configuration dialog
        using (var dialog = new DWGNamingDialog(doc, viewsAndSheets))
        {
            if (dialog.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;
            
            // Get folder location
            string exportFolder = null;
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select folder for DWG export";
                folderDialog.ShowNewFolderButton = true;
                
                if (folderDialog.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;
                
                exportFolder = folderDialog.SelectedPath;
            }
            
            // Get selected export options
            var exportOptions = dialog.GetSelectedExportOptions();
            if (exportOptions == null)
            {
                exportOptions = new DWGExportOptions();
                exportOptions.MergedViews = false;
            }
            
            var successCount = 0;
            var failedExports = new List<string>();
            
            using (var tx = new Transaction(doc, "Export Views to DWG"))
            {
                tx.Start();
                
                foreach (var view in viewsAndSheets)
                {
                    try
                    {
                        var fileName = dialog.GetFileName(view);
                        var viewIds = new List<ElementId> { view.Id };
                        
                        doc.Export(exportFolder, fileName, viewIds, exportOptions);
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        failedExports.Add($"{view.Name}: {ex.Message}");
                    }
                }
                
                tx.Commit();
            }
            
            // Show results
            var resultMessage = $"Successfully exported {successCount} of {viewsAndSheets.Count} views.";
            if (failedExports.Any())
            {
                resultMessage += $"\n\nFailed exports:\n{string.Join("\n", failedExports.Take(5))}";
                if (failedExports.Count > 5)
                    resultMessage += $"\n...and {failedExports.Count - 5} more.";
            }
            
            TaskDialog.Show("Export Complete", resultMessage);
        }
        
        return Result.Succeeded;
    }
}

// Dialog for configuring DWG naming
public class DWGNamingDialog : System.Windows.Forms.Form
{
    private static readonly string AppDataPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "revit-scripts"
    );
    private static readonly string FormatHistoryFile = Path.Combine(AppDataPath, "ExportSelectedViewsToDWG");
    
    private Document doc;
    private List<Autodesk.Revit.DB.View> views;
    private Dictionary<Autodesk.Revit.DB.View, Dictionary<string, string>> viewParameters;
    private List<string> availableParameters;
    private List<string> filteredParameters;
    private string formatString = "{Sheet Number}_{Sheet Name}";
    private List<ExportDWGSettings> exportSettings;
    private bool isPlaceholderActive = true;
    
    private System.Windows.Forms.TextBox txtFormatString;
    private System.Windows.Forms.TextBox txtSearch;
    private ListBox lstAvailableParams;
    private DataGridView dgvPreview;
    private System.Windows.Forms.ComboBox cmbExportOptions;
    private Button btnInsertParam;
    private Button btnOK;
    private Button btnCancel;
    private SplitContainer splitMain;
    private System.Windows.Forms.Panel pnlTop;
    private System.Windows.Forms.Panel pnlLeft;
    
    public DWGNamingDialog(Document doc, List<Autodesk.Revit.DB.View> views)
    {
        this.doc = doc;
        this.views = views;
        LoadExportSettings();
        InitializeParameters();
        LoadFormatHistory();
        InitializeUI();
        UpdatePreview();
    }
    
    private void LoadExportSettings()
    {
        exportSettings = new List<ExportDWGSettings>();
        
        // Get all DWG export settings in the project
        var collector = new FilteredElementCollector(doc)
            .OfClass(typeof(ExportDWGSettings));
        
        foreach (ExportDWGSettings settings in collector)
        {
            exportSettings.Add(settings);
        }
    }
    
    private void LoadFormatHistory()
    {
        try
        {
            Directory.CreateDirectory(AppDataPath);
            if (File.Exists(FormatHistoryFile))
            {
                var lines = File.ReadAllLines(FormatHistoryFile);
                if (lines.Length > 0 && !string.IsNullOrWhiteSpace(lines[0]))
                {
                    formatString = lines[0];
                }
            }
        }
        catch
        {
            // Use default if load fails
        }
    }
    
    private void SaveFormatHistory()
    {
        try
        {
            Directory.CreateDirectory(AppDataPath);
            var existingLines = new List<string>();
            
            if (File.Exists(FormatHistoryFile))
            {
                existingLines = File.ReadAllLines(FormatHistoryFile).ToList();
                // Remove current format if it exists
                existingLines.RemoveAll(line => line == formatString);
            }
            
            // Add current format at the beginning
            existingLines.Insert(0, formatString);
            
            // Keep only last 20 formats
            if (existingLines.Count > 20)
                existingLines = existingLines.Take(20).ToList();
            
            File.WriteAllLines(FormatHistoryFile, existingLines);
        }
        catch
        {
            // Ignore save errors
        }
    }
    
    private void InitializeParameters()
    {
        viewParameters = new Dictionary<Autodesk.Revit.DB.View, Dictionary<string, string>>();
        var paramSet = new HashSet<string>();
        
        // Add standard view/sheet properties
        paramSet.Add("View Name");
        paramSet.Add("View Type");
        
        foreach (var view in views)
        {
            var parameters = new Dictionary<string, string>();
            
            // Standard properties
            parameters["View Name"] = view.Name;
            parameters["View Type"] = view.ViewType.ToString();
            
            if (view is ViewSheet sheet)
            {
                parameters["Sheet Number"] = sheet.SheetNumber;
                parameters["Sheet Name"] = sheet.Name;
                paramSet.Add("Sheet Number");
                paramSet.Add("Sheet Name");
            }
            
            // Instance parameters
            foreach (Parameter param in view.Parameters)
            {
                if (param.HasValue)
                {
                    string value = null;
                    switch (param.StorageType)
                    {
                        case StorageType.String:
                            value = param.AsString();
                            break;
                        case StorageType.Integer:
                            value = param.AsInteger().ToString();
                            break;
                        case StorageType.Double:
                            value = param.AsValueString();
                            break;
                    }
                    
                    if (!string.IsNullOrEmpty(value))
                    {
                        var paramName = param.Definition.Name;
                        parameters[paramName] = value;
                        paramSet.Add(paramName);
                    }
                }
            }
            
            viewParameters[view] = parameters;
        }
        
        availableParameters = paramSet.OrderBy(p => p).ToList();
        filteredParameters = new List<string>(availableParameters);
    }
    
    private void InitializeUI()
    {
        this.Text = "Configure DWG Export Naming";
        this.Size = new System.Drawing.Size(1000, 700);
        this.StartPosition = FormStartPosition.CenterScreen;
        
        // Top panel for format string and export options
        pnlTop = new System.Windows.Forms.Panel
        {
            Dock = DockStyle.Top,
            Height = 100,
            Padding = new Padding(12)
        };
        
        var lblFormat = new Label
        {
            Text = "Format String:",
            Location = new System.Drawing.Point(12, 12),
            Size = new System.Drawing.Size(100, 20)
        };
        
        txtFormatString = new System.Windows.Forms.TextBox
        {
            Text = formatString,
            Location = new System.Drawing.Point(12, 35),
            Size = new System.Drawing.Size(600, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        txtFormatString.TextChanged += TxtFormatString_TextChanged;
        txtFormatString.KeyPress += TxtFormatString_KeyPress;
        
        btnInsertParam = new Button
        {
            Text = "Insert Selected Parameter",
            Location = new System.Drawing.Point(620, 35),
            Size = new System.Drawing.Size(150, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnInsertParam.Click += BtnInsertParam_Click;
        
        var lblExportOptions = new Label
        {
            Text = "DWG Export Settings:",
            Location = new System.Drawing.Point(12, 65),
            Size = new System.Drawing.Size(120, 20)
        };
        
        cmbExportOptions = new System.Windows.Forms.ComboBox
        {
            Location = new System.Drawing.Point(135, 65),
            Size = new System.Drawing.Size(300, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        
        // Populate export options
        cmbExportOptions.Items.Add("<Default>");
        foreach (var settings in exportSettings)
        {
            cmbExportOptions.Items.Add(settings.Name);
        }
        cmbExportOptions.SelectedIndex = 0;
        
        pnlTop.Controls.AddRange(new System.Windows.Forms.Control[] {
            lblFormat, txtFormatString, btnInsertParam,
            lblExportOptions, cmbExportOptions
        });
        
        // Bottom panel for buttons
        var pnlBottom = new System.Windows.Forms.Panel
        {
            Dock = DockStyle.Bottom,
            Height = 50
        };
        
        btnOK = new Button
        {
            Text = "OK",
            Size = new System.Drawing.Size(75, 30),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnOK.Location = new System.Drawing.Point(this.ClientSize.Width - btnOK.Width - 100, 10);
        btnOK.DialogResult = DialogResult.OK;
        btnOK.Click += BtnOK_Click;
        
        btnCancel = new Button
        {
            Text = "Cancel",
            Size = new System.Drawing.Size(75, 30),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnCancel.Location = new System.Drawing.Point(this.ClientSize.Width - btnCancel.Width - 15, 10);
        btnCancel.DialogResult = DialogResult.Cancel;
        
        pnlBottom.Controls.AddRange(new System.Windows.Forms.Control[] { btnOK, btnCancel });
        
        // Main split container
        splitMain = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            SplitterDistance = 250, // 25% for parameters
            TabStop = false // Prevent tab focus on splitter
        };
        
        // Left panel for parameters
        pnlLeft = new System.Windows.Forms.Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12)
        };
        
        txtSearch = new System.Windows.Forms.TextBox
        {
            Location = new System.Drawing.Point(12, 35),
            Size = new System.Drawing.Size(220, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            ForeColor = SystemColors.GrayText,
            Text = "Filter Parameters"
        };
        txtSearch.TextChanged += TxtSearch_TextChanged;
        txtSearch.KeyPress += TxtSearch_KeyPress;
        txtSearch.Enter += TxtSearch_Enter;
        txtSearch.Leave += TxtSearch_Leave;
        
        var lblAvailable = new Label
        {
            Text = "Available Parameters:",
            Location = new System.Drawing.Point(12, 65),
            Size = new System.Drawing.Size(200, 20)
        };
        
        lstAvailableParams = new ListBox
        {
            Location = new System.Drawing.Point(12, 85),
            Size = new System.Drawing.Size(220, 380),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };
        lstAvailableParams.DoubleClick += LstAvailableParams_DoubleClick;
        lstAvailableParams.KeyPress += LstAvailableParams_KeyPress;
        lstAvailableParams.Items.AddRange(availableParameters.ToArray());
        
        var btnAddParam = new Button
        {
            Text = "Add Parameter",
            Location = new System.Drawing.Point(12, 470),
            Size = new System.Drawing.Size(220, 25),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };
        btnAddParam.Click += BtnAddParam_Click;
        
        pnlLeft.Controls.AddRange(new System.Windows.Forms.Control[] {
            txtSearch, lblAvailable, lstAvailableParams, btnAddParam
        });
        
        splitMain.Panel1.Controls.Add(pnlLeft);
        
        // Right panel for preview
        var pnlRight = new System.Windows.Forms.Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 12, 12, 12) // No left padding to use full width
        };
        
        var lblPreview = new Label
        {
            Text = "Preview:",
            Location = new System.Drawing.Point(0, 12),
            Size = new System.Drawing.Size(100, 20)
        };
        
        dgvPreview = new DataGridView
        {
            Location = new System.Drawing.Point(0, 35),
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            RowHeadersVisible = false
        };
        
        // Size the DataGridView to fit within the panel properly
        dgvPreview.Size = new System.Drawing.Size(
            pnlRight.Width - pnlRight.Padding.Left - pnlRight.Padding.Right,
            pnlRight.Height - 35 - pnlRight.Padding.Top - pnlRight.Padding.Bottom
        );
        dgvPreview.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        
        pnlRight.Controls.AddRange(new System.Windows.Forms.Control[] {
            lblPreview, dgvPreview
        });
        
        splitMain.Panel2.Controls.Add(pnlRight);
        
        // Add controls to form
        this.Controls.Add(splitMain);
        this.Controls.Add(pnlBottom);
        this.Controls.Add(pnlTop);
        
        // Set splitter position to 25% after form loads
        this.Load += (s, e) => { 
            splitMain.SplitterDistance = this.Width / 4;
            txtFormatString.Focus(); // Set initial focus to format string
        };
        
        // Set accept and cancel buttons
        this.AcceptButton = btnOK;
        this.CancelButton = btnCancel;
    }
    
    private void TxtSearch_TextChanged(object sender, EventArgs e)
    {
        // Skip filtering if showing placeholder
        if (isPlaceholderActive)
            return;
            
        var searchText = txtSearch.Text.ToLower();
        lstAvailableParams.Items.Clear();
        
        if (string.IsNullOrWhiteSpace(searchText))
        {
            filteredParameters = new List<string>(availableParameters);
        }
        else
        {
            filteredParameters = availableParameters
                .Where(p => p.ToLower().Contains(searchText))
                .ToList();
        }
        
        lstAvailableParams.Items.AddRange(filteredParameters.ToArray());
    }
    
    private void TxtSearch_Enter(object sender, EventArgs e)
    {
        if (isPlaceholderActive)
        {
            txtSearch.Text = "";
            txtSearch.ForeColor = SystemColors.WindowText;
            isPlaceholderActive = false;
        }
    }
    
    private void TxtSearch_Leave(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtSearch.Text))
        {
            txtSearch.Text = "Filter Parameters";
            txtSearch.ForeColor = SystemColors.GrayText;
            isPlaceholderActive = true;
            
            // Reset the list to show all parameters
            lstAvailableParams.Items.Clear();
            lstAvailableParams.Items.AddRange(availableParameters.ToArray());
        }
    }
    
    private void TxtFormatString_TextChanged(object sender, EventArgs e)
    {
        formatString = txtFormatString.Text;
        UpdatePreview();
    }
    
    private void BtnInsertParam_Click(object sender, EventArgs e)
    {
        if (lstAvailableParams.SelectedItem != null)
        {
            InsertParameter(lstAvailableParams.SelectedItem.ToString());
        }
    }
    
    private void LstAvailableParams_DoubleClick(object sender, EventArgs e)
    {
        if (lstAvailableParams.SelectedItem != null)
        {
            InsertParameter(lstAvailableParams.SelectedItem.ToString());
        }
    }
    
    private void InsertParameter(string paramName)
    {
        var insertText = $"{{{paramName}}}";
        var selectionStart = txtFormatString.SelectionStart;
        txtFormatString.Text = txtFormatString.Text.Insert(selectionStart, insertText);
        txtFormatString.SelectionStart = selectionStart + insertText.Length;
        txtFormatString.Focus();
    }
    
    private void BtnOK_Click(object sender, EventArgs e)
    {
        SaveFormatHistory();
    }
    
    private void BtnAddParam_Click(object sender, EventArgs e)
    {
        if (lstAvailableParams.SelectedItem != null)
        {
            InsertParameter(lstAvailableParams.SelectedItem.ToString());
        }
    }
    
    private void TxtFormatString_KeyPress(object sender, KeyPressEventArgs e)
    {
        // Prevent Enter from triggering OK when in format string textbox
        if (e.KeyChar == (char)Keys.Enter)
        {
            e.Handled = true;
        }
    }
    
    private void TxtSearch_KeyPress(object sender, KeyPressEventArgs e)
    {
        // Prevent Enter from triggering OK when in search textbox
        if (e.KeyChar == (char)Keys.Enter)
        {
            e.Handled = true;
        }
    }
    
    private void LstAvailableParams_KeyPress(object sender, KeyPressEventArgs e)
    {
        // Add parameter on Enter when in parameter list
        if (e.KeyChar == (char)Keys.Enter && lstAvailableParams.SelectedItem != null)
        {
            InsertParameter(lstAvailableParams.SelectedItem.ToString());
            e.Handled = true;
        }
    }
    
    private void UpdatePreview()
    {
        dgvPreview.Rows.Clear();
        dgvPreview.Columns.Clear();
        
        dgvPreview.Columns.Add("Original", "Original Name");
        dgvPreview.Columns.Add("NewName", "DWG File Name");
        
        // Set equal column widths
        dgvPreview.Columns[0].FillWeight = 50;
        dgvPreview.Columns[1].FillWeight = 50;
        
        foreach (var view in views.Take(15)) // Show first 15 for performance
        {
            var originalName = view.Name;
            var newName = GetFileName(view) + ".dwg"; // Add .dwg for preview
            dgvPreview.Rows.Add(originalName, newName);
        }
        
        if (views.Count > 15)
        {
            dgvPreview.Rows.Add("...", $"... and {views.Count - 15} more");
        }
    }
    
    public string GetFileName(Autodesk.Revit.DB.View view)
    {
        var result = formatString;
        var parameters = viewParameters[view];
        
        // Replace all parameter placeholders
        var matches = Regex.Matches(result, @"\{([^}]+)\}");
        foreach (Match match in matches)
        {
            var paramName = match.Groups[1].Value;
            if (parameters.ContainsKey(paramName))
            {
                var value = CleanFileName(parameters[paramName]);
                result = result.Replace(match.Value, value);
            }
            else
            {
                result = result.Replace(match.Value, "");
            }
        }
        
        // Clean up any multiple underscores
        result = Regex.Replace(result, @"_{2,}", "_");
        result = result.Trim('_');
        
        // Ensure valid filename
        if (string.IsNullOrWhiteSpace(result))
            result = "Untitled";
            
        return result;
    }
    
    private string CleanFileName(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "";
            
        // Remove invalid filename characters but keep spaces
        var invalidChars = Path.GetInvalidFileNameChars().Where(c => c != ' ').ToArray();
        var result = string.Join("", input.Split(invalidChars));
        
        return result;
    }
    
    public DWGExportOptions GetSelectedExportOptions()
    {
        if (cmbExportOptions.SelectedIndex <= 0)
            return null; // Use default
        
        var settingsName = cmbExportOptions.SelectedItem.ToString();
        var settings = exportSettings.FirstOrDefault(s => s.Name == settingsName);
        
        if (settings != null)
        {
            return settings.GetDWGExportOptions();
        }
        
        return null;
    }
}
