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
using Application = System.Windows.Forms.Application;
using System.Runtime.InteropServices;
using Microsoft.WindowsAPICodePack.Dialogs;

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
            Autodesk.Revit.UI.TaskDialog.Show("No Selection", "Please select sheets or views to export.");
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
            Autodesk.Revit.UI.TaskDialog.Show("Invalid Selection", "No exportable views or sheets were selected.");
            return Result.Cancelled;
        }
        
        // Show naming configuration dialog
        using (var dialog = new DWGNamingDialog(doc, viewsAndSheets))
        {
            if (dialog.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;
            
            // Get folder location using CommonOpenFileDialog
            string exportFolder = null;
            var lastPath = dialog.GetLastExportPath();
            
            var folderDialog = new CommonOpenFileDialog
            {
                Title = "Select folder for DWG export",
                IsFolderPicker = true,
                InitialDirectory = lastPath,
                EnsurePathExists = true,
                EnsureFileExists = false
            };
            
            if (folderDialog.ShowDialog(commandData.Application.MainWindowHandle) == CommonFileDialogResult.Ok)
            {
                exportFolder = folderDialog.FileName;
            }
            else
            {
                return Result.Cancelled;
            }
            
            // Save the export path along with other settings
            dialog.SaveExportSettings(exportFolder);
            
            // Get selected export options
            var exportOptions = dialog.GetSelectedExportOptions();
            if (exportOptions == null)
            {
                exportOptions = new DWGExportOptions();
                exportOptions.MergedViews = false;
            }
            
            var successCount = 0;
            var failedExports = new List<string>();
            var cancelled = false;
            
            // Option 1: Using custom progress dialog (comment out if using Option 2)
            using (var progressDialog = new ExportProgressDialog(viewsAndSheets.Count))
            {
                progressDialog.Show(new RevitWindow(commandData.Application.MainWindowHandle));
                
                using (var tx = new Transaction(doc, "Export Views to DWG"))
                {
                    tx.Start();
                    
                    for (int i = 0; i < viewsAndSheets.Count; i++)
                    {
                        if (progressDialog.IsCancelled)
                        {
                            cancelled = true;
                            break;
                        }
                        
                        var view = viewsAndSheets[i];
                        progressDialog.UpdateProgress(i, $"Exporting: {view.Name}");
                        Application.DoEvents(); // Allow UI to update
                        
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
                    
                    if (cancelled)
                        tx.RollBack();
                    else
                        tx.Commit();
                }
            }
            
            // Show results
            if (cancelled)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Cancelled", "DWG export was cancelled by user.");
            }
            else
            {
                var resultMessage = $"Successfully exported {successCount} of {viewsAndSheets.Count} views.";
                if (failedExports.Any())
                {
                    resultMessage += $"\n\nFailed exports:\n{string.Join("\n", failedExports.Take(5))}";
                    if (failedExports.Count > 5)
                        resultMessage += $"\n...and {failedExports.Count - 5} more.";
                }
                
                Autodesk.Revit.UI.TaskDialog.Show("Export Complete", resultMessage);
            }
        }
        
        return Result.Succeeded;
    }
}

public class RevitWindow : IWin32Window
{
    public IntPtr Handle { get; private set; }
    public RevitWindow(IntPtr handle)
    {
        Handle = handle;
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
    private static readonly string ExportSettingsFile = Path.Combine(AppDataPath, "ExportSelectedViewsToDWG_Settings");
    
    private Document doc;
    private List<Autodesk.Revit.DB.View> views;
    private Dictionary<Autodesk.Revit.DB.View, Dictionary<string, string>> viewParameters;
    private List<string> availableParameters;
    private string formatString = "{Sheet Number}_{Sheet Name}";
    private List<ExportDWGSettings> exportSettings;
    private bool isPlaceholderActive = true;
    private List<string> formatHistory;
    private int lastCursorPosition = 0;
    private Dictionary<string, string> pdfNamingPresets; // Store PDF naming presets
    private Stack<string> undoStack = new Stack<string>(); // For undo functionality
    private Stack<string> redoStack = new Stack<string>(); // For redo functionality
    
    private System.Windows.Forms.ComboBox cmbFormatString;
    private System.Windows.Forms.ComboBox cmbPDFNamingPresets;
    private System.Windows.Forms.TextBox txtSearch;
    private DataGridView dgvAvailableParams;
    private DataGridView dgvPreview;
    private System.Windows.Forms.ComboBox cmbExportOptions;
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
        LoadPDFNamingPresets();
        InitializeParameters();
        LoadFormatHistory();
        InitializeUI();
        UpdatePreview();
    }
    
    private void LoadPDFNamingPresets()
    {
        pdfNamingPresets = new Dictionary<string, string>();
        
        try
        {
            // Get the active PDF export settings
            ExportPDFSettings activePdfSettings = ExportPDFSettings.GetActivePredefinedSettings(doc);
            if (activePdfSettings != null)
            {
                AddPDFNamingPreset("Active PDF Settings", activePdfSettings);
            }
            
            // Get all PDF export settings in the project
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(ExportPDFSettings));
            
            foreach (ExportPDFSettings settings in collector)
            {
                var presetName = $"PDF: {settings.Name}";
                AddPDFNamingPreset(presetName, settings);
            }
        }
        catch (Exception ex)
        {
            // Log error but continue - PDF settings might not be available
            System.Diagnostics.Debug.WriteLine($"Error loading PDF naming presets: {ex.Message}");
        }
    }
    
    private void AddPDFNamingPreset(string presetName, ExportPDFSettings settings)
    {
        try
        {
            PDFExportOptions options = settings.GetOptions();
            IList<TableCellCombinedParameterData> namingRules = options.GetNamingRule();
            
            if (namingRules != null && namingRules.Count > 0)
            {
                string formatString = ConvertPDFNamingRuleToFormatString(namingRules);
                if (!string.IsNullOrEmpty(formatString))
                {
                    pdfNamingPresets[presetName] = formatString;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error processing PDF naming preset '{presetName}': {ex.Message}");
        }
    }
    
    private string ConvertPDFNamingRuleToFormatString(IList<TableCellCombinedParameterData> namingRules)
    {
        var formatParts = new List<string>();
        
        for (int i = 0; i < namingRules.Count; i++)
        {
            var rule = namingRules[i];
            var paramPart = "";
            
            // Add prefix if exists
            if (!string.IsNullOrEmpty(rule.Prefix))
            {
                formatParts.Add(rule.Prefix);
            }
            
            // Convert BuiltInParameter enum values to ElementId for comparison
            var sheetNumberId = new ElementId(BuiltInParameter.SHEET_NUMBER);
            var sheetNameId = new ElementId(BuiltInParameter.SHEET_NAME);
            var viewNameId = new ElementId(BuiltInParameter.VIEW_NAME);
            var viewTypeId = new ElementId(BuiltInParameter.VIEW_TYPE);
            var invalidId = new ElementId(BuiltInParameter.INVALID);
            
            // Get the parameter part
            if (rule.ParamId.Equals(sheetNumberId))
            {
                paramPart = "{Sheet Number}";
            }
            else if (rule.ParamId.Equals(sheetNameId))
            {
                paramPart = "{Sheet Name}";
            }
            else if (rule.ParamId.Equals(viewNameId))
            {
                paramPart = "{View Name}";
            }
            else if (rule.ParamId.Equals(viewTypeId))
            {
                paramPart = "{View Type}";
            }
            else if (!rule.ParamId.Equals(invalidId) && rule.ParamId.IntegerValue < 0)
            {
                // For built-in parameters (negative IDs), try to convert to BuiltInParameter
                try
                {
                    var builtInParam = (BuiltInParameter)rule.ParamId.IntegerValue;
                    var paramName = LabelUtils.GetLabelFor(builtInParam);
                    if (!string.IsNullOrEmpty(paramName))
                    {
                        paramPart = $"{{{paramName}}}";
                    }
                }
                catch
                {
                    // If conversion fails, try to get parameter name directly
                    var param = doc.GetElement(rule.ParamId) as ParameterElement;
                    if (param != null)
                    {
                        paramPart = $"{{{param.Name}}}";
                    }
                }
            }
            else if (rule.ParamId.IntegerValue > 0)
            {
                // For custom parameters (positive IDs), get the parameter element
                var param = doc.GetElement(rule.ParamId) as ParameterElement;
                if (param != null)
                {
                    paramPart = $"{{{param.Name}}}";
                }
            }
            
            // Add the parameter if found
            if (!string.IsNullOrEmpty(paramPart))
            {
                formatParts.Add(paramPart);
            }
            
            // Add suffix if exists
            if (!string.IsNullOrEmpty(rule.Suffix))
            {
                formatParts.Add(rule.Suffix);
            }
            
            // Add separator if specified and not the last item
            if (!string.IsNullOrEmpty(rule.Separator) && i < namingRules.Count - 1)
            {
                formatParts.Add(rule.Separator);
            }
        }
        
        return string.Join("", formatParts);
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

        // Sort the export settings alphabetically (case-insensitive)
        exportSettings = exportSettings.OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }
    
    private void LoadFormatHistory()
    {
        formatHistory = new List<string>();
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
                formatHistory.AddRange(lines.Where(l => !string.IsNullOrWhiteSpace(l)).Take(10));
            }
        }
        catch
        {
            // Use default if load fails
        }
        
        if (!formatHistory.Any())
            formatHistory.Add(formatString);
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
    
    private void SaveProjectSetting(string key, string value)
    {
        try
        {
            Directory.CreateDirectory(AppDataPath);
            var lines = File.Exists(ExportSettingsFile) ? File.ReadAllLines(ExportSettingsFile).ToList() : new List<string>();
            var projectName = doc.Title;
            var sectionHeader = $"[{projectName}]";
            int sectionIndex = lines.FindIndex(l => l.Trim() == sectionHeader);
            
            if (sectionIndex < 0)
            {
                lines.Add(sectionHeader);
                sectionIndex = lines.Count - 1;
            }
            
            int endIndex = sectionIndex + 1;
            while (endIndex < lines.Count && !lines[endIndex].Trim().StartsWith("["))
            {
                endIndex++;
            }
            
            bool found = false;
            for (int i = sectionIndex + 1; i < endIndex; i++)
            {
                var trimmed = lines[i].Trim();
                if (trimmed.StartsWith(key + ":"))
                {
                    lines[i] = $"{key}: {value}";
                    found = true;
                    break;
                }
            }
            
            if (!found)
            {
                lines.Insert(endIndex, $"{key}: {value}");
            }
            
            File.WriteAllLines(ExportSettingsFile, lines);
        }
        catch
        {
            // Ignore save errors
        }
    }
    
    private string GetProjectSetting(string key)
    {
        try
        {
            if (File.Exists(ExportSettingsFile))
            {
                var lines = File.ReadAllLines(ExportSettingsFile);
                var projectName = doc.Title;
                var sectionHeader = $"[{projectName}]";
                bool inSection = false;
                
                foreach (var line in lines)
                {
                    var trimmed = line.Trim();
                    if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                    {
                        inSection = trimmed == sectionHeader;
                        continue;
                    }
                    
                    if (inSection)
                    {
                        if (trimmed.StartsWith(key + ":"))
                        {
                            return trimmed.Substring(key.Length + 1).Trim();
                        }
                        else if (key == "export settings" && !trimmed.StartsWith("last path:") && !string.IsNullOrEmpty(trimmed))
                        {
                            return trimmed;
                        }
                    }
                }
            }
        }
        catch { }
        return null;
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
            Height = 130, // Increased height for new dropdown
            Padding = new Padding(12)
        };
        
        var lblFormat = new Label
        {
            Text = "Format String:",
            Location = new System.Drawing.Point(12, 12),
            Size = new System.Drawing.Size(100, 20)
        };
        
        cmbFormatString = new System.Windows.Forms.ComboBox
        {
            Location = new System.Drawing.Point(12, 35),
            Size = new System.Drawing.Size(600, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            DropDownStyle = ComboBoxStyle.DropDown
        };
        cmbFormatString.Items.AddRange(formatHistory.ToArray());
        cmbFormatString.Text = formatString;
        cmbFormatString.TextChanged += CmbFormatString_TextChanged;
        cmbFormatString.KeyDown += CmbFormatString_KeyDown;
        cmbFormatString.Enter += CmbFormatString_Enter;
        cmbFormatString.Leave += CmbFormatString_Leave;
        
        // New PDF naming presets dropdown
        var lblPDFPresets = new Label
        {
            Text = "PDF Export Format String:",
            Location = new System.Drawing.Point(12, 65),
            Size = new System.Drawing.Size(150, 20)
        };
        
        cmbPDFNamingPresets = new System.Windows.Forms.ComboBox
        {
            Location = new System.Drawing.Point(165, 65),
            Size = new System.Drawing.Size(300, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        
        // Populate PDF presets
        cmbPDFNamingPresets.Items.Add("<None>");
        foreach (var preset in pdfNamingPresets)
        {
            cmbPDFNamingPresets.Items.Add(preset.Key);
        }
        cmbPDFNamingPresets.SelectedIndex = 0;
        cmbPDFNamingPresets.SelectedIndexChanged += CmbPDFNamingPresets_SelectedIndexChanged;
        
        var lblExportOptions = new Label
        {
            Text = "DWG Export Settings:",
            Location = new System.Drawing.Point(12, 95),
            Size = new System.Drawing.Size(120, 20)
        };
        
        cmbExportOptions = new System.Windows.Forms.ComboBox
        {
            Location = new System.Drawing.Point(135, 95),
            Size = new System.Drawing.Size(300, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        
        // Populate export options
        cmbExportOptions.Items.Add("<Default>");
        foreach (var settings in exportSettings)
        {
            cmbExportOptions.Items.Add(settings.Name);
        }
        
        // Load saved choice
        string savedChoice = GetProjectSetting("export settings") ?? "<Default>";
        var selectedIndex = cmbExportOptions.Items.IndexOf(savedChoice);
        cmbExportOptions.SelectedIndex = selectedIndex >= 0 ? selectedIndex : 0;
        
        lblFormat.TabStop = false;
        cmbFormatString.TabIndex = 0; // First tab stop
        lblPDFPresets.TabStop = false;
        cmbPDFNamingPresets.TabIndex = 1;
        lblExportOptions.TabStop = false;
        cmbExportOptions.TabIndex = 2;
        
        pnlTop.Controls.AddRange(new System.Windows.Forms.Control[] {
            lblFormat, cmbFormatString,
            lblPDFPresets, cmbPDFNamingPresets,
            lblExportOptions, cmbExportOptions
        });
        
        pnlTop.Resize += (s, e) => {
            cmbFormatString.Width = pnlTop.ClientSize.Width - cmbFormatString.Left - 12;
        };
        
        // Bottom panel for buttons
        var pnlBottom = new System.Windows.Forms.Panel
        {
            Dock = DockStyle.Bottom,
            Height = 50,
            BackColor = SystemColors.Control
        };
        
        btnOK = new Button
        {
            Text = "OK",
            Size = new System.Drawing.Size(75, 30),
            DialogResult = DialogResult.OK,
            TabIndex = 6,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnOK.Click += BtnOK_Click;
        
        btnCancel = new Button
        {
            Text = "Cancel",
            Size = new System.Drawing.Size(75, 30),
            DialogResult = DialogResult.Cancel,
            TabIndex = 7,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        
        // Position buttons in resize event to ensure proper placement
        pnlBottom.Resize += (s, e) => {
            btnCancel.Location = new System.Drawing.Point(pnlBottom.Width - btnCancel.Width - 15, 10);
            btnOK.Location = new System.Drawing.Point(pnlBottom.Width - btnOK.Width - btnCancel.Width - 25, 10);
        };
        
        pnlBottom.Controls.Add(btnOK);
        pnlBottom.Controls.Add(btnCancel);
        
        // Main split container
        splitMain = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            SplitterDistance = 300, // 30% for parameters
            TabStop = false, // Prevent tab focus on splitter
            SplitterWidth = 4
        };
        
        // Left panel for parameters
        pnlLeft = new System.Windows.Forms.Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12, 12, 6, 12)
        };
        
        var lblAvailable = new Label
        {
            Text = "Available Parameters:",
            Location = new System.Drawing.Point(0, 0),
            Size = new System.Drawing.Size(200, 20)
        };
        
        txtSearch = new System.Windows.Forms.TextBox
        {
            Location = new System.Drawing.Point(0, 25),
            Size = new System.Drawing.Size(276, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            ForeColor = SystemColors.GrayText,
            Text = "Filter Parameters",
            TabIndex = 3,
            TabStop = true
        };
        txtSearch.TextChanged += TxtSearch_TextChanged;
        txtSearch.Enter += TxtSearch_Enter;
        txtSearch.Leave += TxtSearch_Leave;
        
        dgvAvailableParams = new DataGridView
        {
            Location = new System.Drawing.Point(0, 55),
            Size = new System.Drawing.Size(276, 400),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            ReadOnly = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            RowHeadersVisible = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            StandardTab = true,
            ScrollBars = ScrollBars.Vertical, // Changed to only vertical since columns are auto-sized
            RowTemplate = { Height = 18 }, // Set row height to 18 pixels
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersHeight = 23 // Set a fixed header height
        };
        dgvAvailableParams.DoubleClick += DgvAvailableParams_DoubleClick;
        dgvAvailableParams.KeyDown += DgvAvailableParams_KeyDown;
        
        // Setup columns
        dgvAvailableParams.Columns.Add("Parameter", "Parameter");
        dgvAvailableParams.Columns.Add("Sample", "Sample Value");
        dgvAvailableParams.Columns[0].FillWeight = 50;
        dgvAvailableParams.Columns[1].FillWeight = 50;
        
        PopulateParametersGrid();
        
        var btnAddParam = new Button
        {
            Text = "Insert Selected Parameter",
            Location = new System.Drawing.Point(0, 460),
            Size = new System.Drawing.Size(276, 25),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };
        btnAddParam.Click += BtnAddParam_Click;
        
        txtSearch.TabIndex = 3;
        lblAvailable.TabStop = false;
        dgvAvailableParams.TabIndex = 4;
        btnAddParam.TabIndex = 5;
        
        pnlLeft.Controls.AddRange(new System.Windows.Forms.Control[] {
            lblAvailable, txtSearch, dgvAvailableParams, btnAddParam
        });
        
        pnlLeft.Resize += (s, e) => {
            var padding = pnlLeft.Padding;
            var availableWidth = pnlLeft.ClientSize.Width - padding.Left - padding.Right;
            var availableHeight = pnlLeft.ClientSize.Height - padding.Top - padding.Bottom;
            
            // Update txtSearch width
            txtSearch.Width = availableWidth;
            
            // Update dgvAvailableParams size
            dgvAvailableParams.Width = availableWidth;
            dgvAvailableParams.Height = availableHeight - dgvAvailableParams.Top - 35; // Leave space for button
            
            // Update button position and width
            btnAddParam.Top = dgvAvailableParams.Bottom + 5;
            btnAddParam.Width = availableWidth;
        };
        
        splitMain.Panel1.Controls.Add(pnlLeft);
        
        // Right panel for preview
        var pnlRight = new System.Windows.Forms.Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(6, 12, 12, 12)
        };
        
        var lblPreview = new Label
        {
            Text = "Preview:",
            Location = new System.Drawing.Point(0, 0),
            Size = new System.Drawing.Size(100, 20)
        };
        
        dgvPreview = new DataGridView
        {
            Location = new System.Drawing.Point(0, 25),
            Size = new System.Drawing.Size(pnlRight.ClientSize.Width, pnlRight.ClientSize.Height - 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            RowHeadersVisible = false,
            ScrollBars = ScrollBars.Vertical,
            RowTemplate = { Height = 18 }, // Set row height to 18 pixels for consistency
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersHeight = 23 // Set a fixed header height
        };
        
        lblPreview.TabStop = false;
        dgvPreview.TabStop = false; // Preview is read-only, skip in tab order
        
        pnlRight.Controls.AddRange(new System.Windows.Forms.Control[] {
            lblPreview, dgvPreview
        });
        
        splitMain.Panel2.Controls.Add(pnlRight);
        
        // Add controls to form
        this.Controls.Add(splitMain);
        this.Controls.Add(pnlBottom);
        this.Controls.Add(pnlTop);
        
        // Set splitter position to 30% after form loads
        this.Load += (s, e) => { 
            splitMain.SplitterDistance = (int)(this.Width * 0.3);
            // Force initial resize
            pnlLeft.PerformLayout();
            pnlRight.PerformLayout();
        };
        
        // Set focus after form is shown
        this.Shown += (s, e) => {
            this.ActiveControl = cmbFormatString;
            cmbFormatString.Focus();
            cmbFormatString.SelectionStart = 0;
            cmbFormatString.SelectionLength = 0;
            lastCursorPosition = 0;
        };
        
        // Set accept and cancel buttons
        this.AcceptButton = btnOK;
        this.CancelButton = btnCancel;
        
        // Set tab order for panels
        pnlTop.TabIndex = 0;
        splitMain.TabIndex = 1;
        pnlBottom.TabIndex = 2;
    }
    
    private void CmbPDFNamingPresets_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (cmbPDFNamingPresets.SelectedIndex > 0)
        {
            var selectedPreset = cmbPDFNamingPresets.SelectedItem.ToString();
            if (pdfNamingPresets.ContainsKey(selectedPreset))
            {
                // Save current state for undo
                if (!string.IsNullOrEmpty(formatString))
                {
                    undoStack.Push(formatString);
                    redoStack.Clear();
                }
                
                formatString = pdfNamingPresets[selectedPreset];
                cmbFormatString.Text = formatString;
                lastCursorPosition = formatString.Length;
                UpdatePreview();
            }
        }
    }
    
    private void PopulateParametersGrid()
    {
        dgvAvailableParams.Rows.Clear();
        
        // Get first view for sample values
        var firstView = views.FirstOrDefault();
        var sampleParams = firstView != null ? viewParameters[firstView] : new Dictionary<string, string>();
        
        var parametersToShow = isPlaceholderActive || string.IsNullOrWhiteSpace(txtSearch.Text) 
            ? availableParameters 
            : availableParameters.Where(p => p.ToLower().Contains(txtSearch.Text.ToLower())).ToList();
        
        foreach (var param in parametersToShow)
        {
            var sampleValue = sampleParams.ContainsKey(param) ? sampleParams[param] : "";
            if (sampleValue.Length > 50)
                sampleValue = sampleValue.Substring(0, 47) + "...";
            
            dgvAvailableParams.Rows.Add(param, sampleValue);
        }
    }
    
    private void CmbFormatString_TextChanged(object sender, EventArgs e)
    {
        if (formatString != cmbFormatString.Text)
        {
            // Save current state for undo before changing
            if (!string.IsNullOrEmpty(formatString))
            {
                undoStack.Push(formatString);
                redoStack.Clear(); // Clear redo stack when new change is made
            }
            
            formatString = cmbFormatString.Text;
            UpdatePreview();
        }
    }
    
    private void CmbFormatString_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.F4)
        {
            cmbFormatString.DroppedDown = !cmbFormatString.DroppedDown;
            e.Handled = true;
            cmbFormatString.BeginInvoke(new Action(() =>
            {
                cmbFormatString.SelectionStart = lastCursorPosition;
                cmbFormatString.SelectionLength = 0;
            }));
        }
        else if (e.Control && e.KeyCode == Keys.Z)
        {
            // Undo
            if (undoStack.Count > 0)
            {
                redoStack.Push(formatString);
                formatString = undoStack.Pop();
                cmbFormatString.Text = formatString;
                lastCursorPosition = formatString.Length;
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        else if (e.Control && e.KeyCode == Keys.Y)
        {
            // Redo
            if (redoStack.Count > 0)
            {
                undoStack.Push(formatString);
                formatString = redoStack.Pop();
                cmbFormatString.Text = formatString;
                lastCursorPosition = formatString.Length;
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
    }
    
    private void CmbFormatString_Enter(object sender, EventArgs e)
    {
        cmbFormatString.BeginInvoke(new Action(() =>
        {
            cmbFormatString.SelectionStart = lastCursorPosition;
            cmbFormatString.SelectionLength = 0;
        }));
    }
    
    private void CmbFormatString_Leave(object sender, EventArgs e)
    {
        lastCursorPosition = cmbFormatString.SelectionStart;
    }
    
    private void TxtSearch_TextChanged(object sender, EventArgs e)
    {
        // Skip filtering if showing placeholder
        if (isPlaceholderActive)
            return;
            
        PopulateParametersGrid();
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
            PopulateParametersGrid();
        }
    }
    
    private void BtnAddParam_Click(object sender, EventArgs e)
    {
        InsertSelectedParameter();
    }
    
    private void DgvAvailableParams_DoubleClick(object sender, EventArgs e)
    {
        InsertSelectedParameter();
    }
    
    private void DgvAvailableParams_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            InsertSelectedParameter();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }
        else if (e.KeyCode == Keys.Tab)
        {
            // Let the form handle tab navigation
            e.Handled = false;
        }
    }
    
    private void InsertSelectedParameter()
    {
        if (dgvAvailableParams.SelectedRows.Count > 0)
        {
            var paramName = dgvAvailableParams.SelectedRows[0].Cells[0].Value.ToString();
            var insertText = $"{{{paramName}}}";
            
            // Save current state for undo
            undoStack.Push(formatString);
            redoStack.Clear();
            
            var selectionStart = lastCursorPosition;
            var currentText = cmbFormatString.Text;
            cmbFormatString.Text = currentText.Insert(selectionStart, insertText);
            lastCursorPosition = selectionStart + insertText.Length;
            cmbFormatString.Focus();
            cmbFormatString.SelectionStart = lastCursorPosition;
            cmbFormatString.SelectionLength = 0;
        }
    }
    
    private void BtnOK_Click(object sender, EventArgs e)
    {
        SaveFormatHistory();
        SaveProjectSetting("export settings", cmbExportOptions.SelectedItem?.ToString() ?? "<Default>");
    }
    
    public string GetLastExportPath()
    {
        return GetProjectSetting("last path") ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    }
    
    public void SaveExportSettings(string exportPath)
    {
        SaveProjectSetting("export settings", cmbExportOptions.SelectedItem?.ToString() ?? "<Default>");
        SaveProjectSetting("last path", exportPath);
    }
    
    private void UpdatePreview()
    {
        dgvPreview.Rows.Clear();
        dgvPreview.Columns.Clear();
        
        dgvPreview.Columns.Add("Original", "Original Name");
        dgvPreview.Columns.Add("NewName", "DWG File Name");
        
        // Set column widths
        dgvPreview.Columns[0].FillWeight = 30;
        dgvPreview.Columns[1].FillWeight = 70;
        
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

// Progress dialog for export operation
public class ExportProgressDialog : System.Windows.Forms.Form
{
    private ProgressBar progressBar;
    private Label lblStatus;
    private Label lblProgress;
    private Button btnCancel;
    private bool isCancelled = false;
    private int totalItems;
    
    public bool IsCancelled => isCancelled;
    
    public ExportProgressDialog(int totalItems)
    {
        this.totalItems = totalItems;
        InitializeUI();
    }
    
    private void InitializeUI()
    {
        this.Text = "Exporting to DWG";
        this.Size = new System.Drawing.Size(450, 180); // Increased height
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.ControlBox = false; // Remove close button
        
        lblStatus = new Label
        {
            Text = "Preparing export...",
            Location = new System.Drawing.Point(12, 20),
            Size = new System.Drawing.Size(410, 20),
            AutoEllipsis = true
        };
        
        lblProgress = new Label
        {
            Text = "0 of " + totalItems,
            Location = new System.Drawing.Point(12, 45),
            Size = new System.Drawing.Size(100, 20)
        };
        
        progressBar = new ProgressBar
        {
            Location = new System.Drawing.Point(12, 70),
            Size = new System.Drawing.Size(410, 23),
            Minimum = 0,
            Maximum = totalItems,
            Value = 0,
            Style = ProgressBarStyle.Continuous
        };
        
        btnCancel = new Button
        {
            Text = "Cancel",
            Location = new System.Drawing.Point(347, 105), // Moved up slightly
            Size = new System.Drawing.Size(75, 23),
            DialogResult = DialogResult.Cancel
        };
        btnCancel.Click += (s, e) => { 
            isCancelled = true;
            btnCancel.Enabled = false;
            btnCancel.Text = "Cancelling...";
        };
        
        this.Controls.AddRange(new System.Windows.Forms.Control[] {
            lblStatus, lblProgress, progressBar, btnCancel
        });
        
        // Add ESC key handling
        this.KeyPreview = true;
        this.KeyDown += (s, e) => {
            if (e.KeyCode == Keys.Escape)
            {
                btnCancel.PerformClick();
            }
        };
    }
    
    public void UpdateProgress(int current, string status)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => UpdateProgress(current, status)));
            return;
        }
        
        progressBar.Value = Math.Min(current + 1, progressBar.Maximum);
        lblStatus.Text = status;
        lblProgress.Text = $"{current + 1} of {totalItems}";
        // Update percentage in title
        var percentage = (int)((current + 1) * 100.0 / totalItems);
        this.Text = $"Exporting to DWG - {percentage}%";
        
        this.Refresh();
    }
}
