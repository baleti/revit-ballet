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
using System.Reflection;
using System.Threading;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ExportSelectedViewsToPDF : IExternalCommand
{
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetWindowText(IntPtr hWnd, string lpString);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    private const uint BM_CLICK = 0x00F5;

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = commandData.Application.ActiveUIDocument;
        var doc = uidoc.Document;
        
        // Get current selection
        var selectedIds = uidoc.Selection.GetElementIds();
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
        using (var dialog = new PDFNamingDialog(doc, viewsAndSheets))
        {
            if (dialog.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;
            
            // Get folder location
            string exportFolder = null;
            var lastPath = dialog.GetLastExportPath();
            
            var folderDialog = new CommonOpenFileDialog
            {
                Title = "Select folder for PDF export",
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
            
            // Save settings
            dialog.SaveExportSettings(exportFolder);
            
            // Get base export settings (null for default)
            var baseSettings = dialog.GetSelectedExportSettings();
            string selectedPrinter = dialog.GetSelectedPrinter();
            
            var successCount = 0;
            var failedExports = new List<string>();
            var cancelled = false;

            bool useNativePDF = Type.GetType("Autodesk.Revit.DB.PDFExportOptions, RevitAPI") != null;
            
            using (var progressDialog = new PDFExportProgressDialog(viewsAndSheets.Count))
            {
                progressDialog.Show(new PDFRevitWindow(commandData.Application.MainWindowHandle));
                
                using (var tx = new Transaction(doc, "Export Views to PDF"))
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
                        var fileName = dialog.GetFileName(view);
                        progressDialog.UpdateProgress(i, $"Exporting: {fileName}.pdf");
                        Application.DoEvents();
                        
                        try
                        {
                            var viewIds = new List<ElementId> { view.Id };

                            if (useNativePDF)
                            {
                                Type pdfOptionsType = Type.GetType("Autodesk.Revit.DB.PDFExportOptions, RevitAPI");
                                dynamic options = Activator.CreateInstance(pdfOptionsType);

                                if (baseSettings != null)
                                {
                                    options = baseSettings.GetOptions();
                                }

                                options.FileName = fileName;
                                options.Combine = true;

                                MethodInfo exportMethod = typeof(Document).GetMethod("Export", new Type[] { typeof(string), typeof(IList<ElementId>), pdfOptionsType });
                                exportMethod.Invoke(doc, new object[] { exportFolder, viewIds, options });
                            }
                            else
                            {
                                // Fallback to printing
                                PrintManager pm = doc.PrintManager;
                                pm.PrintToFile = true;
                                string fullPath = Path.Combine(exportFolder, fileName + ".pdf");
                                pm.PrintToFileName = fullPath;
                                pm.SelectNewPrintDriver(selectedPrinter ?? "Microsoft Print to PDF");
                                pm.Apply();
                                pm.SubmitPrint(view);

                                // Automate the save dialog
                                IntPtr dlg = IntPtr.Zero;
                                int attempts = 0;
                                while (dlg == IntPtr.Zero && attempts < 50)
                                {
                                    dlg = FindWindow("#32770", "Save Print Output As");
                                    Thread.Sleep(100);
                                    attempts++;
                                }

                                if (dlg != IntPtr.Zero)
                                {
                                    IntPtr edit = FindWindowEx(dlg, IntPtr.Zero, "Edit", null);
                                    if (edit != IntPtr.Zero)
                                    {
                                        SetWindowText(edit, fullPath);
                                    }

                                    IntPtr saveBtn = FindWindowEx(dlg, IntPtr.Zero, "Button", "&Save");
                                    if (saveBtn != IntPtr.Zero)
                                    {
                                        SendMessage(saveBtn, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                                    }
                                }
                            }
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
                Autodesk.Revit.UI.TaskDialog.Show("Export Cancelled", "PDF export was cancelled by user.");
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

public class PDFRevitWindow : IWin32Window
{
    public IntPtr Handle { get; private set; }
    public PDFRevitWindow(IntPtr handle)
    {
        Handle = handle;
    }
}

public class PDFNamingDialog : System.Windows.Forms.Form
{
    private static readonly string AppDataPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "revit-scripts"
    );
    private static readonly string FormatHistoryFile = Path.Combine(AppDataPath, "ExportSelectedViewsToPDF");
    private static readonly string ExportSettingsFile = Path.Combine(AppDataPath, "ExportSelectedViewsToPDF_Settings");
    
    private static readonly bool IsNativePDFAvailable = Type.GetType("Autodesk.Revit.DB.PDFExportOptions, RevitAPI") != null;
    
    private Document doc;
    private List<Autodesk.Revit.DB.View> views;
    private Dictionary<Autodesk.Revit.DB.View, Dictionary<string, string>> viewParameters;
    private List<string> availableParameters;
    private string formatString = "{Sheet Number}_{Sheet Name}";
    private List<dynamic> exportSettings;
    private bool isPlaceholderActive = true;
    private List<string> formatHistory;
    private int lastCursorPosition = 0;
    
    private System.Windows.Forms.ComboBox cmbFormatString;
    private System.Windows.Forms.TextBox txtSearch;
    private DataGridView dgvAvailableParams;
    private DataGridView dgvPreview;
    private System.Windows.Forms.ComboBox cmbExportOptions;
    private System.Windows.Forms.ComboBox cmbPrinterDriver;
    private Button btnOK;
    private Button btnCancel;
    private SplitContainer splitMain;
    private System.Windows.Forms.Panel pnlTop;
    private System.Windows.Forms.Panel pnlLeft;
    
    public PDFNamingDialog(Document doc, List<Autodesk.Revit.DB.View> views)
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
        exportSettings = new List<dynamic>();
        
        Type pdfSettingsType = Type.GetType("Autodesk.Revit.DB.ExportPDFSettings, RevitAPI");
        if (pdfSettingsType != null)
        {
            var collector = new FilteredElementCollector(doc).OfClass(pdfSettingsType);
            
            foreach (dynamic settings in collector)
            {
                exportSettings.Add(settings);
            }
        }
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
        this.Text = "Configure PDF Export Naming";
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
        
        var lblExportOptions = new Label
        {
            Text = "PDF Export Settings:",
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
        foreach (dynamic settings in exportSettings)
        {
            cmbExportOptions.Items.Add(settings.Name);
        }
        
        // Load saved choice
        string savedChoice = GetProjectSetting("export settings") ?? "<Default>";
        var selectedIndex = cmbExportOptions.Items.IndexOf(savedChoice);
        cmbExportOptions.SelectedIndex = selectedIndex >= 0 ? selectedIndex : 0;
        
        lblFormat.TabStop = false;
        cmbFormatString.TabIndex = 0; // First tab stop
        lblExportOptions.TabStop = false;
        cmbExportOptions.TabIndex = 1;
        
        pnlTop.Controls.AddRange(new System.Windows.Forms.Control[] {
            lblFormat, cmbFormatString,
            lblExportOptions, cmbExportOptions
        });

        if (!IsNativePDFAvailable)
        {
            var lblPrinter = new Label
            {
                Text = "PDF Printer Driver:",
                Location = new System.Drawing.Point(12, 95),
                Size = new System.Drawing.Size(120, 20)
            };
            
            cmbPrinterDriver = new System.Windows.Forms.ComboBox
            {
                Location = new System.Drawing.Point(135, 95),
                Size = new System.Drawing.Size(300, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            
            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                cmbPrinterDriver.Items.Add(printer);
            }
            
            string savedPrinter = GetProjectSetting("printer driver");
            if (!string.IsNullOrEmpty(savedPrinter) && cmbPrinterDriver.Items.Contains(savedPrinter))
            {
                cmbPrinterDriver.SelectedItem = savedPrinter;
            }
            else if (cmbPrinterDriver.Items.Count > 0)
            {
                cmbPrinterDriver.SelectedIndex = 0;
            }
            
            lblPrinter.TabStop = false;
            cmbPrinterDriver.TabIndex = 2;
            
            pnlTop.Controls.Add(lblPrinter);
            pnlTop.Controls.Add(cmbPrinterDriver);
            pnlTop.Height += 40; // Increase height for the new control
        }
        
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
            TabIndex = 5,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnOK.Click += BtnOK_Click;
        
        btnCancel = new Button
        {
            Text = "Cancel",
            Size = new System.Drawing.Size(75, 30),
            DialogResult = DialogResult.Cancel,
            TabIndex = 6,
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
        
        txtSearch.TabIndex = 2;
        lblAvailable.TabStop = false;
        dgvAvailableParams.TabIndex = 3;
        btnAddParam.TabIndex = 4;
        
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
        formatString = cmbFormatString.Text;
        UpdatePreview();
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
        if (cmbPrinterDriver != null)
        {
            SaveProjectSetting("printer driver", cmbPrinterDriver.SelectedItem?.ToString());
        }
    }
    
    public string GetLastExportPath()
    {
        return GetProjectSetting("last path") ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    }
    
    public void SaveExportSettings(string exportPath)
    {
        SaveProjectSetting("export settings", cmbExportOptions.SelectedItem?.ToString() ?? "<Default>");
        if (cmbPrinterDriver != null)
        {
            SaveProjectSetting("printer driver", cmbPrinterDriver.SelectedItem?.ToString());
        }
        SaveProjectSetting("last path", exportPath);
    }
    
    private void UpdatePreview()
    {
        dgvPreview.Rows.Clear();
        dgvPreview.Columns.Clear();
        
        dgvPreview.Columns.Add("Original", "Original Name");
        dgvPreview.Columns.Add("NewName", "PDF File Name");
        
        // Set column widths
        dgvPreview.Columns[0].FillWeight = 30;
        dgvPreview.Columns[1].FillWeight = 70;
        
        foreach (var view in views.Take(15)) // Show first 15 for performance
        {
            var originalName = view.Name;
            var newName = GetFileName(view) + ".pdf"; // Add .pdf for preview
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
    
    public dynamic GetSelectedExportSettings()
    {
        if (cmbExportOptions.SelectedIndex <= 0)
            return null;
        
        var settingsName = cmbExportOptions.SelectedItem.ToString();
        return exportSettings.FirstOrDefault(s => s.Name == settingsName);
    }
    
    public string GetSelectedPrinter()
    {
        return cmbPrinterDriver?.SelectedItem?.ToString();
    }
}

public class PDFExportProgressDialog : System.Windows.Forms.Form
{
    private ProgressBar progressBar;
    private Label lblStatus;
    private Label lblProgress;
    private Button btnCancel;
    private bool isCancelled = false;
    private int totalItems;
    
    public bool IsCancelled => isCancelled;
    
    public PDFExportProgressDialog(int totalItems)
    {
        this.totalItems = totalItems;
        InitializeUI();
    }
    
    private void InitializeUI()
    {
        this.Text = "Exporting to PDF";
        this.Size = new System.Drawing.Size(650, 180); // Wider dialog
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.ControlBox = false; // Remove close button
        
        lblStatus = new Label
        {
            Text = "Preparing export...",
            Location = new System.Drawing.Point(12, 20),
            Size = new System.Drawing.Size(610, 20),
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
            Size = new System.Drawing.Size(610, 23),
            Minimum = 0,
            Maximum = totalItems,
            Value = 0,
            Style = ProgressBarStyle.Continuous
        };
        
        btnCancel = new Button
        {
            Text = "Cancel",
            Location = new System.Drawing.Point(547, 105), // Adjusted for wider dialog
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
        this.Text = $"Exporting to PDF - {percentage}%";
        
        this.Refresh();
    }
}
