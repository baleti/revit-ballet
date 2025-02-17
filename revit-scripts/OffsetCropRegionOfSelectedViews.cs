using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture; // For Viewport
using Autodesk.Revit.UI;

// Alias Windows Forms namespace to avoid conflicts with Revit types.
using WinForms = System.Windows.Forms;

namespace OffsetCropRegion
{
    [Transaction(TransactionMode.Manual)]
    public class OffsetCropRegionOfSelectedViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Gather selected views. If a viewport is selected, get its associated view.
            List<Autodesk.Revit.DB.View> selectedViews = new List<Autodesk.Revit.DB.View>();
            foreach (var id in uidoc.Selection.GetElementIds())
            {
                Element elem = doc.GetElement(id);
                if (elem is Autodesk.Revit.DB.View view)
                {
                    selectedViews.Add(view);
                }
                else if (elem is Viewport viewport)
                {
                    // Get the view associated with the viewport.
                    Autodesk.Revit.DB.View vpView = doc.GetElement(viewport.ViewId) as Autodesk.Revit.DB.View;
                    if (vpView != null)
                        selectedViews.Add(vpView);
                }
            }
            // Remove duplicate views.
            selectedViews = selectedViews.Distinct().ToList();

            if (!selectedViews.Any())
            {
                TaskDialog.Show("Offset Crop Region", "No valid views (or viewports) selected.");
                return Result.Cancelled;
            }

            // Determine if we can offer 4 separate offset values.
            // We require that all selected views have a rectangular crop region (exactly 4 segments).
            bool allRectangular = selectedViews.All(v => IsRectangular(v));

            // Use document title as the project identifier.
            string projectName = doc.Title;

            // Load stored settings (if any) for this command.
            Dictionary<string, OffsetCropRegionSettings> settingsDict = SettingsManager.LoadSettings();
            OffsetCropRegionSettings defaultSettings;
            if (settingsDict.ContainsKey(projectName))
            {
                defaultSettings = settingsDict[projectName];
            }
            else
            {
                // Provide defaults: 50 mm for uniform and for each side.
                defaultSettings = new OffsetCropRegionSettings
                {
                    UseFourOffsets = allRectangular,
                    UniformOffset = 50,
                    TopOffset = 50,
                    LeftOffset = 50,
                    BottomOffset = 50,
                    RightOffset = 50
                };
            }
            // If not all views support rectangular crop regions, force uniform mode.
            if (!allRectangular)
                defaultSettings.UseFourOffsets = false;

            // Show the dialog.
            OffsetInputForm form = new OffsetInputForm(defaultSettings);
            WinForms.DialogResult dr = form.ShowDialog();
            if (dr != WinForms.DialogResult.OK)
            {
                return Result.Cancelled;
            }

            // Retrieve offset values from dialog and convert mm -> feet (1 ft = 304.8 mm).
            double mmToFeet = 1.0 / 304.8;
            bool useFourOffsets = defaultSettings.UseFourOffsets;
            double uniformOffsetFeet = form.UniformOffset * mmToFeet;
            double topOffsetFeet = form.TopOffset * mmToFeet;
            double leftOffsetFeet = form.LeftOffset * mmToFeet;
            double bottomOffsetFeet = form.BottomOffset * mmToFeet;
            double rightOffsetFeet = form.RightOffset * mmToFeet;

            // Start a transaction to update the crop regions.
            using (Transaction tx = new Transaction(doc, "Offset Crop Region"))
            {
                tx.Start();
                foreach (Autodesk.Revit.DB.View view in selectedViews)
                {
                    // Process only if the view's crop box is active.
                    if (!view.CropBoxActive)
                        continue;
                    BoundingBoxXYZ bbox = view.CropBox;
                    if (bbox == null)
                        continue;

                    // Determine if this view’s crop region is rectangular.
                    bool rectangular = IsRectangular(view);

                    XYZ newMin, newMax;
                    if (rectangular)
                    {
                        if (useFourOffsets)
                        {
                            newMin = new XYZ(bbox.Min.X - leftOffsetFeet,
                                             bbox.Min.Y - bottomOffsetFeet,
                                             bbox.Min.Z);
                            newMax = new XYZ(bbox.Max.X + rightOffsetFeet,
                                             bbox.Max.Y + topOffsetFeet,
                                             bbox.Max.Z);
                        }
                        else
                        {
                            // Uniform offset applied on all sides.
                            newMin = new XYZ(bbox.Min.X - uniformOffsetFeet,
                                             bbox.Min.Y - uniformOffsetFeet,
                                             bbox.Min.Z);
                            newMax = new XYZ(bbox.Max.X + uniformOffsetFeet,
                                             bbox.Max.Y + uniformOffsetFeet,
                                             bbox.Max.Z);
                        }
                    }
                    else
                    {
                        // For nonrectangular crop regions, apply uniform offset.
                        newMin = new XYZ(bbox.Min.X - uniformOffsetFeet,
                                         bbox.Min.Y - uniformOffsetFeet,
                                         bbox.Min.Z);
                        newMax = new XYZ(bbox.Max.X + uniformOffsetFeet,
                                         bbox.Max.Y + uniformOffsetFeet,
                                         bbox.Max.Z);
                    }
                    // Update the crop box.
                    bbox.Min = newMin;
                    bbox.Max = newMax;
                    view.CropBox = bbox;
                }
                tx.Commit();
            }

            // Update and save the settings.
            if (settingsDict.ContainsKey(projectName))
                settingsDict[projectName] = form.GetSettings();
            else
                settingsDict.Add(projectName, form.GetSettings());
            SettingsManager.SaveSettings(settingsDict);

            return Result.Succeeded;
        }

        /// <summary>
        /// Checks whether the view’s crop region is rectangular.
        /// For views of type ViewPlan we try to retrieve the crop shape;
        /// for other view types, we assume a rectangular crop region.
        /// </summary>
        private bool IsRectangular(Autodesk.Revit.DB.View view)
        {
            if (view == null || view.CropBox == null)
                return false;

            // For this example, if the view is a ViewPlan, attempt to use reflection to get CropRegionShapeManager.
            if (view is Autodesk.Revit.DB.ViewPlan plan)
            {
                try
                {
                    var managerProperty = plan.GetType().GetProperty("CropRegionShapeManager");
                    if (managerProperty != null)
                    {
                        object manager = managerProperty.GetValue(plan, null);
                        if (manager != null)
                        {
                            var method = manager.GetType().GetMethod("GetCropRegionShape");
                            if (method != null)
                            {
                                CurveLoop cropLoop = method.Invoke(manager, null) as CurveLoop;
                                if (cropLoop != null)
                                {
                                    int count = cropLoop.ToList().Count;
                                    return count == 4;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    return false;
                }
            }
            // For other view types, assume rectangular.
            return true;
        }
    }

    #region Settings Management

    /// <summary>
    /// Represents the offset settings.
    /// </summary>
    public class OffsetCropRegionSettings
    {
        // If true, then separate offsets for Top, Left, Bottom, Right are used.
        public bool UseFourOffsets { get; set; }
        public double UniformOffset { get; set; }
        public double TopOffset { get; set; }
        public double LeftOffset { get; set; }
        public double BottomOffset { get; set; }
        public double RightOffset { get; set; }
    }

    /// <summary>
    /// Handles loading and saving settings to a plain text file.
    /// </summary>
    public static class SettingsManager
    {
        private static string GetSettingsFilePath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "revit-scripts");
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            return Path.Combine(folder, "OffsetCropRegionOfSelectedViews.txt");
        }

        /// <summary>
        /// Loads settings from the text file.
        /// Each line represents a project's settings in the format:
        /// ProjectName|UseFourOffsets|UniformOffset|TopOffset|LeftOffset|BottomOffset|RightOffset
        /// </summary>
        public static Dictionary<string, OffsetCropRegionSettings> LoadSettings()
        {
            Dictionary<string, OffsetCropRegionSettings> settingsDict = new Dictionary<string, OffsetCropRegionSettings>();
            string path = GetSettingsFilePath();
            if (!File.Exists(path))
                return settingsDict;

            try
            {
                foreach (string line in File.ReadAllLines(path))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] tokens = line.Split('|');
                    if (tokens.Length != 7)
                        continue;

                    string projectName = tokens[0];
                    if (!bool.TryParse(tokens[1], out bool useFourOffsets))
                        useFourOffsets = false;
                    if (!double.TryParse(tokens[2], out double uniformOffset))
                        uniformOffset = 50;
                    if (!double.TryParse(tokens[3], out double topOffset))
                        topOffset = 50;
                    if (!double.TryParse(tokens[4], out double leftOffset))
                        leftOffset = 50;
                    if (!double.TryParse(tokens[5], out double bottomOffset))
                        bottomOffset = 50;
                    if (!double.TryParse(tokens[6], out double rightOffset))
                        rightOffset = 50;

                    OffsetCropRegionSettings settings = new OffsetCropRegionSettings
                    {
                        UseFourOffsets = useFourOffsets,
                        UniformOffset = uniformOffset,
                        TopOffset = topOffset,
                        LeftOffset = leftOffset,
                        BottomOffset = bottomOffset,
                        RightOffset = rightOffset
                    };
                    settingsDict[projectName] = settings;
                }
            }
            catch
            {
                // On error, return an empty dictionary.
            }
            return settingsDict;
        }

        /// <summary>
        /// Saves the settings dictionary to the text file using the pipe-delimited format.
        /// </summary>
        public static void SaveSettings(Dictionary<string, OffsetCropRegionSettings> settingsDict)
        {
            string path = GetSettingsFilePath();
            try
            {
                List<string> lines = new List<string>();
                foreach (var kvp in settingsDict)
                {
                    OffsetCropRegionSettings s = kvp.Value;
                    string line = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}",
                        kvp.Key,
                        s.UseFourOffsets,
                        s.UniformOffset,
                        s.TopOffset,
                        s.LeftOffset,
                        s.BottomOffset,
                        s.RightOffset);
                    lines.Add(line);
                }
                File.WriteAllLines(path, lines);
            }
            catch
            {
                // Ignore errors.
            }
        }
    }

    #endregion

    #region Dialog Form

    /// <summary>
    /// A Windows Form dialog for entering offset values.
    /// If using four offsets (when crop region is rectangular), the dialog shows a field labeled
    /// "Offset All (mm)" that updates four individual fields (Top, Left, Bottom, Right) live.
    /// If not, it shows only one "Offset All (mm)" field.
    /// </summary>
    public class OffsetInputForm : WinForms.Form
    {
        // Mode flag: if using four offsets.
        private bool _useFourOffsets;

        // Controls for uniform mode.
        private WinForms.Label _labelUniform;
        private WinForms.TextBox _textBoxUniform;

        // Controls for four-offset mode.
        private WinForms.Label _labelGlobal;
        private WinForms.TextBox _textBoxGlobal;

        private WinForms.Label _labelTop;
        private WinForms.TextBox _textBoxTop;
        private WinForms.Label _labelLeft;
        private WinForms.TextBox _textBoxLeft;
        private WinForms.Label _labelBottom;
        private WinForms.TextBox _textBoxBottom;
        private WinForms.Label _labelRight;
        private WinForms.TextBox _textBoxRight;

        private WinForms.Button _buttonOk;
        private WinForms.Button _buttonCancel;

        private OffsetCropRegionSettings _settings;

        /// <summary>
        /// In uniform mode, returns the offset value entered (in mm).
        /// In four-offset mode, this value is ignored.
        /// </summary>
        public double UniformOffset
        {
            get
            {
                double.TryParse(_textBoxUniform?.Text, out double val);
                return val;
            }
        }
        public double TopOffset
        {
            get
            {
                double.TryParse(_textBoxTop?.Text, out double val);
                return val;
            }
        }
        public double LeftOffset
        {
            get
            {
                double.TryParse(_textBoxLeft?.Text, out double val);
                return val;
            }
        }
        public double BottomOffset
        {
            get
            {
                double.TryParse(_textBoxBottom?.Text, out double val);
                return val;
            }
        }
        public double RightOffset
        {
            get
            {
                double.TryParse(_textBoxRight?.Text, out double val);
                return val;
            }
        }

        public OffsetInputForm(OffsetCropRegionSettings settings)
        {
            _settings = settings;
            _useFourOffsets = settings.UseFourOffsets;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Offset Crop Region";
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            this.StartPosition = WinForms.FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            if (_useFourOffsets)
            {
                // Increase form height to accommodate the extra "Offset All" field plus individual fields.
                this.Width = 320;
                this.Height = 300;

                // "Offset All" field (formerly Global Offset) for live updates.
                _labelGlobal = new WinForms.Label() { Left = 10, Top = 10, Width = 120, Text = "Offset All (mm):" };
                // If all four individual settings are equal, use that value; otherwise, leave blank.
                string globalText = (_settings.TopOffset == _settings.LeftOffset &&
                                     _settings.TopOffset == _settings.BottomOffset &&
                                     _settings.TopOffset == _settings.RightOffset)
                                     ? _settings.TopOffset.ToString() : "";
                _textBoxGlobal = new WinForms.TextBox() { Left = 140, Top = 10, Width = 150, Text = globalText };
                _textBoxGlobal.TextChanged += GlobalTextBox_TextChanged;
                this.Controls.Add(_labelGlobal);
                this.Controls.Add(_textBoxGlobal);

                // Four labels and textboxes.
                _labelTop = new WinForms.Label() { Left = 10, Top = 50, Width = 100, Text = "Top Offset (mm):" };
                _textBoxTop = new WinForms.TextBox() { Left = 140, Top = 50, Width = 150, Text = _settings.TopOffset.ToString() };

                _labelLeft = new WinForms.Label() { Left = 10, Top = 80, Width = 100, Text = "Left Offset (mm):" };
                _textBoxLeft = new WinForms.TextBox() { Left = 140, Top = 80, Width = 150, Text = _settings.LeftOffset.ToString() };

                _labelBottom = new WinForms.Label() { Left = 10, Top = 110, Width = 100, Text = "Bottom Offset (mm):" };
                _textBoxBottom = new WinForms.TextBox() { Left = 140, Top = 110, Width = 150, Text = _settings.BottomOffset.ToString() };

                _labelRight = new WinForms.Label() { Left = 10, Top = 140, Width = 100, Text = "Right Offset (mm):" };
                _textBoxRight = new WinForms.TextBox() { Left = 140, Top = 140, Width = 150, Text = _settings.RightOffset.ToString() };

                this.Controls.Add(_labelTop);
                this.Controls.Add(_textBoxTop);
                this.Controls.Add(_labelLeft);
                this.Controls.Add(_textBoxLeft);
                this.Controls.Add(_labelBottom);
                this.Controls.Add(_textBoxBottom);
                this.Controls.Add(_labelRight);
                this.Controls.Add(_textBoxRight);
            }
            else
            {
                // Uniform mode: only one field, labeled "Offset All (mm)".
                this.Width = 320;
                this.Height = 160;
                _labelUniform = new WinForms.Label() { Left = 10, Top = 20, Width = 120, Text = "Offset All (mm):" };
                _textBoxUniform = new WinForms.TextBox() { Left = 140, Top = 20, Width = 150, Text = _settings.UniformOffset.ToString() };
                this.Controls.Add(_labelUniform);
                this.Controls.Add(_textBoxUniform);
            }

            // OK and Cancel buttons.
            int btnTop = _useFourOffsets ? 220 : 60;
            _buttonOk = new WinForms.Button() { Text = "OK", Left = 90, Width = 80, Top = btnTop, DialogResult = WinForms.DialogResult.OK };
            _buttonCancel = new WinForms.Button() { Text = "Cancel", Left = 180, Width = 80, Top = btnTop, DialogResult = WinForms.DialogResult.Cancel };

            this.Controls.Add(_buttonOk);
            this.Controls.Add(_buttonCancel);
            this.AcceptButton = _buttonOk;
            this.CancelButton = _buttonCancel;
        }

        /// <summary>
        /// When the "Offset All" textbox changes, update all four individual offset textboxes.
        /// </summary>
        private void GlobalTextBox_TextChanged(object sender, EventArgs e)
        {
            string text = _textBoxGlobal.Text;
            _textBoxTop.Text = text;
            _textBoxLeft.Text = text;
            _textBoxBottom.Text = text;
            _textBoxRight.Text = text;
        }

        /// <summary>
        /// Returns a settings object containing the values entered by the user.
        /// </summary>
        public OffsetCropRegionSettings GetSettings()
        {
            if (_useFourOffsets)
            {
                return new OffsetCropRegionSettings
                {
                    UseFourOffsets = true,
                    TopOffset = this.TopOffset,
                    LeftOffset = this.LeftOffset,
                    BottomOffset = this.BottomOffset,
                    RightOffset = this.RightOffset,
                    UniformOffset = 0  // Not used in this mode.
                };
            }
            else
            {
                return new OffsetCropRegionSettings
                {
                    UseFourOffsets = false,
                    UniformOffset = this.UniformOffset,
                    TopOffset = 0,
                    LeftOffset = 0,
                    BottomOffset = 0,
                    RightOffset = 0
                };
            }
        }
    }

    #endregion
}
