using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ChangeLengthsOfSelectedLevels : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = commandData.Application.ActiveUIDocument;
        var doc = uidoc.Document;
        var activeView = doc.ActiveView;
        
        try
        {
            // Get current selection using SelectionModeManager
            var selectedIds = uidoc.GetSelectionIds();
            
            // Filter for Level elements and Views
            var levels = new List<Level>();
            var viewsSet = new HashSet<Autodesk.Revit.DB.View>();
            
            foreach (var id in selectedIds)
            {
                var element = doc.GetElement(id);
                if (element is Level level)
                {
                    levels.Add(level);
                }
                else if (element is Autodesk.Revit.DB.View view && 
                         !(view is ViewSheet) && 
                         !(view is ViewSchedule) &&
                         !(view is View3D) &&
                         !(view is ViewDrafting))
                {
                    // Add view directly if it's a plan, section, or elevation
                    viewsSet.Add(view);
                }
                else if (element is Viewport viewport)
                {
                    // Get view from viewport
                    var viewFromViewport = doc.GetElement(viewport.ViewId) as Autodesk.Revit.DB.View;
                    if (viewFromViewport != null && 
                        !(viewFromViewport is ViewSchedule) &&
                        !(viewFromViewport is View3D) &&
                        !(viewFromViewport is ViewDrafting))
                    {
                        viewsSet.Add(viewFromViewport);
                    }
                }
            }
            
            var views = viewsSet.ToList();
            
            if (!levels.Any())
            {
                TaskDialog.Show("No Levels Selected", 
                    "Please select one or more level elements before running this command.");
                return Result.Cancelled;
            }
            
            // If no views selected, use active view
            if (!views.Any())
            {
                views.Add(activeView);
            }
            
            // Get project units
            var units = doc.GetUnits();
            var lengthFormatOptions = units.GetFormatOptions(SpecTypeId.Length);
            var unitTypeId = lengthFormatOptions.GetUnitTypeId();
            
            // Determine if metric or imperial
            bool isMetric = IsMetricUnit(unitTypeId);
            string unitLabel = GetUnitLabel(unitTypeId);
            
            // Check bubble visibility for the levels across all target views
            bool end0HasBubble = false;
            bool end1HasBubble = false;
            foreach (var view in views)
            {
                foreach (var level in levels)
                {
                    try
                    {
                        if (level.IsBubbleVisibleInView(DatumEnds.End0, view))
                            end0HasBubble = true;
                        if (level.IsBubbleVisibleInView(DatumEnds.End1, view))
                            end1HasBubble = true;
                    }
                    catch
                    {
                        // Level might not be visible in this view
                        continue;
                    }
                }
            }
            
            // Show input dialog
            using (var form = new LevelLengthInputForm(isMetric, unitLabel, end0HasBubble, end1HasBubble))
            {
                if (form.ShowDialog() != WinForms.DialogResult.OK)
                {
                    return Result.Cancelled;
                }
                
                // Convert input values from display units to internal units (feet)
                double startExtension = UnitUtils.ConvertToInternalUnits(
                    form.StartEndValue, unitTypeId);
                double endExtension = UnitUtils.ConvertToInternalUnits(
                    form.EndEndValue, unitTypeId);
                
                // Modify levels in a transaction
                using (var tx = new Transaction(doc, "Change Level Lengths"))
                {
                    tx.Start();
                    
                    var modifiedViews = new Dictionary<string, int>();
                    
                    foreach (var view in views)
                    {
                        int modifiedInView = 0;
                        foreach (var level in levels)
                        {
                            if (ModifyLevelLength(level, view, startExtension, endExtension))
                            {
                                modifiedInView++;
                            }
                        }
                        if (modifiedInView > 0)
                        {
                            modifiedViews[view.Name] = modifiedInView;
                        }
                    }
                    
                    tx.Commit();
                    
                    // Show dialog only if multiple views were modified or if views were explicitly selected
                    bool viewsWereExplicitlySelected = selectedIds.Any(id => 
                    {
                        var elem = doc.GetElement(id);
                        return elem is Autodesk.Revit.DB.View || elem is Viewport;
                    });
                    
                    if (viewsWereExplicitlySelected && modifiedViews.Any())
                    {
                        string viewList = string.Join("\n", modifiedViews.Select(kvp => 
                            $"  • {kvp.Key}: {kvp.Value} level(s)"));
                        
                        TaskDialog.Show("Levels Modified", 
                            $"Modified levels in the following views:\n\n{viewList}");
                    }
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
    
    private bool ModifyLevelLength(Level level, Autodesk.Revit.DB.View view, double startExtension, double endExtension)
    {
        try
        {
            // Get the level's curve - first try view-specific
            var curves = level.GetCurvesInView(DatumExtentType.ViewSpecific, view);
            Curve curve = curves?.FirstOrDefault();
            
            // If no view-specific curve, try model extent
            if (curve == null)
            {
                curves = level.GetCurvesInView(DatumExtentType.Model, view);
                curve = curves?.FirstOrDefault();
            }
            
            if (curve == null)
            {
                return false;
            }
            
            // Only handle straight lines for now
            if (!(curve is Line line))
            {
                return false;
            }
            
            // Get current endpoints
            var startPoint = line.GetEndPoint(0);
            var endPoint = line.GetEndPoint(1);
            
            // Get direction vector
            var direction = (endPoint - startPoint).Normalize();
            
            // Calculate new endpoints
            // End 0 corresponds to start point, End 1 to end point
            // Negative values will shorten, positive will extend
            var newStartPoint = startPoint - direction * startExtension;
            var newEndPoint = endPoint + direction * endExtension;
            
            // Create new line
            var newLine = Line.CreateBound(newStartPoint, newEndPoint);
            
            // Set the new curve
            level.SetCurveInView(DatumExtentType.ViewSpecific, view, newLine);
            
            return true;
        }
        catch
        {
            return false;
        }
    }
    
    private bool IsMetricUnit(ForgeTypeId unitTypeId)
    {
        // Check if the unit is metric
        return unitTypeId == UnitTypeId.Millimeters ||
               unitTypeId == UnitTypeId.Centimeters ||
               unitTypeId == UnitTypeId.Decimeters ||
               unitTypeId == UnitTypeId.Meters;
    }
    
    private string GetUnitLabel(ForgeTypeId unitTypeId)
    {
        if (unitTypeId == UnitTypeId.Feet) return "ft";
        if (unitTypeId == UnitTypeId.FeetFractionalInches) return "ft";
        if (unitTypeId == UnitTypeId.Inches) return "in";
        if (unitTypeId == UnitTypeId.FractionalInches) return "in";
        if (unitTypeId == UnitTypeId.Millimeters) return "mm";
        if (unitTypeId == UnitTypeId.Centimeters) return "cm";
        if (unitTypeId == UnitTypeId.Decimeters) return "dm";
        if (unitTypeId == UnitTypeId.Meters) return "m";
        return "";
    }
}

// WinForm for user input
public class LevelLengthInputForm : WinForms.Form
{
    private WinForms.TextBox txtStartEnd;
    private WinForms.TextBox txtEndEnd;
    private WinForms.Button btnOk;
    private WinForms.Button btnCancel;
    private WinForms.Label lblStartEnd;
    private WinForms.Label lblEndEnd;
    
    public double StartEndValue { get; private set; }
    public double EndEndValue { get; private set; }
    
    public LevelLengthInputForm(bool isMetric, string unitLabel, bool end0HasBubble, bool end1HasBubble)
    {
        InitializeComponent(isMetric, unitLabel, end0HasBubble, end1HasBubble);
    }
    
    private void InitializeComponent(bool isMetric, string unitLabel, bool end0HasBubble, bool end1HasBubble)
    {
        this.Text = "Change Level Lengths";
        this.Width = 320;
        this.Height = 160;
        this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.StartPosition = WinForms.FormStartPosition.CenterScreen;
        
        // Determine label text based on bubble visibility
        string end0Label = $"End 0{(end0HasBubble ? " (with bubble)" : "")} ({unitLabel}):";
        string end1Label = $"End 1{(end1HasBubble ? " (with bubble)" : "")} ({unitLabel}):";
        
        // Start end input
        lblStartEnd = new WinForms.Label
        {
            Text = end0Label,
            Location = new Drawing.Point(12, 20),
            Size = new Drawing.Size(180, 23),
            TextAlign = Drawing.ContentAlignment.MiddleLeft
        };
        
        txtStartEnd = new WinForms.TextBox
        {
            Location = new Drawing.Point(195, 20),
            Size = new Drawing.Size(100, 23),
            Text = "0"
        };
        
        // End end input
        lblEndEnd = new WinForms.Label
        {
            Text = end1Label,
            Location = new Drawing.Point(12, 50),
            Size = new Drawing.Size(180, 23),
            TextAlign = Drawing.ContentAlignment.MiddleLeft
        };
        
        txtEndEnd = new WinForms.TextBox
        {
            Location = new Drawing.Point(195, 50),
            Size = new Drawing.Size(100, 23),
            Text = "0"
        };
        
        // OK button
        btnOk = new WinForms.Button
        {
            Text = "OK",
            Location = new Drawing.Point(139, 90),
            Size = new Drawing.Size(75, 30),
            DialogResult = WinForms.DialogResult.OK
        };
        btnOk.Click += BtnOk_Click;
        
        // Cancel button
        btnCancel = new WinForms.Button
        {
            Text = "Cancel",
            Location = new Drawing.Point(220, 90),
            Size = new Drawing.Size(75, 30),
            DialogResult = WinForms.DialogResult.Cancel
        };
        
        // Add controls
        this.Controls.AddRange(new WinForms.Control[] 
        {
            lblStartEnd,
            txtStartEnd,
            lblEndEnd,
            txtEndEnd,
            btnOk,
            btnCancel
        });
        
        // Set accept and cancel buttons
        this.AcceptButton = btnOk;
        this.CancelButton = btnCancel;
        
        // Focus on first input
        txtStartEnd.Select();
    }
    
    private void BtnOk_Click(object sender, EventArgs e)
    {
        // Validate inputs
        if (!double.TryParse(txtStartEnd.Text, out double startValue))
        {
            WinForms.MessageBox.Show("Please enter a valid number for Start End.", 
                "Invalid Input", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
            txtStartEnd.Focus();
            txtStartEnd.SelectAll();
            this.DialogResult = WinForms.DialogResult.None;
            return;
        }
        
        if (!double.TryParse(txtEndEnd.Text, out double endValue))
        {
            WinForms.MessageBox.Show("Please enter a valid number for End End.", 
                "Invalid Input", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
            txtEndEnd.Focus();
            txtEndEnd.SelectAll();
            this.DialogResult = WinForms.DialogResult.None;
            return;
        }
        
        StartEndValue = startValue;
        EndEndValue = endValue;
    }
}
