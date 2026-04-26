using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
[CommandMeta("Grid, Level")]
public class ChangeLength : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = commandData.Application.ActiveUIDocument;
        var doc = uidoc.Document;
        var activeView = doc.ActiveView;

        try
        {
            var selectedIds = uidoc.GetSelectionIds();

            var datums = new List<DatumPlane>();
            var viewsSet = new HashSet<Autodesk.Revit.DB.View>();

            foreach (var id in selectedIds)
            {
                var element = doc.GetElement(id);
                if (element is Grid || element is Level)
                {
                    datums.Add((DatumPlane)element);
                }
                else if (element is Autodesk.Revit.DB.View view &&
                         !(view is ViewSheet) &&
                         !(view is ViewSchedule) &&
                         !(view is View3D) &&
                         !(view is ViewDrafting))
                {
                    viewsSet.Add(view);
                }
                else if (element is Viewport viewport)
                {
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

            if (!datums.Any())
            {
                CustomGUIs.SetCurrentUIDocument(uidoc);
                var items = new List<Dictionary<string, object>>();
                foreach (Grid g in new FilteredElementCollector(doc).OfClass(typeof(Grid)).Cast<Grid>())
                    items.Add(new Dictionary<string, object> { { "Name", g.Name }, { "Type", "Grid" }, { "ElementIdObject", g.Id } });
                foreach (Level l in new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>())
                    items.Add(new Dictionary<string, object> { { "Name", l.Name }, { "Type", "Level" }, { "ElementIdObject", l.Id } });

                var chosen = CustomGUIs.DataGrid(items, new List<string> { "Name", "Type" }, false);
                if (chosen == null || chosen.Count == 0) return Result.Cancelled;

                foreach (var row in chosen)
                {
                    if (row.TryGetValue("ElementIdObject", out object idObj) && idObj is ElementId eid)
                        if (doc.GetElement(eid) is DatumPlane dp) datums.Add(dp);
                }
                if (!datums.Any()) return Result.Succeeded;
            }

            var views = viewsSet.Any() ? viewsSet.ToList() : new List<Autodesk.Revit.DB.View> { activeView };

            var units = doc.GetUnits();
            var unitTypeId = units.GetFormatOptions(SpecTypeId.Length).GetUnitTypeId();
            string unitLabel = GetUnitLabel(unitTypeId);

            bool end0HasBubble = false;
            bool end1HasBubble = false;
            foreach (var view in views)
            {
                foreach (var datum in datums)
                {
                    try
                    {
                        if (datum.IsBubbleVisibleInView(DatumEnds.End0, view)) end0HasBubble = true;
                        if (datum.IsBubbleVisibleInView(DatumEnds.End1, view)) end1HasBubble = true;
                    }
                    catch { }
                }
            }

            using (var form = new DatumLengthInputForm(unitLabel, end0HasBubble, end1HasBubble))
            {
                if (form.ShowDialog() != WinForms.DialogResult.OK)
                    return Result.Cancelled;

                double startExtension = UnitUtils.ConvertToInternalUnits(form.StartEndValue, unitTypeId);
                double endExtension   = UnitUtils.ConvertToInternalUnits(form.EndEndValue,   unitTypeId);

                using (var tx = new Transaction(doc, "Change Lengths"))
                {
                    tx.Start();

                    var modifiedViews = new Dictionary<string, int>();
                    foreach (var view in views)
                    {
                        int count = 0;
                        foreach (var datum in datums)
                            if (ModifyDatumLength(datum, view, startExtension, endExtension)) count++;
                        if (count > 0) modifiedViews[view.Name] = count;
                    }

                    tx.Commit();

                    bool viewsExplicit = selectedIds.Any(id =>
                    {
                        var elem = doc.GetElement(id);
                        return elem is Autodesk.Revit.DB.View || elem is Viewport;
                    });

                    if (viewsExplicit && modifiedViews.Any())
                    {
                        string viewList = string.Join("\n", modifiedViews.Select(kvp => $"  - {kvp.Key}: {kvp.Value} element(s)"));
                        TaskDialog.Show("Lengths Modified", $"Modified elements in the following views:\n\n{viewList}");
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

    private bool ModifyDatumLength(DatumPlane datum, Autodesk.Revit.DB.View view, double startExt, double endExt)
    {
        try
        {
            Curve curve = datum.GetCurvesInView(DatumExtentType.ViewSpecific, view)?.FirstOrDefault();

            if (curve == null)
            {
                if (datum is Grid grid)
                    curve = grid.Curve;
                else if (datum is Level level)
                    curve = level.GetCurvesInView(DatumExtentType.Model, view)?.FirstOrDefault();
            }

            if (curve == null || !(curve is Line line)) return false;

            var start = line.GetEndPoint(0);
            var end   = line.GetEndPoint(1);
            var dir   = (end - start).Normalize();

            var newLine = Line.CreateBound(start - dir * startExt, end + dir * endExt);
            datum.SetCurveInView(DatumExtentType.ViewSpecific, view, newLine);
            return true;
        }
        catch { return false; }
    }

    private static string GetUnitLabel(ForgeTypeId unitTypeId)
    {
        if (unitTypeId == UnitTypeId.Feet || unitTypeId == UnitTypeId.FeetFractionalInches) return "ft";
        if (unitTypeId == UnitTypeId.Inches || unitTypeId == UnitTypeId.FractionalInches)   return "in";
        if (unitTypeId == UnitTypeId.Millimeters)  return "mm";
        if (unitTypeId == UnitTypeId.Centimeters)  return "cm";
        if (unitTypeId == UnitTypeId.Decimeters)   return "dm";
        if (unitTypeId == UnitTypeId.Meters)       return "m";
        return "";
    }
}

public class DatumLengthInputForm : WinForms.Form
{
    private WinForms.TextBox txtStart, txtEnd;

    public double StartEndValue { get; private set; }
    public double EndEndValue   { get; private set; }

    public DatumLengthInputForm(string unitLabel, bool end0HasBubble, bool end1HasBubble)
    {
        Text = "Change Lengths";
        Width = 320; Height = 160;
        FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        MaximizeBox = false; MinimizeBox = false;
        StartPosition = WinForms.FormStartPosition.CenterScreen;

        string lbl0 = $"End 0{(end0HasBubble ? " (with bubble)" : "")} ({unitLabel}):";
        string lbl1 = $"End 1{(end1HasBubble ? " (with bubble)" : "")} ({unitLabel}):";

        Controls.Add(new WinForms.Label { Text = lbl0, Location = new Drawing.Point(12, 20),  Size = new Drawing.Size(180, 23), TextAlign = Drawing.ContentAlignment.MiddleLeft });
        Controls.Add(new WinForms.Label { Text = lbl1, Location = new Drawing.Point(12, 50),  Size = new Drawing.Size(180, 23), TextAlign = Drawing.ContentAlignment.MiddleLeft });

        txtStart = new WinForms.TextBox { Location = new Drawing.Point(195, 20), Size = new Drawing.Size(100, 23), Text = "0" };
        txtEnd   = new WinForms.TextBox { Location = new Drawing.Point(195, 50), Size = new Drawing.Size(100, 23), Text = "0" };
        Controls.Add(txtStart);
        Controls.Add(txtEnd);

        var btnOk     = new WinForms.Button { Text = "OK",     Location = new Drawing.Point(139, 90), Size = new Drawing.Size(75, 30), DialogResult = WinForms.DialogResult.OK };
        var btnCancel = new WinForms.Button { Text = "Cancel", Location = new Drawing.Point(220, 90), Size = new Drawing.Size(75, 30), DialogResult = WinForms.DialogResult.Cancel };
        Controls.Add(btnOk);
        Controls.Add(btnCancel);
        AcceptButton = btnOk;
        CancelButton = btnCancel;

        btnOk.Click += (s, e) =>
        {
            if (!double.TryParse(txtStart.Text, out double sv))
            {
                WinForms.MessageBox.Show("Please enter a valid number for End 0.", "Invalid Input", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                txtStart.Focus(); txtStart.SelectAll();
                DialogResult = WinForms.DialogResult.None;
                return;
            }
            if (!double.TryParse(txtEnd.Text, out double ev))
            {
                WinForms.MessageBox.Show("Please enter a valid number for End 1.", "Invalid Input", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                txtEnd.Focus(); txtEnd.SelectAll();
                DialogResult = WinForms.DialogResult.None;
                return;
            }
            StartEndValue = sv;
            EndEndValue   = ev;
        };

        txtStart.Select();
    }
}

#endif
