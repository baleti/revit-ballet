// MoveSelectedToCentroid.cs  – Revit 2024 / .NET 4.8 / C# 7.3
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WF = System.Windows.Forms;

namespace MoveSelectedToCentroid
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MoveSelectedToCentroid : IExternalCommand
    {
        private static readonly string StoreDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "revit-scripts");
        private static readonly string StoreFile = Path.Combine(StoreDir, "MoveSelectedToCentroid");

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uidoc.Document;
            ICollection<ElementId> selIds = uidoc.GetSelectionIds();

            if (selIds == null || selIds.Count == 0)
            {
                message = "Please select one or more elements before running the command.";
                return Result.Failed;
            }

            // ------------------------------------------------------------------
            // Detect model units  ------------------------------------------------
            Units projectUnits = doc.GetUnits();
            FormatOptions lengthOpts = projectUnits.GetFormatOptions(SpecTypeId.Length);
            bool isMetric = lengthOpts.GetUnitTypeId() == UnitTypeId.Millimeters
                         || lengthOpts.GetUnitTypeId() == UnitTypeId.Meters;
            double convToFeet = isMetric ? 1.0 / 304.8 /* mm → ft */ : 1.0;

            // ------------------------------------------------------------------
            // Check if every selected element is view-specific  -----------------
            bool onlyViewDependent = selIds.All(id =>
            {
                Element el = doc.GetElement(id);
                return el != null && (el.ViewSpecific || el is View);
            });

            // ------------------------------------------------------------------
            // Load last values (if any)  ----------------------------------------
            double lastX = 0, lastY = 0, lastZ = 0;
            if (File.Exists(StoreFile))
            {
                string[] parts = File.ReadAllText(StoreFile).Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 3)
                {
                    double.TryParse(parts[0], out lastX);
                    double.TryParse(parts[1], out lastY);
                    double.TryParse(parts[2], out lastZ);
                }
            }

            // ------------------------------------------------------------------
            // Show the coordinate dialog  ---------------------------------------
            XYZ target;
            using (var frm = new CoordinateInputForm(
                includeZ: !onlyViewDependent,
                showMetric: isMetric,
                defX: lastX,
                defY: lastY,
                defZ: lastZ))
            {
                if (frm.ShowDialog() != WF.DialogResult.OK)
                    return Result.Cancelled;

                // Convert to feet if the user typed millimetres
                target = new XYZ(frm.X * convToFeet,
                                 frm.Y * convToFeet,
                                 (!onlyViewDependent ? frm.Z * convToFeet : 0));
                // Persist current (raw) values for next time
                try
                {
                    if (!Directory.Exists(StoreDir)) Directory.CreateDirectory(StoreDir);
                    File.WriteAllText(StoreFile, $"{frm.X};{frm.Y};{frm.Z}");
                }
                catch { /* silent — any IO problem is non-fatal */ }
            }

            // ------------------------------------------------------------------
            // Move the elements  ------------------------------------------------
            using (Transaction tx = new Transaction(doc, "Move Selected To Centroid"))
            {
                tx.Start();

                foreach (ElementId id in selIds)
                {
                    Element el = doc.GetElement(id);

                    // Viewport on a sheet – use sheet coordinates, ignore Z
                    if (el is Viewport vp)
                    {
                        XYZ current   = vp.GetBoxCenter();
                        XYZ newCenter = new XYZ(target.X, target.Y, current.Z);
                        vp.SetBoxCenter(newCenter);
                        continue;
                    }

                    // Handle Views (must be placed on sheets)
                    if (el is View view)
                    {
                        // Find if this view is placed on any sheet
                        Viewport viewport = FindViewportForView(doc, view.Id);
                        
                        if (viewport != null)
                        {
                            // View is placed on a sheet, move its viewport
                            XYZ current = viewport.GetBoxCenter();
                            XYZ newCenter = new XYZ(target.X, target.Y, current.Z);
                            viewport.SetBoxCenter(newCenter);
                        }
                        // Skip views not placed on sheets (as per requirements)
                        continue;
                    }

                    // Other elements
                    BoundingBoxXYZ bb = el.get_BoundingBox(null);
                    if (bb == null) continue;

                    XYZ centroid = (bb.Min + bb.Max) * 0.5;
                    XYZ moveVec  = target - centroid;

                    try
                    {
                        ElementTransformUtils.MoveElement(doc, id, moveVec);
                    }
                    catch (Autodesk.Revit.Exceptions.ArgumentException)
                    {
                        // pinned / constrained – silently skip
                    }
                }

                tx.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Finds the viewport that contains the specified view
        /// </summary>
        /// <param name="doc">The current document</param>
        /// <param name="viewId">The ID of the view to find</param>
        /// <returns>The viewport displaying the view, or null if the view isn't placed on any sheet</returns>
        private Viewport FindViewportForView(Document doc, ElementId viewId)
        {
            // Find all viewports in the document
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Viewport));
            
            // Look for a viewport that references our view
            foreach (Viewport vp in collector)
            {
                if (vp.ViewId.Equals(viewId))
                    return vp;
            }
            
            return null;
        }
    }

    /// <summary>
    /// Modal dialog for entering target coords.
    /// </summary>
    internal class CoordinateInputForm : WF.Form
    {
        private readonly bool _includeZ;
        private readonly WF.TextBox _txtX = new WF.TextBox();
        private readonly WF.TextBox _txtY = new WF.TextBox();
        private readonly WF.TextBox _txtZ = new WF.TextBox();

        public double X { get; private set; }
        public double Y { get; private set; }
        public double Z { get; private set; }

        public CoordinateInputForm(bool includeZ, bool showMetric,
                                   double defX, double defY, double defZ)
        {
            _includeZ = includeZ;
            string unitTag = showMetric ? " (mm)" : " (ft)";

            Text            = "Target Coordinates" + unitTag;
            FormBorderStyle = WF.FormBorderStyle.FixedDialog;
            StartPosition   = WF.FormStartPosition.CenterParent;
            Width           = 260;
            Height          = includeZ ? 190 : 150;
            MaximizeBox     = false;
            MinimizeBox     = false;

            var lblX = new WF.Label { Left = 10, Top = 18, Width = 30, Text = "X:" };
            var lblY = new WF.Label { Left = 10, Top = 48, Width = 30, Text = "Y:" };
            var lblZ = new WF.Label { Left = 10, Top = 78, Width = 30, Text = "Z:" };

            _txtX.SetBounds(45, 15, 185, 20);
            _txtY.SetBounds(45, 45, 185, 20);
            _txtZ.SetBounds(45, 75, 185, 20);

            _txtX.Text = defX.ToString();
            _txtY.Text = defY.ToString();
            _txtZ.Text = defZ.ToString();

            var btnOK = new WF.Button
            {
                Text         = "OK",
                Left         = 55,
                Width        = 60,
                Top          = includeZ ? 110 : 80,
                DialogResult = WF.DialogResult.OK
            };
            var btnCancel = new WF.Button
            {
                Text         = "Cancel",
                Left         = 135,
                Width        = 60,
                Top          = includeZ ? 110 : 80,
                DialogResult = WF.DialogResult.Cancel
            };

            btnOK.Click += (s, e) =>
            {
                if (ValidateAndStore()) DialogResult = WF.DialogResult.OK;
            };

            Controls.AddRange(new WF.Control[]
            {
                lblX, lblY,
                _txtX, _txtY,
                btnOK, btnCancel
            });

            if (includeZ)
            {
                Controls.AddRange(new WF.Control[] { lblZ, _txtZ });
            }

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }

        private bool ValidateAndStore()
        {
            if (!double.TryParse(_txtX.Text, out double x) ||
                !double.TryParse(_txtY.Text, out double y) ||
                (_includeZ && !double.TryParse(_txtZ.Text, out double z)))
            {
                WF.MessageBox.Show(
                    "Please enter valid numeric values (use '.' for decimals).",
                    "Invalid Input",
                    WF.MessageBoxButtons.OK,
                    WF.MessageBoxIcon.Warning);
                return false;
            }

            X = x;
            Y = y;
            Z = _includeZ ? double.Parse(_txtZ.Text) : 0;
            return true;
        }
    }
}
