// Revit 2024 API – external command to numerically change the camera focal length of a
// *perspective* 3‑D view. Works whether the crop box is displayed or not: the
// command temporarily enables it (if it was off), scales it, then restores the
// original display state.
// -----------------------------------------------------------------------------
// Author: ChatGPT – July 2025
// -----------------------------------------------------------------------------

#region Namespaces
using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WF = System.Windows.Forms; // alias WinForms to avoid API name clashes
#endregion

namespace SampleCommands
{
    [Transaction(TransactionMode.Manual)]
    public class SetFocalLength : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc   = uiApp.ActiveUIDocument;
            Document   doc     = uiDoc.Document;

            // ---------------- ensure perspective 3‑D view --------------------------------
            View3D v3d = uiDoc.ActiveView as View3D;
            if (v3d == null)
            {
                TaskDialog.Show("Set Focal Length", "The active view is not a 3‑D view.");
                return Result.Cancelled;
            }
            if (!v3d.IsPerspective)
            {
                TaskDialog.Show("Set Focal Length", "The active 3‑D view is orthographic.\nFocal length applies only to perspective views.");
                return Result.Cancelled;
            }

            // ---------------- determine current focal length -----------------------------
            double currentFocal = CameraUtils.CalculateFocalLengthMm(v3d);
            if (double.IsNaN(currentFocal) || currentFocal <= 0)
            {
                TaskDialog.Show("Set Focal Length", "Could not determine the current focal length.");
                return Result.Failed;
            }

            // ---------------- ask the user ----------------------------------------------
            using (FocalLengthForm dlg = new FocalLengthForm(currentFocal))
            {
                if (dlg.ShowDialog() != WF.DialogResult.OK)
                    return Result.Cancelled;

                double desiredFocal = dlg.NewFocalLengthMm;
                if (Math.Abs(desiredFocal - currentFocal) < 0.01)
                    return Result.Cancelled; // unchanged

                bool wasCropActive = v3d.CropBoxActive; // remember original state

                using (Transaction tx = new Transaction(doc, "Set Focal Length"))
                {
                    tx.Start();

                    if (!wasCropActive)
                        v3d.CropBoxActive = true; // temporarily enable it

                    CameraUtils.SetFocalLengthMm(v3d, desiredFocal);

                    if (!wasCropActive)
                        v3d.CropBoxActive = false; // restore original state

                    tx.Commit();
                }
            }
            return Result.Succeeded;
        }
    }

    // -------------------------------------------------------------------------
    // Utility helpers – conversion between focal length and crop‑box geometry.
    // -------------------------------------------------------------------------
    internal static class CameraUtils
    {
        private const double FilmWidthMm = 36.0; // matches Revit navigation wheel

        internal static double CalculateFocalLengthMm(View3D v3d)
        {
            BoundingBoxXYZ crop = v3d.CropBox; // always present for 3‑D views
            if (crop == null) return double.NaN;

            Transform t = crop.Transform;

            // crop‑box width in its local X (horizontal) direction – internal units (ft)
            XYZ half = (crop.Max - crop.Min) * 0.5;
            double widthFt = 2 * half.X;

            // distance from eye to crop‑center – internal units (ft)
            XYZ eye = v3d.GetOrientation().EyePosition;
            XYZ centreLocal = (crop.Max + crop.Min) * 0.5;
            XYZ centreWorld = t.OfPoint(centreLocal);
            double distFt = eye.DistanceTo(centreWorld);

            // convert to mm
            double widthMm = UnitUtils.ConvertFromInternalUnits(widthFt, UnitTypeId.Millimeters);
            double distMm = UnitUtils.ConvertFromInternalUnits(distFt, UnitTypeId.Millimeters);

            double halfAngle = Math.Atan((widthMm / 2.0) / distMm);
            double focalMm  = (FilmWidthMm / 2.0) / Math.Tan(halfAngle);
            return focalMm;
        }

        internal static void SetFocalLengthMm(View3D v3d, double desiredFocalMm)
        {
            double currentFocal = CalculateFocalLengthMm(v3d);
            if (double.IsNaN(currentFocal) || currentFocal == 0) return;

            double scale = currentFocal / desiredFocalMm; // >1 ⇒ zoom‑in, <1 ⇒ zoom‑out

            BoundingBoxXYZ crop = v3d.CropBox;
            if (crop == null) return;

            Transform t = crop.Transform;

            XYZ centre = (crop.Max + crop.Min) * 0.5;
            XYZ half   = (crop.Max - crop.Min) * 0.5;
            half = half.Multiply(scale);

            BoundingBoxXYZ newBox = new BoundingBoxXYZ();
            newBox.Transform = t;
            newBox.Min = centre - half;
            newBox.Max = centre + half;

            v3d.CropBox = newBox; // apply
        }
    }

    // -------------------------------------------------------------------------
    // WinForms dialog – aliased types keep names clear of Revit classes.
    // -------------------------------------------------------------------------
    internal sealed class FocalLengthForm : WF.Form
    {
        private readonly WF.TextBox _txt;
        internal double NewFocalLengthMm { get; private set; }

        internal FocalLengthForm(double currentFocalMm)
        {
            // basic form set‑up
            this.Text = "Set Focal Length (mm)";
            this.FormBorderStyle = WF.FormBorderStyle.FixedDialog;
            this.StartPosition  = WF.FormStartPosition.CenterParent;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.ClientSize = new System.Drawing.Size(230, 90);

            // label
            WF.Label lbl = new WF.Label();
            lbl.Text = "Focal length:";
            lbl.Left = 10;
            lbl.Top = 15;
            lbl.AutoSize = true;

            // textbox
            _txt = new WF.TextBox();
            _txt.Left = 100;
            _txt.Top = 12;
            _txt.Width = 100;
            _txt.Text = currentFocalMm.ToString("F2");

            // buttons
            WF.Button ok = new WF.Button();
            ok.Text = "OK";
            ok.DialogResult = WF.DialogResult.OK;
            ok.Left = 35;
            ok.Width = 70;
            ok.Top = 50;

            WF.Button cancel = new WF.Button();
            cancel.Text = "Cancel";
            cancel.DialogResult = WF.DialogResult.Cancel;
            cancel.Left = 125;
            cancel.Width = 70;
            cancel.Top = 50;

            this.AcceptButton = ok;
            this.CancelButton = cancel;

            this.Controls.AddRange(new WF.Control[] { lbl, _txt, ok, cancel });

            ok.Click += OnOkClick;
        }

        private void OnOkClick(object sender, EventArgs e)
        {
            double newVal;
            if (!double.TryParse(_txt.Text, out newVal) || newVal <= 0)
            {
                WF.MessageBox.Show("Please enter a positive number.", "Invalid input", WF.MessageBoxButtons.OK, WF.MessageBoxIcon.Warning);
                this.DialogResult = WF.DialogResult.None; // keep form open
            }
            else
            {
                NewFocalLengthMm = newVal;
            }
        }
    }
}
