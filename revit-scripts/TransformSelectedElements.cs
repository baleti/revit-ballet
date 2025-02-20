#region Namespaces
using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using WinForms = System.Windows.Forms;
#endregion

namespace TransformSelectedElementsSample
{
  public enum TransformationAxis { X, Y, Z }

  public class TransformationSettings
  {
    // Move parameters:
    public bool DoMove { get; set; }
    public double MoveX { get; set; }
    public double MoveY { get; set; }
    public double MoveZ { get; set; }

    // Mirror parameters (using checkboxes):
    public bool DoMirror { get; set; }
    public bool MirrorX { get; set; }
    public bool MirrorY { get; set; }
    public bool MirrorZ { get; set; }  // Only available for non–view–dependent elements.

    // Rotate parameters:
    public bool DoRotate { get; set; }
    public double RotationAngle { get; set; }
    public TransformationAxis RotationAxis { get; set; }
  }

  [Transaction(TransactionMode.Manual)]
  public class TransformSelectedElements : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Document doc = uidoc.Document;

      var selIds = uidoc.Selection.GetElementIds();
      if (selIds.Count == 0)
      {
        TaskDialog.Show("Error", "Please select one or more elements to transform.");
        return Result.Failed;
      }

      // Determine if all selected elements are view–dependent.
      bool onlyViewDependent = selIds
        .Select(id => doc.GetElement(id))
        .All(e => e.ViewSpecific);

      // Determine if any selected element is hosted.
      bool hasHosted = selIds
        .Select(id => doc.GetElement(id))
        .Any(e => e is FamilyInstance fi && fi.Host != null);

      using (var form = new TransformForm(onlyViewDependent, hasHosted))
      {
        if (form.ShowDialog() != WinForms.DialogResult.OK)
          return Result.Cancelled;

        TransformationSettings settings = form.Settings;

        using (Transaction trans = new Transaction(doc, "Transform Selected Elements"))
        {
          trans.Start();
          foreach (ElementId id in selIds.ToList()) // use ToList() to avoid iteration issues if elements are deleted
          {
            Element elem = doc.GetElement(id);

            // --- Move Operation ---
            if (settings.DoMove)
            {
              BoundingBoxXYZ bbox = elem.get_BoundingBox(doc.ActiveView);
              if (bbox != null)
              {
                if (onlyViewDependent)
                {
                  // In view–dependent (detail) mode, only X and Y moves are applied in the active view’s plane.
                  View activeView = doc.ActiveView;
                  XYZ up = activeView.UpDirection;
                  XYZ viewDir = activeView.ViewDirection;  // e.g., (0,0,-1) in a plan view
                  XYZ right = viewDir.CrossProduct(up);
                  XYZ displacement = (settings.MoveX / 304.8) * right +
                                     (settings.MoveY / 304.8) * up;
                  ElementTransformUtils.MoveElement(doc, id, displacement);
                }
                else
                {
                  // For 3D elements, apply the provided move in all three axes.
                  XYZ displacement = new XYZ(
                    settings.MoveX / 304.8,
                    settings.MoveY / 304.8,
                    settings.MoveZ / 304.8);
                  ElementTransformUtils.MoveElement(doc, id, displacement);
                }
              }
            }

            // --- Mirror Operation ---
            if (settings.DoMirror)
            {
              BoundingBoxXYZ bbox = elem.get_BoundingBox(doc.ActiveView);
              if (bbox != null)
              {
                XYZ center = (bbox.Min + bbox.Max) * 0.5;
                if (onlyViewDependent)
                {
                  // Only X and Y mirror options are available.
                  int mirrorCount = (settings.MirrorX ? 1 : 0) + (settings.MirrorY ? 1 : 0);
                  if (mirrorCount == 1)
                  {
                    View activeView = doc.ActiveView;
                    if (settings.MirrorX)
                    {
                      // Mirror about X: use the view’s right vector.
                      XYZ up = activeView.UpDirection;
                      XYZ viewDir = activeView.ViewDirection;
                      XYZ right = viewDir.CrossProduct(up);
                      Plane plane = Plane.CreateByNormalAndOrigin(right, center);
                      ElementTransformUtils.MirrorElement(doc, id, plane);
                      doc.Delete(id);
                    }
                    else if (settings.MirrorY)
                    {
                      // Mirror about Y: use the view’s up vector.
                      XYZ up = activeView.UpDirection;
                      Plane plane = Plane.CreateByNormalAndOrigin(up, center);
                      ElementTransformUtils.MirrorElement(doc, id, plane);
                      doc.Delete(id);
                    }
                  }
                  else if (mirrorCount == 2)
                  {
                    // Mirroring about both X and Y in the view is equivalent to a 180° rotation about the view direction.
                    View activeView = doc.ActiveView;
                    Line rotationAxis = Line.CreateBound(center, center + activeView.ViewDirection);
                    ElementTransformUtils.RotateElement(doc, id, rotationAxis, Math.PI);
                  }
                }
                else
                {
                  // 3D elements: support mirror about X, Y, and Z.
                  int mirrorCount = (settings.MirrorX ? 1 : 0) +
                                    (settings.MirrorY ? 1 : 0) +
                                    (settings.MirrorZ ? 1 : 0);
                  if (mirrorCount == 1)
                  {
                    if (settings.MirrorX)
                    {
                      Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisX, center);
                      ElementTransformUtils.MirrorElement(doc, id, plane);
                      doc.Delete(id);
                    }
                    else if (settings.MirrorY)
                    {
                      Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisY, center);
                      ElementTransformUtils.MirrorElement(doc, id, plane);
                      doc.Delete(id);
                    }
                    else if (settings.MirrorZ)
                    {
                      Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, center);
                      ElementTransformUtils.MirrorElement(doc, id, plane);
                      doc.Delete(id);
                    }
                  }
                  else if (mirrorCount == 2)
                  {
                    // Two mirror axes: use a 180° rotation about the axis perpendicular to the plane defined by them.
                    if (settings.MirrorX && settings.MirrorY)
                    {
                      Line rotationAxis = Line.CreateBound(center, center + XYZ.BasisZ);
                      ElementTransformUtils.RotateElement(doc, id, rotationAxis, Math.PI);
                    }
                    else if (settings.MirrorX && settings.MirrorZ)
                    {
                      Line rotationAxis = Line.CreateBound(center, center + XYZ.BasisY);
                      ElementTransformUtils.RotateElement(doc, id, rotationAxis, Math.PI);
                    }
                    else if (settings.MirrorY && settings.MirrorZ)
                    {
                      Line rotationAxis = Line.CreateBound(center, center + XYZ.BasisX);
                      ElementTransformUtils.RotateElement(doc, id, rotationAxis, Math.PI);
                    }
                  }
                  else if (mirrorCount == 3)
                  {
                    // All three mirror options: perform sequentially.
                    // First mirror about X and Y (equivalent to 180° rotation about Z)
                    Line rotationAxis = Line.CreateBound(center, center + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(doc, id, rotationAxis, Math.PI);
                    // Then mirror about Z.
                    Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, center);
                    ElementTransformUtils.MirrorElement(doc, id, plane);
                    doc.Delete(id);
                  }
                }
              }
            }

            // --- Rotate Operation ---
            if (settings.DoRotate)
            {
              BoundingBoxXYZ bbox = elem.get_BoundingBox(doc.ActiveView);
              if (bbox != null)
              {
                XYZ center = (bbox.Min + bbox.Max) * 0.5;
                double angleRad = settings.RotationAngle * Math.PI / 180.0;
                if (onlyViewDependent)
                {
                  View activeView = doc.ActiveView;
                  XYZ viewDir = activeView.ViewDirection;
                  Line rotationAxis = Line.CreateBound(center, center + viewDir);
                  ElementTransformUtils.RotateElement(doc, id, rotationAxis, angleRad);
                }
                else
                {
                  XYZ axis = GetAxis(settings.RotationAxis);
                  Line rotationAxis = Line.CreateBound(center, center + axis);
                  ElementTransformUtils.RotateElement(doc, id, rotationAxis, angleRad);
                }
              }
            }
          }
          trans.Commit();
        }
      }
      return Result.Succeeded;
    }

    private XYZ GetAxis(TransformationAxis axis)
    {
      switch (axis)
      {
        case TransformationAxis.X: return XYZ.BasisX;
        case TransformationAxis.Y: return XYZ.BasisY;
        case TransformationAxis.Z:
        default: return XYZ.BasisZ;
      }
    }
  }

  // The WinForm uses a tabbed layout with separate tabs for Move, Mirror, and Rotate.
  public class TransformForm : WinForms.Form
  {
    private bool _onlyViewDependent;
    private bool _disableY;
    private WinForms.TabControl tabControl;
    private WinForms.TabPage tabMove, tabMirror, tabRotate;

    // Move tab controls.
    private WinForms.Label lblMoveX, lblMoveY, lblMoveZ;
    private WinForms.TextBox txtMoveX, txtMoveY, txtMoveZ;

    // Mirror tab controls (using checkboxes).
    private WinForms.CheckBox cbMirrorX, cbMirrorY, cbMirrorZ;

    // Rotate tab controls.
    private WinForms.Label lblRotateAngle;
    private WinForms.TextBox txtRotateAngle;
    private WinForms.RadioButton rbRotateX, rbRotateY, rbRotateZ;

    private WinForms.Button btnOK, btnCancel;

    public TransformationSettings Settings { get; private set; }

    /// <summary>
    /// onlyViewDependent: if true, only in‑plane (X/Y) operations are allowed.
    /// disableY: if true, the Y axis field in the Move tab is disabled because the selected element(s) are hosted.
    /// </summary>
    public TransformForm(bool onlyViewDependent, bool disableY)
    {
      _onlyViewDependent = onlyViewDependent;
      _disableY = disableY;

      this.Text = "Transform Selected Elements";
      this.Width = 300;
      this.Height = 300;
      this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
      this.StartPosition = WinForms.FormStartPosition.CenterScreen;
      this.MaximizeBox = false;
      this.MinimizeBox = false;

      tabControl = new WinForms.TabControl();
      tabControl.Dock = WinForms.DockStyle.Top;
      tabControl.Height = 200;

      // ---- Move Tab ----
      tabMove = new WinForms.TabPage("Move");
      lblMoveX = new WinForms.Label() { Text = "Move X (mm):", Left = 10, Top = 20, Width = 90 };
      txtMoveX = new WinForms.TextBox() { Left = 110, Top = 20, Width = 100, Text = "0" };
      lblMoveY = new WinForms.Label() { Text = "Move Y (mm):", Left = 10, Top = 50, Width = 90 };
      txtMoveY = new WinForms.TextBox() { Left = 110, Top = 50, Width = 100, Text = "0" };
      lblMoveZ = new WinForms.Label() { Text = "Move Z (mm):", Left = 10, Top = 80, Width = 90 };
      txtMoveZ = new WinForms.TextBox() { Left = 110, Top = 80, Width = 100, Text = "0" };
      tabMove.Controls.AddRange(new WinForms.Control[] { lblMoveX, txtMoveX, lblMoveY, txtMoveY, lblMoveZ, txtMoveZ });

      // ---- Mirror Tab ----
      tabMirror = new WinForms.TabPage("Mirror");
      cbMirrorX = new WinForms.CheckBox() { Text = "Mirror about X", Left = 10, Top = 20, Width = 150 };
      cbMirrorY = new WinForms.CheckBox() { Text = "Mirror about Y", Left = 10, Top = 50, Width = 150 };
      cbMirrorZ = new WinForms.CheckBox() { Text = "Mirror about Z", Left = 10, Top = 80, Width = 150 };
      tabMirror.Controls.AddRange(new WinForms.Control[] { cbMirrorX, cbMirrorY, cbMirrorZ });

      // ---- Rotate Tab ----
      tabRotate = new WinForms.TabPage("Rotate");
      lblRotateAngle = new WinForms.Label() { Text = "Angle (deg):", Left = 10, Top = 20, Width = 90 };
      txtRotateAngle = new WinForms.TextBox() { Left = 110, Top = 20, Width = 100, Text = "0" };
      rbRotateX = new WinForms.RadioButton() { Text = "Rotate about X", Left = 10, Top = 50, Width = 150 };
      rbRotateY = new WinForms.RadioButton() { Text = "Rotate about Y", Left = 10, Top = 80, Width = 150 };
      rbRotateZ = new WinForms.RadioButton() { Text = "Rotate about Z", Left = 10, Top = 110, Width = 150 };
      rbRotateX.Checked = true;
      tabRotate.Controls.AddRange(new WinForms.Control[] { lblRotateAngle, txtRotateAngle, rbRotateX, rbRotateY, rbRotateZ });

      tabControl.TabPages.AddRange(new WinForms.TabPage[] { tabMove, tabMirror, tabRotate });
      this.Controls.Add(tabControl);

      btnOK = new WinForms.Button() { Text = "OK", Left = 50, Width = 80, Top = 220, DialogResult = WinForms.DialogResult.OK };
      btnCancel = new WinForms.Button() { Text = "Cancel", Left = 150, Width = 80, Top = 220, DialogResult = WinForms.DialogResult.Cancel };
      this.Controls.Add(btnOK);
      this.Controls.Add(btnCancel);
      this.AcceptButton = btnOK;
      this.CancelButton = btnCancel;

      if (_onlyViewDependent)
      {
        // In view–dependent mode, hide Z controls.
        lblMoveZ.Visible = false;
        txtMoveZ.Visible = false;
        cbMirrorZ.Visible = false;
        rbRotateX.Visible = false;
        rbRotateY.Visible = false;
        rbRotateZ.Text = "Rotate about view normal";
        rbRotateZ.Checked = true;
      }
      if (_disableY)
      {
        txtMoveY.Enabled = false;
        txtMoveY.Text = "0";
      }
    }

    protected override void OnFormClosing(WinForms.FormClosingEventArgs e)
    {
      base.OnFormClosing(e);
      if (this.DialogResult == WinForms.DialogResult.OK)
      {
        TransformationSettings settings = new TransformationSettings();

        // --- Move Tab ---
        double.TryParse(txtMoveX.Text, out double mvX);
        double.TryParse(txtMoveY.Text, out double mvY);
        double.TryParse(txtMoveZ.Text, out double mvZ);
        if (Math.Abs(mvX) > 0 || Math.Abs(mvY) > 0 || Math.Abs(mvZ) > 0)
        {
          settings.DoMove = true;
          settings.MoveX = mvX;
          settings.MoveY = mvY;
          settings.MoveZ = mvZ;
        }
        else
          settings.DoMove = false;

        // --- Mirror Tab ---
        settings.MirrorX = cbMirrorX.Checked;
        settings.MirrorY = cbMirrorY.Checked;
        settings.MirrorZ = (!_onlyViewDependent && cbMirrorZ.Checked);
        settings.DoMirror = settings.MirrorX || settings.MirrorY || settings.MirrorZ;

        // --- Rotate Tab ---
        double.TryParse(txtRotateAngle.Text, out double angle);
        if (Math.Abs(angle) > 0)
        {
          settings.DoRotate = true;
          settings.RotationAngle = angle;
          if (_onlyViewDependent)
            settings.RotationAxis = TransformationAxis.Z;
          else
          {
            if (rbRotateX.Checked)
              settings.RotationAxis = TransformationAxis.X;
            else if (rbRotateY.Checked)
              settings.RotationAxis = TransformationAxis.Y;
            else if (rbRotateZ.Checked)
              settings.RotationAxis = TransformationAxis.Z;
            else
              settings.RotationAxis = TransformationAxis.X;
          }
        }
        else
          settings.DoRotate = false;

        this.Settings = settings;
      }
    }
  }
}
