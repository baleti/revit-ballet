#region Namespaces
using System;
using System.Linq;
using System.Collections.Generic;
using System.Drawing;  // For System.Drawing.Color and Font
using System.Globalization;
using System.IO;
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

    // Mirror parameters:
    public bool DoMirror { get; set; }
    public bool MirrorX { get; set; }
    public bool MirrorY { get; set; }
    public bool MirrorZ { get; set; }  // Only available for non–view–dependent elements.

    // Rotate parameters:
    public bool DoRotate { get; set; }
    public double RotationAngle { get; set; }
    public TransformationAxis RotationAxis { get; set; }
    
    // Set Origin parameters (applied only if enabled)
    public bool DoSetOrigin { get; set; }
    public double OriginX { get; set; }
    public double OriginY { get; set; }
    public double OriginZ { get; set; }
  }

  [Transaction(TransactionMode.Manual)]
  public class TransformSelectedElements : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Document doc = uidoc.Document;

      var selIds = uidoc.GetSelectionIds();
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

      // Determine if any hosted element is hosted on a Wall.
      bool hostedOnWall = selIds
        .Select(id => doc.GetElement(id))
        .OfType<FamilyInstance>()
        .Any(fi => fi.Host is Wall);

      // Determine if the project is metric.
      Units projectUnits = doc.GetUnits();
      FormatOptions lengthOptions = projectUnits.GetFormatOptions(SpecTypeId.Length);
      bool isMetric = (lengthOptions.GetUnitTypeId() == UnitTypeId.Millimeters ||
                       lengthOptions.GetUnitTypeId() == UnitTypeId.Meters);
      // When metric, assume user input is in millimeters and convert to feet.
      double moveConversionFactor = isMetric ? (1.0 / 304.8) : 1.0;

      using (var form = new TransformForm(onlyViewDependent, hasHosted, hostedOnWall))
      {
        if (form.ShowDialog() != WinForms.DialogResult.OK)
          return Result.Cancelled;

        TransformationSettings settings = form.Settings;

        using (Transaction trans = new Transaction(doc, "Transform Selected Elements"))
        {
          trans.Start();
          foreach (ElementId id in selIds.ToList()) // Use ToList() to avoid iteration issues if elements are deleted.
          {
            Element elem = doc.GetElement(id);

            // --- Move Operation ---
            if (settings.DoMove)
            {
              BoundingBoxXYZ bbox = elem.get_BoundingBox(doc.ActiveView);
              if (bbox != null)
              {
                if (elem is FamilyInstance fi && fi.Host != null && fi.Host is Wall wall)
                {
                  LocationCurve hostLoc = wall.Location as LocationCurve;
                  if (hostLoc != null)
                  {
                    XYZ start = hostLoc.Curve.GetEndPoint(0);
                    XYZ end = hostLoc.Curve.GetEndPoint(1);
                    XYZ wallTangent = (end - start).Normalize();
                    XYZ wallPerp = new XYZ(-wallTangent.Y, wallTangent.X, 0);
                    XYZ displacement = (settings.MoveX * moveConversionFactor) * wallTangent +
                                       (settings.MoveY * moveConversionFactor) * wallPerp;
                    ElementTransformUtils.MoveElement(doc, id, displacement);
                  }
                }
                else if (onlyViewDependent)
                {
                  View activeView = doc.ActiveView;
                  XYZ up = activeView.UpDirection;
                  XYZ viewDir = activeView.ViewDirection;
                  XYZ right = viewDir.CrossProduct(up);
                  XYZ displacement = (settings.MoveX * moveConversionFactor) * right +
                                     (settings.MoveY * moveConversionFactor) * up;
                  ElementTransformUtils.MoveElement(doc, id, displacement);
                }
                else
                {
                  XYZ displacement = new XYZ(
                    settings.MoveX * moveConversionFactor,
                    settings.MoveY * moveConversionFactor,
                    settings.MoveZ * moveConversionFactor);
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
                  int mirrorCount = (settings.MirrorX ? 1 : 0) + (settings.MirrorY ? 1 : 0);
                  if (mirrorCount == 1)
                  {
                    View activeView = doc.ActiveView;
                    if (settings.MirrorX)
                    {
                      XYZ up = activeView.UpDirection;
                      XYZ viewDir = activeView.ViewDirection;
                      XYZ right = viewDir.CrossProduct(up);
                      Plane plane = Plane.CreateByNormalAndOrigin(right, center);
                      ElementTransformUtils.MirrorElement(doc, id, plane);
                      doc.Delete(id);
                    }
                    else if (settings.MirrorY)
                    {
                      XYZ up = activeView.UpDirection;
                      Plane plane = Plane.CreateByNormalAndOrigin(up, center);
                      ElementTransformUtils.MirrorElement(doc, id, plane);
                      doc.Delete(id);
                    }
                  }
                  else if (mirrorCount == 2)
                  {
                    View activeView = doc.ActiveView;
                    Line rotationAxis = Line.CreateBound(center, center + activeView.ViewDirection);
                    ElementTransformUtils.RotateElement(doc, id, rotationAxis, Math.PI);
                  }
                }
                else
                {
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
                    Line rotationAxis = Line.CreateBound(center, center + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(doc, id, rotationAxis, Math.PI);
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
            
            // --- Set Origin Operation ---
            if (settings.DoSetOrigin)
            {
              double originX = settings.OriginX;
              double originY = settings.OriginY;
              double originZ = settings.OriginZ;

              XYZ newOrigin = new XYZ(originX, originY, originZ);

              if (elem is RevitLinkInstance rli)
              {
                Transform linkTransform = rli.GetTransform();
                XYZ currentOrigin = linkTransform.Origin;
                XYZ displacement = newOrigin - currentOrigin;
                ElementTransformUtils.MoveElement(doc, id, displacement);
              }
              else if (elem.Location != null)
              {
                if (elem.Location is LocationPoint locPoint)
                {
                  locPoint.Point = newOrigin;
                }
                else if (elem.Location is LocationCurve locCurve)
                {
                  XYZ currentOrigin = locCurve.Curve.GetEndPoint(0);
                  XYZ displacement = newOrigin - currentOrigin;
                  ElementTransformUtils.MoveElement(doc, id, displacement);
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

  // The TransformForm now displays all options in one view.
  // It loads and saves the last values to %APPDATA%\revit-scripts\TransformSelectedElements.
  // A "Reset" button resets all fields, and a "Set new origin" checkbox controls the Set Origin options.
  public class TransformForm : WinForms.Form
  {
    private bool _onlyViewDependent;
    private bool _disableY;
    private bool _hostedOnWall;

    // GroupBoxes for organization.
    private WinForms.GroupBox groupMove;
    private WinForms.GroupBox groupMirror;
    private WinForms.GroupBox groupRotate;
    private WinForms.GroupBox groupSetOrigin;

    // Move controls.
    private WinForms.Label lblMoveX, lblMoveY, lblMoveZ;
    private WinForms.TextBox txtMoveX, txtMoveY, txtMoveZ;
    private WinForms.Label lblLocalX, lblLocalY, lblLocalZ;
    private WinForms.TextBox txtLocalX, txtLocalY, txtLocalZ;

    // Mirror controls.
    private WinForms.CheckBox cbMirrorX, cbMirrorY, cbMirrorZ;

    // Rotate controls.
    private WinForms.Label lblRotateAngle;
    private WinForms.TextBox txtRotateAngle;
    private WinForms.RadioButton rbRotateX, rbRotateY, rbRotateZ;

    // Set Origin controls.
    private WinForms.CheckBox cbSetOrigin;
    private WinForms.Label lblOriginX, lblOriginY, lblOriginZ;
    private WinForms.TextBox txtOriginX, txtOriginY, txtOriginZ;

    // Buttons.
    private WinForms.Button btnOK, btnCancel, btnReset;

    public TransformationSettings Settings { get; private set; }

    // File path to persist settings.
    private string settingsFilePath;

    public TransformForm(bool onlyViewDependent, bool disableY, bool hostedOnWall)
    {
      _onlyViewDependent = onlyViewDependent;
      _disableY = disableY;
      _hostedOnWall = hostedOnWall;

      this.Text = "Transform Selected Elements";
      this.Width = 400; // Increased width.
      this.Height = 650;
      this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
      this.StartPosition = WinForms.FormStartPosition.CenterScreen;
      this.MaximizeBox = false;
      this.MinimizeBox = false;

      // Determine settings file path.
      string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
      string folder = Path.Combine(appData, "revit-scripts");
      if (!Directory.Exists(folder))
        Directory.CreateDirectory(folder);
      settingsFilePath = Path.Combine(folder, "TransformSelectedElements");

      // --- GroupBox: Move ---
      groupMove = new WinForms.GroupBox();
      groupMove.Text = "Move";
      int moveGroupHeight = _disableY ? (_onlyViewDependent ? 120 : (_hostedOnWall ? 150 : 180)) : 110;
      groupMove.SetBounds(10, 10, 370, moveGroupHeight);

      lblMoveX = new WinForms.Label() { Text = "Move X (mm):", Left = 10, Top = 20, Width = 80 };
      txtMoveX = new WinForms.TextBox() { Left = 100, Top = 20, Width = 120, Text = "0" };
      lblMoveY = new WinForms.Label() { Text = "Move Y (mm):", Left = 10, Top = 50, Width = 80 };
      txtMoveY = new WinForms.TextBox() { Left = 100, Top = 50, Width = 120, Text = "0" };
      lblMoveZ = new WinForms.Label() { Text = "Move Z (mm):", Left = 10, Top = 80, Width = 80 };
      txtMoveZ = new WinForms.TextBox() { Left = 100, Top = 80, Width = 120, Text = "0" };
      groupMove.Controls.AddRange(new WinForms.Control[] { lblMoveX, txtMoveX, lblMoveY, txtMoveY, lblMoveZ, txtMoveZ });

      if (_disableY)
      {
        txtMoveX.Enabled = false;
        txtMoveY.Enabled = false;
        txtMoveZ.Enabled = false;
        txtMoveX.BackColor = System.Drawing.Color.LightGray;
        txtMoveY.BackColor = System.Drawing.Color.LightGray;
        txtMoveZ.BackColor = System.Drawing.Color.LightGray;

        int localYStart = 110;
        lblLocalX = new WinForms.Label() { Text = "Local X (mm):", Left = 10, Top = localYStart, Width = 80 };
        txtLocalX = new WinForms.TextBox() { Left = 100, Top = localYStart, Width = 120, Text = "0" };
        groupMove.Controls.Add(lblLocalX);
        groupMove.Controls.Add(txtLocalX);
        localYStart += 30;
        lblLocalY = new WinForms.Label() { Text = "Local Y (mm):", Left = 10, Top = localYStart, Width = 80 };
        txtLocalY = new WinForms.TextBox() { Left = 100, Top = localYStart, Width = 120, Text = "0" };
        groupMove.Controls.Add(lblLocalY);
        groupMove.Controls.Add(txtLocalY);
        if (!_onlyViewDependent && !_hostedOnWall)
        {
          localYStart += 30;
          lblLocalZ = new WinForms.Label() { Text = "Local Z (mm):", Left = 10, Top = localYStart, Width = 80 };
          txtLocalZ = new WinForms.TextBox() { Left = 100, Top = localYStart, Width = 120, Text = "0" };
          groupMove.Controls.Add(lblLocalZ);
          groupMove.Controls.Add(txtLocalZ);
        }
      }

      // --- GroupBox: Mirror ---
      groupMirror = new WinForms.GroupBox();
      groupMirror.Text = "Mirror";
      groupMirror.SetBounds(10, groupMove.Bottom + 10, 370, 120);
      cbMirrorX = new WinForms.CheckBox() { Text = "Mirror about X", Left = 10, Top = 20, Width = 150 };
      cbMirrorY = new WinForms.CheckBox() { Text = "Mirror about Y", Left = 10, Top = 45, Width = 150 };
      cbMirrorZ = new WinForms.CheckBox() { Text = "Mirror about Z", Left = 10, Top = 70, Width = 150 };
      groupMirror.Controls.AddRange(new WinForms.Control[] { cbMirrorX, cbMirrorY, cbMirrorZ });
      if (_onlyViewDependent)
        cbMirrorZ.Visible = false;

      // --- GroupBox: Rotate ---
      groupRotate = new WinForms.GroupBox();
      groupRotate.Text = "Rotate";
      groupRotate.SetBounds(10, groupMirror.Bottom + 10, 370, _onlyViewDependent ? 100 : 130);
      lblRotateAngle = new WinForms.Label() { Text = "Angle (deg):", Left = 10, Top = 20, Width = 80 };
      txtRotateAngle = new WinForms.TextBox() { Left = 100, Top = 20, Width = 120, Text = "0" };
      groupRotate.Controls.Add(lblRotateAngle);
      groupRotate.Controls.Add(txtRotateAngle);
      if (_onlyViewDependent)
      {
        rbRotateZ = new WinForms.RadioButton() { Text = "Rotate about view normal", Left = 10, Top = 50, Width = 200 };
        rbRotateZ.Checked = true;
        groupRotate.Controls.Add(rbRotateZ);
      }
      else
      {
        rbRotateX = new WinForms.RadioButton() { Text = "Rotate about X", Left = 10, Top = 50, Width = 150 };
        rbRotateY = new WinForms.RadioButton() { Text = "Rotate about Y", Left = 10, Top = 75, Width = 150 };
        rbRotateZ = new WinForms.RadioButton() { Text = "Rotate about Z", Left = 10, Top = 100, Width = 150 };
        rbRotateX.Checked = true;
        groupRotate.Controls.AddRange(new WinForms.Control[] { rbRotateX, rbRotateY, rbRotateZ });
      }

      // --- GroupBox: Set Origin ---
      groupSetOrigin = new WinForms.GroupBox();
      groupSetOrigin.Text = "Set Origin";
      groupSetOrigin.SetBounds(10, groupRotate.Bottom + 10, 370, 150);

      // Checkbox to enable setting origin.
      cbSetOrigin = new WinForms.CheckBox() { Text = "Set new origin", Left = 10, Top = 20, Width = 150 };
      cbSetOrigin.CheckedChanged += new EventHandler(cbSetOrigin_CheckedChanged);
      groupSetOrigin.Controls.Add(cbSetOrigin);

      lblOriginX = new WinForms.Label() { Text = "Origin X (mm):", Left = 10, Top = 50, Width = 80 };
      txtOriginX = new WinForms.TextBox() { Left = 100, Top = 50, Width = 120, Text = "0" };
      lblOriginY = new WinForms.Label() { Text = "Origin Y (mm):", Left = 10, Top = 80, Width = 80 };
      txtOriginY = new WinForms.TextBox() { Left = 100, Top = 80, Width = 120, Text = "0" };
      lblOriginZ = new WinForms.Label() { Text = "Origin Z (mm):", Left = 10, Top = 110, Width = 80 };
      txtOriginZ = new WinForms.TextBox() { Left = 100, Top = 110, Width = 120, Text = "0" };
      groupSetOrigin.Controls.AddRange(new WinForms.Control[] { lblOriginX, txtOriginX, lblOriginY, txtOriginY, lblOriginZ, txtOriginZ });

      this.Controls.AddRange(new WinForms.Control[] { groupMove, groupMirror, groupRotate, groupSetOrigin });

      // OK, Cancel, and Reset buttons.
      btnOK = new WinForms.Button() { Text = "OK", Left = 40, Width = 80, Top = groupSetOrigin.Bottom + 20, DialogResult = WinForms.DialogResult.OK };
      btnCancel = new WinForms.Button() { Text = "Cancel", Left = 140, Width = 80, Top = groupSetOrigin.Bottom + 20, DialogResult = WinForms.DialogResult.Cancel };
      btnReset = new WinForms.Button() { Text = "Reset", Left = 240, Width = 80, Top = groupSetOrigin.Bottom + 20 };
      btnReset.Click += new EventHandler(btnReset_Click);
      this.Controls.AddRange(new WinForms.Control[] { btnOK, btnCancel, btnReset });
      this.AcceptButton = btnOK;
      this.CancelButton = btnCancel;

      UpdateOriginControls();
      LoadSettings();
    }

    private void cbSetOrigin_CheckedChanged(object sender, EventArgs e)
    {
      UpdateOriginControls();
    }

    private void UpdateOriginControls()
    {
      bool enabled = cbSetOrigin.Checked;
      txtOriginX.Enabled = enabled;
      txtOriginY.Enabled = enabled;
      txtOriginZ.Enabled = enabled;
      var bg = enabled ? SystemColors.Window : System.Drawing.Color.LightGray;
      txtOriginX.BackColor = bg;
      txtOriginY.BackColor = bg;
      txtOriginZ.BackColor = bg;
    }

    protected override void OnFormClosing(WinForms.FormClosingEventArgs e)
    {
      base.OnFormClosing(e);
      if (this.DialogResult == WinForms.DialogResult.OK)
      {
        TransformationSettings settings = new TransformationSettings();

        // --- Move ---
        if (_disableY)
        {
          double.TryParse(txtLocalX.Text, out double mvX);
          double.TryParse(txtLocalY.Text, out double mvY);
          double mvZ = 0;
          bool showLocalZ = (!_onlyViewDependent && !_hostedOnWall);
          if (showLocalZ && txtLocalZ != null)
            double.TryParse(txtLocalZ.Text, out mvZ);
          if (Math.Abs(mvX) > 0 || Math.Abs(mvY) > 0 || Math.Abs(mvZ) > 0)
          {
            settings.DoMove = true;
            settings.MoveX = mvX;
            settings.MoveY = mvY;
            settings.MoveZ = mvZ;
          }
          else
            settings.DoMove = false;
        }
        else
        {
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
        }

        // --- Mirror ---
        settings.MirrorX = cbMirrorX.Checked;
        settings.MirrorY = cbMirrorY.Checked;
        settings.MirrorZ = (!_onlyViewDependent && cbMirrorZ.Checked);
        settings.DoMirror = settings.MirrorX || settings.MirrorY || settings.MirrorZ;

        // --- Rotate ---
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

        // --- Set Origin ---
        settings.DoSetOrigin = cbSetOrigin.Checked;
        if (settings.DoSetOrigin)
        {
          double.TryParse(txtOriginX.Text, out double originX);
          double.TryParse(txtOriginY.Text, out double originY);
          double.TryParse(txtOriginZ.Text, out double originZ);
          settings.OriginX = originX;
          settings.OriginY = originY;
          settings.OriginZ = originZ;
        }

        this.Settings = settings;
        SaveSettings();
      }
    }

    private void btnReset_Click(object sender, EventArgs e)
    {
      txtMoveX.Text = "0"; txtMoveY.Text = "0"; txtMoveZ.Text = "0";
      if (txtLocalX != null) txtLocalX.Text = "0";
      if (txtLocalY != null) txtLocalY.Text = "0";
      if (txtLocalZ != null) txtLocalZ.Text = "0";
      cbMirrorX.Checked = false; cbMirrorY.Checked = false; cbMirrorZ.Checked = false;
      txtRotateAngle.Text = "0";
      if (_onlyViewDependent)
      {
        if (rbRotateZ != null) rbRotateZ.Checked = true;
      }
      else
      {
        if (rbRotateX != null) rbRotateX.Checked = true;
        if (rbRotateY != null) rbRotateY.Checked = false;
        if (rbRotateZ != null) rbRotateZ.Checked = false;
      }
      cbSetOrigin.Checked = false;
      txtOriginX.Text = "0"; txtOriginY.Text = "0"; txtOriginZ.Text = "0";
    }

    private void LoadSettings()
    {
      if (File.Exists(settingsFilePath))
      {
        try
        {
          var lines = File.ReadAllLines(settingsFilePath);
          foreach (var line in lines)
          {
            if (string.IsNullOrWhiteSpace(line)) continue;
            var parts = line.Split(new char[] { '=' }, 2);
            if (parts.Length < 2) continue;
            string key = parts[0].Trim();
            string value = parts[1].Trim();
            switch (key)
            {
              case "MoveX": txtMoveX.Text = value; break;
              case "MoveY": txtMoveY.Text = value; break;
              case "MoveZ": txtMoveZ.Text = value; break;
              case "LocalX": if (txtLocalX != null) txtLocalX.Text = value; break;
              case "LocalY": if (txtLocalY != null) txtLocalY.Text = value; break;
              case "LocalZ": if (txtLocalZ != null) txtLocalZ.Text = value; break;
              case "MirrorX": cbMirrorX.Checked = bool.TryParse(value, out bool bx) && bx; break;
              case "MirrorY": cbMirrorY.Checked = bool.TryParse(value, out bool by) && by; break;
              case "MirrorZ": cbMirrorZ.Checked = bool.TryParse(value, out bool bz) && bz; break;
              case "RotateAngle": txtRotateAngle.Text = value; break;
              case "RotateAxis":
                if (!_onlyViewDependent)
                {
                  if (value == "X")
                    rbRotateX.Checked = true;
                  else if (value == "Y")
                    rbRotateY.Checked = true;
                  else if (value == "Z")
                    rbRotateZ.Checked = true;
                }
                break;
              case "SetOrigin": cbSetOrigin.Checked = bool.TryParse(value, out bool so) && so; break;
              case "OriginX": txtOriginX.Text = value; break;
              case "OriginY": txtOriginY.Text = value; break;
              case "OriginZ": txtOriginZ.Text = value; break;
            }
          }
          UpdateOriginControls();
        }
        catch { /* Ignore errors */ }
      }
    }

    private void SaveSettings()
    {
      var lines = new List<string>();
      lines.Add("MoveX=" + txtMoveX.Text);
      lines.Add("MoveY=" + txtMoveY.Text);
      lines.Add("MoveZ=" + txtMoveZ.Text);
      if (txtLocalX != null) lines.Add("LocalX=" + txtLocalX.Text);
      if (txtLocalY != null) lines.Add("LocalY=" + txtLocalY.Text);
      if (txtLocalZ != null) lines.Add("LocalZ=" + txtLocalZ.Text);
      lines.Add("MirrorX=" + cbMirrorX.Checked.ToString());
      lines.Add("MirrorY=" + cbMirrorY.Checked.ToString());
      lines.Add("MirrorZ=" + cbMirrorZ.Checked.ToString());
      lines.Add("RotateAngle=" + txtRotateAngle.Text);
      if (!_onlyViewDependent)
      {
        if (rbRotateX != null && rbRotateX.Checked)
          lines.Add("RotateAxis=X");
        else if (rbRotateY != null && rbRotateY.Checked)
          lines.Add("RotateAxis=Y");
        else if (rbRotateZ != null && rbRotateZ.Checked)
          lines.Add("RotateAxis=Z");
      }
      lines.Add("SetOrigin=" + cbSetOrigin.Checked.ToString());
      lines.Add("OriginX=" + txtOriginX.Text);
      lines.Add("OriginY=" + txtOriginY.Text);
      lines.Add("OriginZ=" + txtOriginZ.Text);
      try
      {
        File.WriteAllLines(settingsFilePath, lines);
      }
      catch { /* Ignore errors */ }
    }
  }
}
