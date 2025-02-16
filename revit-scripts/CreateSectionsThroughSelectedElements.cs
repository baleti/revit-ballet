using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;
using System.ComponentModel;
using System.Text.RegularExpressions;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

namespace MyRevitCommands
{
  // Holds settings plus project name.
  public class SectionSettings
  {
      public ElementId SelectedSectionTypeId { get; set; }
      // Global orientation mode: "Element", "Host", or "CurrentView"
      public string OrientationMode { get; set; }
      // Local orientation: one of "0°", "90°", "180°", "270°"
      public string LocalOrientation { get; set; }
      public string ProjectName { get; set; }
  }

  // A panel that draws a subtle light–blue border when any child control is focused.
  public class FocusPanel : WinForms.Panel
  {
      public bool IsFocused { get; set; }
      protected override void OnPaint(WinForms.PaintEventArgs e)
      {
          base.OnPaint(e);
          if (IsFocused)
          {
              using (var pen = new Drawing.Pen(Drawing.Color.LightBlue, 1))
              {
                  e.Graphics.DrawRectangle(pen, 0, 0, this.Width - 1, this.Height - 1);
              }
          }
      }
  }

  // Custom settings form.
  public class SectionSettingsForm : WinForms.Form
  {
      private WinForms.Label lblSectionType;
      private FocusPanel fpSectionType;
      private WinForms.TextBox txtSearch;
      private WinForms.ListBox lstSectionTypes;

      private FocusPanel fpOrientation;
      private WinForms.GroupBox gbGlobalOrientation;
      private WinForms.RadioButton rbElements;
      private WinForms.RadioButton rbHosts;
      private WinForms.RadioButton rbCurrentView;
      private WinForms.Label lblLocalOrientation;
      private WinForms.ComboBox cmbLocalOrientation;

      private WinForms.Button btnOK;
      private WinForms.Button btnCancel;

      private List<ViewFamilyType> sectionTypes;
      private List<ViewFamilyType> filteredSectionTypes;
      public SectionSettings ResultSettings { get; private set; }
      private bool hasHostedElements;
      private string currentProjectName;
      private string settingsFilePath;

      public SectionSettingsForm(List<ViewFamilyType> availableSectionTypes, bool hasHostedElements, string projectName)
      {
          sectionTypes = availableSectionTypes;
          filteredSectionTypes = new List<ViewFamilyType>(availableSectionTypes);
          this.hasHostedElements = hasHostedElements;
          this.currentProjectName = projectName;
          InitializeComponents();
          LoadSettings();
      }

      private void InitializeComponents()
      {
          this.Text = "Section Creation Settings";
          this.Width = 400;
          this.Height = 500;
          this.StartPosition = WinForms.FormStartPosition.CenterScreen;
          this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
          this.MaximizeBox = false;
          this.MinimizeBox = false;
          this.KeyPreview = true;

          // Create Section Type label.
          lblSectionType = new WinForms.Label() { Text = "Section Type", Dock = WinForms.DockStyle.Top, Height = 20, TabIndex = 0 };

          // Create search box.
          txtSearch = new WinForms.TextBox() { Dock = WinForms.DockStyle.Fill, TabIndex = 1 };
          // Wrap in FocusPanel.
          fpSectionType = new FocusPanel() { Dock = WinForms.DockStyle.Top, Height = txtSearch.Height + 4 };
          fpSectionType.Controls.Add(txtSearch);
          txtSearch.Enter += (s, e) => { fpSectionType.IsFocused = true; fpSectionType.Invalidate(); };
          txtSearch.Leave += (s, e) => { fpSectionType.IsFocused = false; fpSectionType.Invalidate(); };

          // Create list box.
          lstSectionTypes = new WinForms.ListBox() { Dock = WinForms.DockStyle.Fill, TabIndex = 2 };
          // Wrap in FocusPanel.
          FocusPanel fpList = new FocusPanel() { Dock = WinForms.DockStyle.Top, Height = 150 + 4 };
          fpList.Controls.Add(lstSectionTypes);
          lstSectionTypes.Enter += (s, e) => { fpList.IsFocused = true; fpList.Invalidate(); };
          lstSectionTypes.Leave += (s, e) => { fpList.IsFocused = false; fpList.Invalidate(); };

          // Global Orientation group.
          gbGlobalOrientation = new WinForms.GroupBox() { Text = "Global Orientation", Dock = WinForms.DockStyle.Top, Height = 100, TabIndex = 3 };
          rbElements = new WinForms.RadioButton() { Text = "Elements", Dock = WinForms.DockStyle.Top, Checked = true, TabIndex = 0 };
          rbHosts = new WinForms.RadioButton() { Text = "Hosts", Dock = WinForms.DockStyle.Top, TabIndex = 1 };
          rbCurrentView = new WinForms.RadioButton() { Text = "Current View", Dock = WinForms.DockStyle.Top, TabIndex = 2 };
          if (!hasHostedElements)
              rbHosts.Enabled = false;
          // Use a nested TableLayoutPanel inside the group.
          WinForms.TableLayoutPanel tlpGlobal = new WinForms.TableLayoutPanel();
          tlpGlobal.Dock = WinForms.DockStyle.Fill;
          tlpGlobal.RowCount = 3;
          tlpGlobal.ColumnCount = 1;
          tlpGlobal.Controls.Add(rbElements, 0, 0);
          tlpGlobal.Controls.Add(rbHosts, 0, 1);
          tlpGlobal.Controls.Add(rbCurrentView, 0, 2);
          gbGlobalOrientation.Controls.Add(tlpGlobal);

          // Local Orientation label and dropdown.
          lblLocalOrientation = new WinForms.Label() { Text = "Local Orientation", Dock = WinForms.DockStyle.Top, Height = 20, TabIndex = 4 };
          cmbLocalOrientation = new WinForms.ComboBox() { Dock = WinForms.DockStyle.Top, DropDownStyle = WinForms.ComboBoxStyle.DropDownList, TabIndex = 5 };
          cmbLocalOrientation.Items.AddRange(new string[] { "0°", "90°", "180°", "270°" });
          cmbLocalOrientation.SelectedIndex = 0;

          // Wrap the entire orientation area in a FocusPanel.
          fpOrientation = new FocusPanel() { Dock = WinForms.DockStyle.Top, Height = 160 };
          // Use a TableLayoutPanel to organize Global and Local orientation.
          WinForms.TableLayoutPanel tlpOrientation = new WinForms.TableLayoutPanel();
          tlpOrientation.Dock = WinForms.DockStyle.Fill;
          tlpOrientation.RowCount = 3;
          tlpOrientation.ColumnCount = 1;
          tlpOrientation.Controls.Add(gbGlobalOrientation, 0, 0);
          tlpOrientation.Controls.Add(lblLocalOrientation, 0, 1);
          tlpOrientation.Controls.Add(cmbLocalOrientation, 0, 2);
          fpOrientation.Controls.Add(tlpOrientation);

          // Buttons.
          btnOK = new WinForms.Button() { Text = "OK", DialogResult = WinForms.DialogResult.OK, Dock = WinForms.DockStyle.Bottom, TabIndex = 6 };
          btnCancel = new WinForms.Button() { Text = "Cancel", DialogResult = WinForms.DialogResult.Cancel, Dock = WinForms.DockStyle.Bottom, TabIndex = 7 };
          this.AcceptButton = btnOK;
          this.CancelButton = btnCancel;

          // Main layout.
          WinForms.TableLayoutPanel panel = new WinForms.TableLayoutPanel();
          panel.Dock = WinForms.DockStyle.Fill;
          panel.RowCount = 6;
          panel.ColumnCount = 1;
          panel.AutoSize = true;
          panel.Controls.Add(lblSectionType, 0, 0);
          panel.Controls.Add(fpSectionType, 0, 1); // instead of fpSearch
          panel.Controls.Add(fpList, 0, 2);
          panel.Controls.Add(fpOrientation, 0, 3);
          panel.Controls.Add(btnOK, 0, 4);
          panel.Controls.Add(btnCancel, 0, 5);
          this.Controls.Add(panel);

          txtSearch.TextChanged += TxtSearch_TextChanged;
          this.KeyDown += SectionSettingsForm_KeyDown;

          UpdateListBox();
      }

      private void SectionSettingsForm_KeyDown(object sender, WinForms.KeyEventArgs e)
      {
          if (e.KeyCode == WinForms.Keys.Enter)
          {
              this.DialogResult = WinForms.DialogResult.OK;
              this.Close();
          }
          else if (e.KeyCode == WinForms.Keys.Escape)
          {
              this.DialogResult = WinForms.DialogResult.Cancel;
              this.Close();
          }
      }

      private void TxtSearch_TextChanged(object sender, EventArgs e)
      {
          string filter = txtSearch.Text.ToLower();
          filteredSectionTypes = sectionTypes.FindAll(x => x.Name.ToLower().Contains(filter));
          UpdateListBox();
      }

      private void UpdateListBox()
      {
          lstSectionTypes.Items.Clear();
          foreach (var st in filteredSectionTypes)
          {
              lstSectionTypes.Items.Add($"{st.Id.IntegerValue} - {st.Name}");
          }
          if (lstSectionTypes.Items.Count > 0)
              lstSectionTypes.SelectedIndex = 0;
      }

      private void LoadSettings()
      {
          string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
          string folder = Path.Combine(appData, "revit-scripts");
          if (!Directory.Exists(folder))
              Directory.CreateDirectory(folder);
          settingsFilePath = Path.Combine(folder, $"CreateSectionsThroughSelectedElements_{currentProjectName}.txt");
          if (File.Exists(settingsFilePath))
          {
              try
              {
                  var lines = File.ReadAllLines(settingsFilePath);
                  Dictionary<string, string> settings = lines
                      .Select(line => line.Split(new char[] { '=' }, 2))
                      .Where(parts => parts.Length == 2)
                      .ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim());
                  if (settings.ContainsKey("ProjectName") && settings["ProjectName"] != currentProjectName)
                      return;
                  if (settings.ContainsKey("SelectedSectionTypeId"))
                  {
                      int id = int.Parse(settings["SelectedSectionTypeId"]);
                      int index = sectionTypes.FindIndex(x => x.Id.IntegerValue == id);
                      if (index >= 0)
                      {
                          var st = sectionTypes[index];
                          int listIndex = filteredSectionTypes.IndexOf(st);
                          if (listIndex >= 0)
                              lstSectionTypes.SelectedIndex = listIndex;
                      }
                  }
                  if (settings.ContainsKey("OrientationMode"))
                  {
                      string mode = settings["OrientationMode"];
                      if (mode == "Element")
                          rbElements.Checked = true;
                      else if (mode == "Host")
                          rbHosts.Checked = true;
                      else if (mode == "CurrentView")
                          rbCurrentView.Checked = true;
                  }
                  if (settings.ContainsKey("LocalOrientation"))
                  {
                      string lo = settings["LocalOrientation"];
                      if (cmbLocalOrientation.Items.Contains(lo))
                          cmbLocalOrientation.SelectedItem = lo;
                  }
              }
              catch { }
          }
      }

      public void SaveSettings()
      {
          try
          {
              List<string> lines = new List<string>();
              if (lstSectionTypes.SelectedIndex >= 0)
              {
                  var st = filteredSectionTypes[lstSectionTypes.SelectedIndex];
                  lines.Add("SelectedSectionTypeId=" + st.Id.IntegerValue.ToString());
              }
              string mode = rbElements.Checked ? "Element" : rbHosts.Checked ? "Host" : "CurrentView";
              lines.Add("OrientationMode=" + mode);
              lines.Add("LocalOrientation=" + cmbLocalOrientation.SelectedItem.ToString());
              lines.Add("ProjectName=" + currentProjectName);
              File.WriteAllLines(settingsFilePath, lines);
          }
          catch { }
      }

      protected override void OnFormClosing(WinForms.FormClosingEventArgs e)
      {
          if (this.DialogResult == WinForms.DialogResult.OK)
          {
              SectionSettings settings = new SectionSettings();
              if (lstSectionTypes.SelectedIndex >= 0)
              {
                  var st = filteredSectionTypes[lstSectionTypes.SelectedIndex];
                  settings.SelectedSectionTypeId = st.Id;
              }
              else
              {
                  settings.SelectedSectionTypeId = ElementId.InvalidElementId;
              }
              settings.OrientationMode = rbElements.Checked ? "Element" : rbHosts.Checked ? "Host" : "CurrentView";
              settings.LocalOrientation = cmbLocalOrientation.SelectedItem.ToString();
              settings.ProjectName = currentProjectName;
              ResultSettings = settings;
              SaveSettings();
          }
          base.OnFormClosing(e);
      }
  }

  public class CreateSectionsThroughSelectedElements : IExternalCommand
  {
      // Helper: Rotate baseDir about Z by the angle specified in localOption.
      private XYZ RotateVectorWithLocal(XYZ baseDir, string localOption)
      {
          // localOption is now one of: "0°", "90°", "180°", "270°"
          string anglePart = localOption.Replace("°", "");
          if (!double.TryParse(anglePart, out double degrees))
              return baseDir;
          double radians = degrees * Math.PI / 180.0;
          double cos = Math.Cos(radians);
          double sin = Math.Sin(radians);
          XYZ final = new XYZ(baseDir.X * cos - baseDir.Y * sin, baseDir.X * sin + baseDir.Y * cos, baseDir.Z);
          return final.Normalize();
      }

      public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
      {
          UIDocument uidoc = commandData.Application.ActiveUIDocument;
          Document doc = uidoc.Document;
          View activeView = doc.ActiveView;

          List<ViewFamilyType> sectionTypes = new FilteredElementCollector(doc)
              .OfClass(typeof(ViewFamilyType))
              .Cast<ViewFamilyType>()
              .Where(x => x.ViewFamily == ViewFamily.Section)
              .ToList();
          if (!sectionTypes.Any())
          {
              message = "No Section view types available.";
              return Result.Failed;
          }

          ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
          bool hasHosted = selIds.Select(id => doc.GetElement(id))
              .OfType<FamilyInstance>()
              .Any(fi => fi.Host != null);
          string projectName = doc.Title;

          SectionSettingsForm settingsForm = new SectionSettingsForm(sectionTypes, hasHosted, projectName);
          WinForms.DialogResult dr = settingsForm.ShowDialog();
          if (dr != WinForms.DialogResult.OK || settingsForm.ResultSettings == null)
          {
              message = "Operation cancelled.";
              return Result.Cancelled;
          }
          SectionSettings settings = settingsForm.ResultSettings;
          if (selIds.Count == 0)
          {
              message = "Please select at least one element.";
              return Result.Failed;
          }

          using (Transaction tx = new Transaction(doc, "Create Sections Through Selected Elements"))
          {
              tx.Start();
              foreach (ElementId id in selIds)
              {
                  Element elem = doc.GetElement(id);
                  BoundingBoxXYZ bbox = elem.get_BoundingBox(null) ?? elem.get_BoundingBox(activeView);
                  XYZ center = null;
                  double elementRotation = 0.0;
                  if (elem.Location is LocationPoint lp)
                  {
                      center = lp.Point;
                      elementRotation = lp.Rotation;
                  }
                  else if (elem.Location is LocationCurve lc)
                  {
                      center = lc.Curve.Evaluate(0.5, true);
                      XYZ curveDir = (lc.Curve.GetEndPoint(1) - lc.Curve.GetEndPoint(0)).Normalize();
                      elementRotation = Math.Atan2(curveDir.Y, curveDir.X);
                  }
                  else if (bbox != null)
                  {
                      XYZ localCenter = (bbox.Min + bbox.Max) * 0.5;
                      center = bbox.Transform.OfPoint(localCenter);
                  }
                  else continue;

                  XYZ sectionUp = XYZ.BasisZ;
                  XYZ sectionViewDir = null;
                  XYZ sectionRight = null;
                  if (settings.OrientationMode == "CurrentView")
                  {
                      XYZ currentHorizontal = new XYZ(activeView.ViewDirection.X, activeView.ViewDirection.Y, 0);
                      if (currentHorizontal.IsAlmostEqualTo(XYZ.Zero))
                          currentHorizontal = XYZ.BasisX;
                      sectionViewDir = currentHorizontal.Normalize();
                      sectionViewDir = RotateVectorWithLocal(sectionViewDir, settings.LocalOrientation);
                      sectionRight = sectionUp.CrossProduct(sectionViewDir).Normalize();
                  }
                  else if (settings.OrientationMode == "Host")
                  {
                      FamilyInstance fi = elem as FamilyInstance;
                      if (fi != null && fi.Host != null)
                      {
                          Element host = fi.Host;
                          if (host is Wall hostWall)
                          {
                              LocationCurve hostLoc = hostWall.Location as LocationCurve;
                              if (hostLoc != null)
                              {
                                  XYZ start = hostLoc.Curve.GetEndPoint(0);
                                  XYZ end = hostLoc.Curve.GetEndPoint(1);
                                  XYZ wallAxis = (end - start).Normalize();
                                  // For a proper cut, use the perpendicular to the wall's centerline.
                                  XYZ baseDir = new XYZ(-wallAxis.Y, wallAxis.X, 0);
                                  sectionViewDir = RotateVectorWithLocal(baseDir, settings.LocalOrientation);
                                  sectionRight = sectionUp.CrossProduct(sectionViewDir).Normalize();
                              }
                          }
                          else
                          {
                              sectionViewDir = new XYZ(Math.Cos(elementRotation), Math.Sin(elementRotation), 0);
                              sectionViewDir = RotateVectorWithLocal(sectionViewDir, settings.LocalOrientation);
                              sectionRight = sectionUp.CrossProduct(sectionViewDir).Normalize();
                          }
                      }
                      else
                      {
                          sectionViewDir = new XYZ(Math.Cos(elementRotation), Math.Sin(elementRotation), 0);
                          sectionViewDir = RotateVectorWithLocal(sectionViewDir, settings.LocalOrientation);
                          sectionRight = sectionUp.CrossProduct(sectionViewDir).Normalize();
                      }
                  }
                  else // "Element" mode.
                  {
                      if (elem.Category != null && elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Walls)
                      {
                          Wall wall = elem as Wall;
                          if (wall != null)
                          {
                              LocationCurve wallLoc = wall.Location as LocationCurve;
                              if (wallLoc != null)
                              {
                                  XYZ start = wallLoc.Curve.GetEndPoint(0);
                                  XYZ end = wallLoc.Curve.GetEndPoint(1);
                                  XYZ wallAxis = (end - start).Normalize();
                                  // Use the perpendicular to the wall's centerline.
                                  XYZ baseDir = new XYZ(-wallAxis.Y, wallAxis.X, 0);
                                  sectionViewDir = RotateVectorWithLocal(baseDir, settings.LocalOrientation);
                                  sectionRight = sectionUp.CrossProduct(sectionViewDir).Normalize();
                              }
                          }
                          else
                          {
                              sectionViewDir = new XYZ(Math.Cos(elementRotation), Math.Sin(elementRotation), 0);
                              sectionViewDir = RotateVectorWithLocal(sectionViewDir, settings.LocalOrientation);
                              sectionRight = sectionUp.CrossProduct(sectionViewDir).Normalize();
                          }
                      }
                      else
                      {
                          sectionViewDir = new XYZ(Math.Cos(elementRotation), Math.Sin(elementRotation), 0);
                          sectionViewDir = RotateVectorWithLocal(sectionViewDir, settings.LocalOrientation);
                          sectionRight = sectionUp.CrossProduct(sectionViewDir).Normalize();
                      }
                  }

                  Transform sectionTransform = Transform.Identity;
                  sectionTransform.Origin = center;
                  sectionTransform.BasisX = sectionRight;
                  sectionTransform.BasisY = sectionUp;
                  sectionTransform.BasisZ = sectionViewDir;

                  List<XYZ> worldCorners = new List<XYZ>();
                  if (bbox != null)
                  {
                      for (int ix = 0; ix < 2; ix++)
                      {
                          double x = (ix == 0) ? bbox.Min.X : bbox.Max.X;
                          for (int iy = 0; iy < 2; iy++)
                          {
                              double y = (iy == 0) ? bbox.Min.Y : bbox.Max.Y;
                              for (int iz = 0; iz < 2; iz++)
                              {
                                  double z = (iz == 0) ? bbox.Min.Z : bbox.Max.Z;
                                  XYZ localPt = new XYZ(x, y, z);
                                  XYZ worldPt = bbox.Transform.OfPoint(localPt);
                                  worldCorners.Add(worldPt);
                              }
                          }
                      }
                  }
                  else
                  {
                      worldCorners.Add(center);
                  }
                  List<XYZ> localPoints = worldCorners.Select(pt => sectionTransform.Inverse.OfPoint(pt)).ToList();
                  double minX = localPoints.Min(pt => pt.X);
                  double maxX = localPoints.Max(pt => pt.X);
                  double minY = localPoints.Min(pt => pt.Y);
                  double maxY = localPoints.Max(pt => pt.Y);
                  double width = maxX - minX;
                  double height = maxY - minY;
                  double horizontalMargin = width * 0.03;
                  double verticalMargin = height * 0.03;
                  if (elem.Category != null && elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Walls)
                  {
                      Parameter wallWidthParam = elem.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                      double wallThickness = (wallWidthParam != null) ? wallWidthParam.AsDouble() : width;
                      double effectiveHalfWidth = (wallThickness * (1 + 2 * 0.03)) / 2.0;
                      minX = -effectiveHalfWidth;
                      maxX = effectiveHalfWidth;
                      minY -= verticalMargin;
                      maxY += verticalMargin;
                  }
                  else
                  {
                      minX -= horizontalMargin;
                      maxX += horizontalMargin;
                      minY -= verticalMargin;
                      maxY += verticalMargin;
                  }
                  double cropMinZ = -0.5;
                  double cropMaxZ = 0.5;
                  BoundingBoxXYZ sectionBB = new BoundingBoxXYZ
                  {
                      Transform = sectionTransform,
                      Min = new XYZ(minX, minY, cropMinZ),
                      Max = new XYZ(maxX, maxY, cropMaxZ)
                  };

                  ViewSection sectionView = ViewSection.CreateSection(doc, settings.SelectedSectionTypeId, sectionBB);
                  Parameter hideParam = sectionView.LookupParameter("Hide at scales coarser than");
                  if (hideParam != null && !hideParam.IsReadOnly)
                      hideParam.Set(activeView.Scale);

                  double maxAbsZ = localPoints.Max(pt => Math.Abs(pt.Z));
                  Parameter farClipParam = sectionView.LookupParameter("Far Clip Offset");
                  if (farClipParam != null && !farClipParam.IsReadOnly)
                      farClipParam.Set(maxAbsZ);

                  string typeName = "Unknown";
                  ElementId typeId = elem.GetTypeId();
                  if (typeId != null)
                  {
                      Element typeElem = doc.GetElement(typeId);
                      if (typeElem != null)
                          typeName = typeElem.Name;
                  }
                  string baseName = $"Section {typeName}";
                  List<View> existingViews = new FilteredElementCollector(doc)
                      .OfClass(typeof(View))
                      .Cast<View>()
                      .Where(v => v.Name.StartsWith(baseName, StringComparison.OrdinalIgnoreCase))
                      .ToList();
                  string finalName;
                  if (!existingViews.Any(v => v.Name.Equals(baseName, StringComparison.OrdinalIgnoreCase)))
                      finalName = baseName;
                  else
                  {
                      int maxOrdinal = 1;
                      foreach (View v in existingViews)
                      {
                          string name = v.Name;
                          if (name.Equals(baseName, StringComparison.OrdinalIgnoreCase))
                              maxOrdinal = Math.Max(maxOrdinal, 1);
                          else if (name.StartsWith(baseName + " ", StringComparison.OrdinalIgnoreCase))
                          {
                              string ordinalStr = name.Substring((baseName + " ").Length).Trim();
                              if (int.TryParse(ordinalStr, out int ordinal))
                                  maxOrdinal = Math.Max(maxOrdinal, ordinal);
                          }
                      }
                      finalName = $"{baseName} {maxOrdinal + 1}";
                  }
                  sectionView.Name = finalName;
              }
              tx.Commit();
          }
          return Result.Succeeded;
      }
  }
}
