using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Forms = System.Windows.Forms;
using System.Drawing;

namespace RevitCustomCommands
{
    [Transaction(TransactionMode.Manual)]
    public class OutlineSelectedElementsInCurrentViewWithDetailLines : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Get the Revit application and document
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;
                View activeView = doc.ActiveView;

                // Check if the active view is valid for detail lines
                if (!(activeView.ViewType == ViewType.FloorPlan || 
                      activeView.ViewType == ViewType.CeilingPlan || 
                      activeView.ViewType == ViewType.Section || 
                      activeView.ViewType == ViewType.Elevation || 
                      activeView is View3D || 
                      activeView is ViewSheet || 
                      activeView is ViewDrafting))
                {
                    TaskDialog.Show("Error", "Detail lines cannot be created in this view type.");
                    return Result.Failed;
                }

                // Get selected elements
                ICollection<ElementId> selectedElementIds = uidoc.Selection.GetElementIds();
                
                if (selectedElementIds.Count == 0)
                {
                    TaskDialog.Show("No Selection", "Please select at least one element to outline.");
                    return Result.Failed;
                }

                // Prompt user for offset value in millimeters
                double offsetMm = PromptForOffset();
                if (offsetMm < 0)
                {
                    // User cancelled
                    return Result.Cancelled;
                }
                
                // Convert offset from mm to feet (Revit's internal unit)
                double offsetFeet = offsetMm / 304.8;

                // Get all line styles in the document
                FilteredElementCollector lineStyleCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(GraphicsStyle))
                    .WhereElementIsNotElementType();

                List<GraphicsStyle> lineStyles = lineStyleCollector
                    .Cast<GraphicsStyle>()
                    .Where(gs => gs.GraphicsStyleCategory.Parent != null &&
                                gs.GraphicsStyleCategory.Parent.Id.IntegerValue == (int)BuiltInCategory.OST_Lines)
                    .ToList();

                if (lineStyles.Count == 0)
                {
                    TaskDialog.Show("Error", "No line styles found in the document.");
                    return Result.Failed;
                }

                // Prepare data for line style selection
                List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
                
                foreach (GraphicsStyle lineStyle in lineStyles)
                {
                    Dictionary<string, object> entry = new Dictionary<string, object>
                    {
                        { "Title", lineStyle.Name },
                        { "SheetFolder", lineStyle.GraphicsStyleCategory.Name },
                        { "Id", lineStyle.Id.IntegerValue.ToString() }
                    };
                    entries.Add(entry);
                }

                // Prompt user to select line style
                List<string> propertyNames = new List<string> { "Title", "SheetFolder", "Id" };
                List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);
                
                if (selectedEntries == null || selectedEntries.Count == 0)
                {
                    TaskDialog.Show("Outline Elements", "No line style was selected.");
                    return Result.Cancelled;
                }

                // Get selected line style
                int selectedLineStyleId = Convert.ToInt32(selectedEntries[0]["Id"]);
                GraphicsStyle selectedLineStyle = lineStyles.FirstOrDefault(ls => ls.Id.IntegerValue == selectedLineStyleId);
                
                if (selectedLineStyle == null)
                {
                    TaskDialog.Show("Error", "Selected line style not found.");
                    return Result.Failed;
                }

                // Start transaction
                using (Transaction trans = new Transaction(doc, "Outline Selected Elements with Detail Lines"))
                {
                    trans.Start();
                    
                    int linesCreated = 0;
                    
                    // Process each selected element
                    foreach (ElementId elementId in selectedElementIds)
                    {
                        Element elem = doc.GetElement(elementId);
                        
                        // Get element's bounding box in the current view
                        BoundingBoxXYZ bbox = elem.get_BoundingBox(activeView);
                        
                        if (bbox == null) continue;
                        
                        // Apply offset to bounding box
                        bbox.Min = new XYZ(bbox.Min.X - offsetFeet, bbox.Min.Y - offsetFeet, bbox.Min.Z - offsetFeet);
                        bbox.Max = new XYZ(bbox.Max.X + offsetFeet, bbox.Max.Y + offsetFeet, bbox.Max.Z + offsetFeet);
                        
                        // Create rectangle from bounding box based on view orientation
                        List<Curve> rectangleLines = CreateRectangleFromBoundingBox(bbox, activeView);
                        
                        // Create detail lines for each rectangle line
                        foreach (Curve curve in rectangleLines)
                        {
                            if (curve != null && curve.Length > 0)
                            {
                                DetailCurve detailLine = doc.Create.NewDetailCurve(activeView, curve);
                                
                                // Apply the selected line style
                                if (detailLine != null && detailLine.LineStyle != null)
                                {
                                    detailLine.LineStyle = selectedLineStyle;
                                    linesCreated++;
                                }
                            }
                        }
                    }
                    
                    trans.Commit();
                    
                    TaskDialog.Show("Outline Complete", $"Created {linesCreated} detail lines to outline selected elements with {offsetMm}mm offset.");
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        /// <summary>
        /// Prompts user for offset value in millimeters
        /// </summary>
        /// <returns>Offset value in millimeters; -1 if cancelled</returns>
        private double PromptForOffset()
        {
            using (Forms.Form promptForm = new Forms.Form())
            {
                promptForm.Width = 350;
                promptForm.Height = 150;
                promptForm.Text = "Specify Outline Offset";
                promptForm.StartPosition = Forms.FormStartPosition.CenterScreen;
                promptForm.FormBorderStyle = Forms.FormBorderStyle.FixedDialog;
                promptForm.MaximizeBox = false;
                promptForm.MinimizeBox = false;

                Forms.Label textLabel = new Forms.Label() { Left = 20, Top = 20, Width = 300, Text = "Enter offset value in millimeters:" };
                Forms.NumericUpDown inputBox = new Forms.NumericUpDown() { Left = 20, Top = 45, Width = 120 };
                inputBox.Minimum = 0;
                inputBox.Maximum = 1000;
                inputBox.DecimalPlaces = 2;
                inputBox.Value = 10; // Default value
                
                Forms.Button confirmButton = new Forms.Button() { Text = "OK", Left = 150, Width = 80, Top = 80, DialogResult = Forms.DialogResult.OK };
                Forms.Button cancelButton = new Forms.Button() { Text = "Cancel", Left = 240, Width = 80, Top = 80, DialogResult = Forms.DialogResult.Cancel };
                
                confirmButton.Click += (sender, e) => { promptForm.Close(); };
                cancelButton.Click += (sender, e) => { promptForm.Close(); };
                
                promptForm.Controls.Add(textLabel);
                promptForm.Controls.Add(inputBox);
                promptForm.Controls.Add(confirmButton);
                promptForm.Controls.Add(cancelButton);
                promptForm.AcceptButton = confirmButton;
                promptForm.CancelButton = cancelButton;

                if (promptForm.ShowDialog() == Forms.DialogResult.OK)
                {
                    return (double)inputBox.Value;
                }
                else
                {
                    return -1; // Cancelled
                }
            }
        }

        /// <summary>
        /// Creates rectangle curves from a bounding box based on view orientation
        /// </summary>
        private List<Curve> CreateRectangleFromBoundingBox(BoundingBoxXYZ bbox, View view)
        {
            List<Curve> rectangleLines = new List<Curve>();
            
            XYZ min = bbox.Min;
            XYZ max = bbox.Max;
            
            if (view is View3D)
            {
                // For 3D views, create a full 3D box (12 lines)
                // Bottom rectangle
                rectangleLines.Add(Line.CreateBound(new XYZ(min.X, min.Y, min.Z), new XYZ(max.X, min.Y, min.Z)));
                rectangleLines.Add(Line.CreateBound(new XYZ(max.X, min.Y, min.Z), new XYZ(max.X, max.Y, min.Z)));
                rectangleLines.Add(Line.CreateBound(new XYZ(max.X, max.Y, min.Z), new XYZ(min.X, max.Y, min.Z)));
                rectangleLines.Add(Line.CreateBound(new XYZ(min.X, max.Y, min.Z), new XYZ(min.X, min.Y, min.Z)));
                
                // Top rectangle
                rectangleLines.Add(Line.CreateBound(new XYZ(min.X, min.Y, max.Z), new XYZ(max.X, min.Y, max.Z)));
                rectangleLines.Add(Line.CreateBound(new XYZ(max.X, min.Y, max.Z), new XYZ(max.X, max.Y, max.Z)));
                rectangleLines.Add(Line.CreateBound(new XYZ(max.X, max.Y, max.Z), new XYZ(min.X, max.Y, max.Z)));
                rectangleLines.Add(Line.CreateBound(new XYZ(min.X, max.Y, max.Z), new XYZ(min.X, min.Y, max.Z)));
                
                // Vertical lines
                rectangleLines.Add(Line.CreateBound(new XYZ(min.X, min.Y, min.Z), new XYZ(min.X, min.Y, max.Z)));
                rectangleLines.Add(Line.CreateBound(new XYZ(max.X, min.Y, min.Z), new XYZ(max.X, min.Y, max.Z)));
                rectangleLines.Add(Line.CreateBound(new XYZ(max.X, max.Y, min.Z), new XYZ(max.X, max.Y, max.Z)));
                rectangleLines.Add(Line.CreateBound(new XYZ(min.X, max.Y, min.Z), new XYZ(min.X, max.Y, max.Z)));
            }
            else if (view.ViewType == ViewType.FloorPlan || view.ViewType == ViewType.CeilingPlan)
            {
                // For plan views, create a horizontal rectangle
                rectangleLines.Add(Line.CreateBound(new XYZ(min.X, min.Y, 0), new XYZ(max.X, min.Y, 0)));
                rectangleLines.Add(Line.CreateBound(new XYZ(max.X, min.Y, 0), new XYZ(max.X, max.Y, 0)));
                rectangleLines.Add(Line.CreateBound(new XYZ(max.X, max.Y, 0), new XYZ(min.X, max.Y, 0)));
                rectangleLines.Add(Line.CreateBound(new XYZ(min.X, max.Y, 0), new XYZ(min.X, min.Y, 0)));
            }
            else if (view.ViewType == ViewType.Elevation || view.ViewType == ViewType.Section)
            {
                // For elevations and sections, need to determine which plane to use
                // based on view direction
                XYZ viewDir = view.ViewDirection;
                
                if (Math.Abs(viewDir.X) > Math.Abs(viewDir.Y) && Math.Abs(viewDir.X) > Math.Abs(viewDir.Z))
                {
                    // Looking along X-axis (YZ plane)
                    double x = (viewDir.X > 0) ? min.X : max.X;
                    rectangleLines.Add(Line.CreateBound(new XYZ(x, min.Y, min.Z), new XYZ(x, max.Y, min.Z)));
                    rectangleLines.Add(Line.CreateBound(new XYZ(x, max.Y, min.Z), new XYZ(x, max.Y, max.Z)));
                    rectangleLines.Add(Line.CreateBound(new XYZ(x, max.Y, max.Z), new XYZ(x, min.Y, max.Z)));
                    rectangleLines.Add(Line.CreateBound(new XYZ(x, min.Y, max.Z), new XYZ(x, min.Y, min.Z)));
                }
                else if (Math.Abs(viewDir.Y) > Math.Abs(viewDir.X) && Math.Abs(viewDir.Y) > Math.Abs(viewDir.Z))
                {
                    // Looking along Y-axis (XZ plane)
                    double y = (viewDir.Y > 0) ? min.Y : max.Y;
                    rectangleLines.Add(Line.CreateBound(new XYZ(min.X, y, min.Z), new XYZ(max.X, y, min.Z)));
                    rectangleLines.Add(Line.CreateBound(new XYZ(max.X, y, min.Z), new XYZ(max.X, y, max.Z)));
                    rectangleLines.Add(Line.CreateBound(new XYZ(max.X, y, max.Z), new XYZ(min.X, y, max.Z)));
                    rectangleLines.Add(Line.CreateBound(new XYZ(min.X, y, max.Z), new XYZ(min.X, y, min.Z)));
                }
                else
                {
                    // Looking along Z-axis (XY plane)
                    double z = (viewDir.Z > 0) ? min.Z : max.Z;
                    rectangleLines.Add(Line.CreateBound(new XYZ(min.X, min.Y, z), new XYZ(max.X, min.Y, z)));
                    rectangleLines.Add(Line.CreateBound(new XYZ(max.X, min.Y, z), new XYZ(max.X, max.Y, z)));
                    rectangleLines.Add(Line.CreateBound(new XYZ(max.X, max.Y, z), new XYZ(min.X, max.Y, z)));
                    rectangleLines.Add(Line.CreateBound(new XYZ(min.X, max.Y, z), new XYZ(min.X, min.Y, z)));
                }
            }
            else
            {
                // For all other views, create a simple rectangle in the XY plane
                rectangleLines.Add(Line.CreateBound(new XYZ(min.X, min.Y, 0), new XYZ(max.X, min.Y, 0)));
                rectangleLines.Add(Line.CreateBound(new XYZ(max.X, min.Y, 0), new XYZ(max.X, max.Y, 0)));
                rectangleLines.Add(Line.CreateBound(new XYZ(max.X, max.Y, 0), new XYZ(min.X, max.Y, 0)));
                rectangleLines.Add(Line.CreateBound(new XYZ(min.X, max.Y, 0), new XYZ(min.X, min.Y, 0)));
            }
            
            return rectangleLines;
        }
    }
}
