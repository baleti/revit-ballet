using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;

[Transaction(TransactionMode.Manual)]
public class SectionBox3DFromSelectedView : IExternalCommand
{
    private bool GetZoomToFitSetting()
    {
        string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string configDir = Path.Combine(appDataPath, "revit-scripts");
        string configFile = Path.Combine(configDir, "SectionBox3DFromView");
        
        // Create directory if it doesn't exist
        if (!Directory.Exists(configDir))
        {
            Directory.CreateDirectory(configDir);
        }
        
        // Create config file with default if it doesn't exist
        if (!File.Exists(configFile))
        {
            File.WriteAllText(configFile, "ZoomToFit = True");
            return true;
        }
        
        // Read the config file
        try
        {
            string content = File.ReadAllText(configFile);
            string[] lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (string line in lines)
            {
                if (line.Trim().StartsWith("ZoomToFit", StringComparison.OrdinalIgnoreCase))
                {
                    string[] parts = line.Split('=');
                    if (parts.Length >= 2)
                    {
                        string value = parts[1].Trim().ToLower();
                        return value == "true" || value == "1";
                    }
                }
            }
        }
        catch
        {
            // If any error reading file, default to true
            return true;
        }
        
        // Default to true if setting not found
        return true;
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;
        View currentView = doc.ActiveView;
        View targetView = null;

        // First, check if exactly one element is selected.
        var selectedIds = uiDoc.GetSelectionIds();
        if (selectedIds.Count == 1)
        {
            Element selElem = doc.GetElement(selectedIds.First());
            // If the element belongs to the OST_Viewers category, retrieve its corresponding view using the VIEW_NAME parameter.
            if (selElem.Category != null &&
                selElem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Viewers)
            {
                Parameter nameParam = selElem.get_Parameter(BuiltInParameter.VIEW_NAME);
                if (nameParam != null)
                {
                    string viewName = nameParam.AsString();
                    if (!string.IsNullOrEmpty(viewName))
                    {
                        // Find the view with the given name (ignoring case).
                        targetView = new FilteredElementCollector(doc)
                                        .OfClass(typeof(View))
                                        .Cast<View>()
                                        .FirstOrDefault(v => v.Name.Equals(viewName, StringComparison.OrdinalIgnoreCase));
                    }
                }
            }
            // Otherwise, if the selected element is a View (but not a sheet), use it.
            else if (selElem is View && !(selElem is ViewSheet))
            {
                targetView = selElem as View;
            }
        }

        // If no valid selection was found, fall back to existing logic.
        if (targetView == null)
        {
            // If we're on a sheet, get the viewport's view.
            if (currentView.ViewType == ViewType.DrawingSheet)
            {
                var viewports = selectedIds.Select(id => doc.GetElement(id))
                                           .OfType<Viewport>()
                                           .ToList();
                if (viewports.Count != 1)
                {
                    TaskDialog.Show("Error", "Please select a single viewport on the sheet.");
                    return Result.Failed;
                }
                targetView = doc.GetElement(viewports.First().ViewId) as View;
            }
            // Otherwise, if the active view has a crop box, use it.
            else if (currentView.CropBoxActive)
            {
                targetView = currentView;
            }
            else
            {
                TaskDialog.Show("Error", "Active view does not have an active crop region.");
                return Result.Failed;
            }
        }

        // Check if the target view is a 3D view with an active section box
        BoundingBoxXYZ sectionBox = null;
        
        if (targetView is View3D target3DView && target3DView.IsSectionBoxActive)
        {
            // Use the section box from the target 3D view directly
            sectionBox = target3DView.GetSectionBox();
        }
        else
        {
            // Original logic for non-3D views
            if (!targetView.CropBoxActive)
            {
                TaskDialog.Show("Error", "The target view does not have an active crop region.");
                return Result.Failed;
            }

            // Get the crop box from the target view.
            BoundingBoxXYZ cropBox = targetView.CropBox;

            // Convert the crop box's eight corners from its local coordinate system to world coordinates.
            List<XYZ> worldPoints = new List<XYZ>();
            foreach (double x in new double[] { cropBox.Min.X, cropBox.Max.X })
            {
                foreach (double y in new double[] { cropBox.Min.Y, cropBox.Max.Y })
                {
                    foreach (double z in new double[] { cropBox.Min.Z, cropBox.Max.Z })
                    {
                        XYZ localPt = new XYZ(x, y, z);
                        XYZ worldPt = cropBox.Transform.OfPoint(localPt);
                        worldPoints.Add(worldPt);
                    }
                }
            }

            // Compute the axis-aligned bounding box in world coordinates.
            double minX = worldPoints.Min(pt => pt.X);
            double minY = worldPoints.Min(pt => pt.Y);
            double minZ = worldPoints.Min(pt => pt.Z);
            double maxX = worldPoints.Max(pt => pt.X);
            double maxY = worldPoints.Max(pt => pt.Y);
            double maxZ = worldPoints.Max(pt => pt.Z);

            // --- Adjust the Z extents if the target view is a plan view using PLAN_VIEW_RANGE ---
            if (targetView is ViewPlan viewPlan)
            {
                // Get the PLAN_VIEW_RANGE parameter. It should return a string with the view range.
                Parameter rangeParam = viewPlan.get_Parameter(BuiltInParameter.PLAN_VIEW_RANGE);
                if (rangeParam != null)
                {
                    string rangeString = rangeParam.AsString();
                    if (!string.IsNullOrEmpty(rangeString))
                    {
                        double viewRangeTop = 0;
                        double viewRangeBottom = 0;
                        string[] parts = rangeString.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string part in parts)
                        {
                            string trimmed = part.Trim();
                            if (trimmed.StartsWith("Top:"))
                            {
                                string value = trimmed.Substring("Top:".Length).Trim();
                                viewRangeTop = ParseFeetInches(value);
                            }
                            else if (trimmed.StartsWith("Bottom:"))
                            {
                                string value = trimmed.Substring("Bottom:".Length).Trim();
                                viewRangeBottom = ParseFeetInches(value);
                            }
                        }
                        // Display a TaskDialog with the parsed view range info.
                        TaskDialog.Show("View Range Info",
                            $"PLAN_VIEW_RANGE string:\n{rangeString}\n\nParsed Values:\nTop: {viewRangeTop} ft\nBottom: {viewRangeBottom} ft");

                        // Override the Z extents computed from the crop region.
                        // (Assuming Top is above Bottom.)
                        minZ = viewRangeBottom;
                        maxZ = viewRangeTop;
                    }
                    else
                    {
                        TaskDialog.Show("View Range Info", "PLAN_VIEW_RANGE parameter is empty.");
                    }
                }
                else
                {
                    TaskDialog.Show("View Range Info", "PLAN_VIEW_RANGE parameter was not found.");
                }
            }
            // --------------------------------------------------------------------------

            // Create a new section box bounding box.
            sectionBox = new BoundingBoxXYZ();
            sectionBox.Min = new XYZ(minX, minY, minZ);
            sectionBox.Max = new XYZ(maxX, maxY, maxZ);
            sectionBox.Transform = Transform.Identity;
        }

        // Locate the default 3D view.
        string default3DViewName = "{3D - " + uiApp.Application.Username + "}";
        View3D default3DView = new FilteredElementCollector(doc)
                                .OfClass(typeof(View3D))
                                .Cast<View3D>()
                                .FirstOrDefault(v => v.Name == default3DViewName);
        if (default3DView == null)
        {
            TaskDialog.Show("Error", "The default 3D view could not be found.");
            return Result.Failed;
        }

        // Set the section box on the default 3D view.
        using (Transaction trans = new Transaction(doc, "Set 3D Section Box"))
        {
            trans.Start();
            default3DView.IsSectionBoxActive = true;
            default3DView.SetSectionBox(sectionBox);
            trans.Commit();
        }

        // Activate the default 3D view.
        uiDoc.ActiveView = default3DView;
        
        // Check configuration setting for zoom to fit
        bool zoomToFit = GetZoomToFitSetting();
        
        if (zoomToFit)
        {
            // Zoom to fit the section box
            UIView uiView = uiDoc
                            .GetOpenUIViews()
                            .FirstOrDefault(v => v.ViewId == default3DView.Id);

            if (uiView != null)
            {
                uiView.ZoomToFit();
            }
        }
        
        return Result.Succeeded;
    }

    /// <summary>
    /// Parses a feet-inches string (e.g., "8'-0\"") into a double (feet).
    /// </summary>
    private double ParseFeetInches(string s)
    {
        if (string.IsNullOrEmpty(s))
            return 0;

        s = s.Trim();
        // Expecting a format like: 8'-0"
        string[] parts = s.Split('\'');
        if (parts.Length < 2)
            return 0;

        double feet = 0;
        double inches = 0;
        double.TryParse(parts[0].Trim(), out feet);

        // The inches part may include extra characters like double quotes or dashes.
        string inchPart = parts[1].Replace("\"", "").Replace("-", "").Trim();
        double.TryParse(inchPart, out inches);

        return feet + (inches / 12.0);
    }
}
