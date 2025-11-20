using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class RelativeLocationsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get currently selected elements
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("Error", "No elements selected. Please select one or more group instances.");
                    return Result.Failed;
                }

                // Collect data for all elements in selected groups
                List<ElementData> dataList = new List<ElementData>();

                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    Group group = elem as Group;
                    if (group == null)
                    {
                        continue; // Skip if not a group
                    }

                    // Get the group's origin
                    LocationPoint groupLocation = group.Location as LocationPoint;
                    if (groupLocation == null)
                    {
                        continue; // Skip if no point location
                    }
                    XYZ groupOrigin = groupLocation.Point;

                    // Get all member elements
                    ICollection<ElementId> memberIds = group.GetMemberIds();
                    foreach (ElementId memberId in memberIds)
                    {
                        Element member = doc.GetElement(memberId);
                        if (member == null) continue;

                        Location memberLocation = member.Location;
                        XYZ relativePos = null;
                        string locationType = "N/A";

                        if (memberLocation is LocationPoint memberPointLoc)
                        {
                            relativePos = memberPointLoc.Point - groupOrigin;
                            locationType = "Point";
                        }
                        else if (memberLocation is LocationCurve memberCurveLoc)
                        {
                            Curve curve = memberCurveLoc.Curve;
                            if (curve != null)
                            {
                                XYZ start = curve.GetEndPoint(0);
                                XYZ end = curve.GetEndPoint(1);
                                XYZ mid = (start + end) / 2.0;
                                relativePos = mid - groupOrigin;
                                locationType = "Curve Midpoint";
                            }
                        }
                        // Add more location types if needed, e.g., LocationPosition for some elements

                        dataList.Add(new ElementData
                        {
                            GroupName = group.Name,
                            ElementId = (int)member.Id.AsLong(),
                            ElementName = member.Name,
                            Category = member.Category?.Name ?? "N/A",
                            LocationType = locationType,
                            RelativeX = relativePos?.X ?? double.NaN,
                            RelativeY = relativePos?.Y ?? double.NaN,
                            RelativeZ = relativePos?.Z ?? double.NaN
                        });
                    }
                }

                if (dataList.Count == 0)
                {
                    TaskDialog.Show("Info", "No valid group members found.");
                    return Result.Succeeded;
                }

                // Display in a DataGridView
                using (System.Windows.Forms.Form form = new System.Windows.Forms.Form())
                {
                    form.Text = "Element Locations Relative to Group Origins";
                    form.Size = new System.Drawing.Size(800, 600);

                    DataGridView dgv = new DataGridView();
                    dgv.Dock = DockStyle.Fill;
                    dgv.DataSource = dataList;
                    dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                    form.Controls.Add(dgv);
                    form.ShowDialog();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private class ElementData
        {
            public string GroupName { get; set; }
            public int ElementId { get; set; }
            public string ElementName { get; set; }
            public string Category { get; set; }
            public string LocationType { get; set; }
            public double RelativeX { get; set; }
            public double RelativeY { get; set; }
            public double RelativeZ { get; set; }
        }
    }
}
