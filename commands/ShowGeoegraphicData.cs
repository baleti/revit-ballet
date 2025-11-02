using System;
using System.Text;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace MyRevitAddin
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ShowGeoegraphicData : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            StringBuilder infoBuilder = new StringBuilder();

            // --- Retrieve Survey Point (Shared Base Point) ---
            IList<Element> sharedBasePoints = new FilteredElementCollector(doc)
                                                .OfCategory(BuiltInCategory.OST_SharedBasePoint)
                                                .ToElements();
            if (sharedBasePoints.Count > 0)
            {
                Element surveyPoint = sharedBasePoints[0];
                double spEasting = surveyPoint.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM)?.AsDouble() ?? 0.0;
                double spNorthing = surveyPoint.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM)?.AsDouble() ?? 0.0;
                double spElevation = surveyPoint.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)?.AsDouble() ?? 0.0;

                infoBuilder.AppendLine("Survey Point:");
                infoBuilder.AppendLine($"  Easting: {spEasting}");
                infoBuilder.AppendLine($"  Northing: {spNorthing}");
                infoBuilder.AppendLine($"  Elevation: {spElevation}");
            }
            else
            {
                infoBuilder.AppendLine("Survey Point not found.");
            }

            // --- Retrieve Project Base Point ---
            IList<Element> projectBasePoints = new FilteredElementCollector(doc)
                                                .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                                                .ToElements();
            if (projectBasePoints.Count > 0)
            {
                Element projectBasePoint = projectBasePoints[0];
                double pbpEasting = projectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM)?.AsDouble() ?? 0.0;
                double pbpNorthing = projectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM)?.AsDouble() ?? 0.0;
                double pbpElevation = projectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)?.AsDouble() ?? 0.0;
                double pbpAngle = projectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM)?.AsDouble() ?? 0.0;

                infoBuilder.AppendLine("\nProject Base Point:");
                infoBuilder.AppendLine($"  Easting: {pbpEasting}");
                infoBuilder.AppendLine($"  Northing: {pbpNorthing}");
                infoBuilder.AppendLine($"  Elevation: {pbpElevation}");
                infoBuilder.AppendLine($"  Angle to True North: {pbpAngle}");
            }
            else
            {
                infoBuilder.AppendLine("Project Base Point not found.");
            }

            // --- Retrieve Site Location from Document ---
            SiteLocation docSiteLocation = doc.SiteLocation;
            if (docSiteLocation != null)
            {
                infoBuilder.AppendLine("\nSite Location from Document:");
                infoBuilder.AppendLine($"  Latitude: {docSiteLocation.Latitude}");
                infoBuilder.AppendLine($"  Longitude: {docSiteLocation.Longitude}");
                infoBuilder.AppendLine($"  Elevation: {docSiteLocation.Elevation}");
            }
            else
            {
                infoBuilder.AppendLine("Site Location not available from Document.");
            }

            // --- Retrieve ProjectLocation details (ProjectPosition and SiteLocation) ---
            ProjectLocation projectLocation = doc.ActiveProjectLocation;
            if (projectLocation != null)
            {
                // Get the project position at the origin (0,0,0) in the project coordinate system.
                XYZ origin = new XYZ(0, 0, 0);
                ProjectPosition projPos = projectLocation.GetProjectPosition(origin);
                if (projPos != null)
                {
                    infoBuilder.AppendLine("\nProject Location (ProjectPosition at (0,0,0)):");
                    infoBuilder.AppendLine($"  Angle: {projPos.Angle}");
                    infoBuilder.AppendLine($"  EastWest: {projPos.EastWest}");
                    infoBuilder.AppendLine($"  Elevation: {projPos.Elevation}");
                    infoBuilder.AppendLine($"  NorthSouth: {projPos.NorthSouth}");
                }
                else
                {
                    infoBuilder.AppendLine("No project position at the origin point.");
                }

                // Get the SiteLocation from the ProjectLocation using the API method.
                SiteLocation projSiteLocation = projectLocation.GetSiteLocation();
                if (projSiteLocation != null)
                {
                    // Convert latitude and longitude from radians to degrees for display.
                    const double angleRatio = Math.PI / 180;
                    infoBuilder.AppendLine("\nProject Location's Site Location:");
                    infoBuilder.AppendLine($"  Latitude: {projSiteLocation.Latitude / angleRatio}°");
                    infoBuilder.AppendLine($"  Longitude: {projSiteLocation.Longitude / angleRatio}°");
                    infoBuilder.AppendLine($"  TimeZone: {projSiteLocation.TimeZone}");
                }
                else
                {
                    infoBuilder.AppendLine("Project Location's Site Location not available.");
                }
            }
            else
            {
                infoBuilder.AppendLine("Project Location not available.");
            }

            // --- Display the information ---
            TaskDialog.Show("Location Information", infoBuilder.ToString());

            return Result.Succeeded;
        }
    }
}
