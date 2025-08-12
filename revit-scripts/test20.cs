using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace RevitDetailLines
{
    [Transaction(TransactionMode.Manual)]
    public class DrawXinViewCentreDetailLines : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            View activeView = doc.ActiveView;

            BoundingBoxXYZ viewBBox = activeView.CropBox;
            if (viewBBox == null)
            {
                message = "View does not support cropping or is a template.";
                return Result.Failed;
            }

            try
            {
                // Get view orientation vectors
                XYZ viewRight = activeView.RightDirection;
                XYZ viewUp = activeView.UpDirection;
                
                // Get the view's coordinate transformation
                Transform viewTransform = activeView.CropBox.Transform;

                // Get the untransformed min/max points
                XYZ min = viewBBox.Min;
                XYZ max = viewBBox.Max;

                // Calculate the center point in untransformed coordinates
                XYZ center = new XYZ(
                    (min.X + max.X) / 2,
                    (min.Y + max.Y) / 2,
                    (min.Z + max.Z) / 2
                );

                // Transform the center point to view coordinates
                XYZ boxCenter = viewTransform.OfPoint(center);
                
                // Calculate the size of the X based on view dimensions
                double viewWidth = Math.Abs(max.X - min.X);
                double viewHeight = Math.Abs(max.Y - min.Y);
                double xSize = Math.Min(viewWidth, viewHeight) * 0.2;
                
                // Create the four corners of the X using view orientation vectors
                XYZ topLeft = boxCenter - viewRight * xSize/2 + viewUp * xSize/2;
                XYZ topRight = boxCenter + viewRight * xSize/2 + viewUp * xSize/2;
                XYZ bottomLeft = boxCenter - viewRight * xSize/2 - viewUp * xSize/2;
                XYZ bottomRight = boxCenter + viewRight * xSize/2 - viewUp * xSize/2;

                using (Transaction trans = new Transaction(doc, "Draw X Shape"))
                {
                    trans.Start();

                    // Create the two crossing lines
                    DetailLine line1 = doc.Create.NewDetailCurve(
                        activeView, 
                        Line.CreateBound(topLeft, bottomRight)
                    ) as DetailLine;

                    DetailLine line2 = doc.Create.NewDetailCurve(
                        activeView, 
                        Line.CreateBound(bottomLeft, topRight)
                    ) as DetailLine;

                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
