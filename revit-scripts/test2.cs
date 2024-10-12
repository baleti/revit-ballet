using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class CreateDetailLines : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData, 
      ref string message, 
      ElementSet elements)
    {
        // Get the current document
        Document doc = commandData.Application.ActiveUIDocument.Document;

        // Get the first available ViewPlan (could be any appropriate view for placing detail lines)
        ViewPlan viewPlan = new FilteredElementCollector(doc)
                                .OfClass(typeof(ViewPlan))
                                .FirstOrDefault() as ViewPlan;

        if (viewPlan == null)
        {
            message = "No suitable ViewPlan found.";
            return Result.Failed;
        }

        // Start a new transaction
        using (Transaction tx = new Transaction(doc, "Create Detail Lines"))
        {
            tx.Start();

            // Create a new CurveArray to store the detail curves
            CurveArray curveArray = new CurveArray();

            // Initialize random number generator
            Random rand = new Random();

            // Define the range for the random points
            double minX = 0, maxX = 100;
            double minY = 0, maxY = 100;
            double minZ = 0, maxZ = 0; // Keep Z constant for 2D placement in the view

            for (int i = 0; i < 500; i++)
            {
                // Generate random start and end points
                XYZ startPoint = new XYZ(
                    minX + (rand.NextDouble() * (maxX - minX)),
                    minY + (rand.NextDouble() * (maxY - minY)),
                    minZ + (rand.NextDouble() * (maxZ - minZ))
                );

                XYZ endPoint = new XYZ(
                    minX + (rand.NextDouble() * (maxX - minX)),
                    minY + (rand.NextDouble() * (maxY - minY)),
                    minZ + (rand.NextDouble() * (maxZ - minZ))
                );

                // Create a new line
                Line line = Line.CreateBound(startPoint, endPoint);
                curveArray.Append(line);
            }

            // Use NewDetailCurveArray to create all the lines in the active view
            doc.Create.NewDetailCurveArray(viewPlan, curveArray);

            tx.Commit();
        }

        return Result.Succeeded;
    }
}
