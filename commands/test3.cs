using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Diagnostics;
using System.Text;

[Transaction(TransactionMode.Manual)]
public class CreateLines : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData, 
      ref string message, 
      ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Check if the document is a family document
        if (!doc.IsFamilyDocument)
        {
            message = "This command can only be executed in a family document.";
            return Result.Failed;
        }

        // Collect start time for overall operation
        Stopwatch totalStopwatch = new Stopwatch();
        totalStopwatch.Start();

        // Store the times for each line creation
        double[] lineCreationTimes = new double[500];
        StringBuilder details = new StringBuilder();

        Random random = new Random();

        using (Transaction trans = new Transaction(doc, "Create Lines"))
        {
            trans.Start();

            // Create 1500 lines
            for (int i = 0; i < 200; i++)
            {
                Stopwatch lineStopwatch = new Stopwatch();
                lineStopwatch.Start();

                // Define random start and end points for the line
                double startX = random.NextDouble() * 100;
                double startY = random.NextDouble() * 100;
                double endX = startX + random.NextDouble() * 10;
                double endY = startY + random.NextDouble() * 10;

                XYZ startPoint = new XYZ(startX, startY, 0);
                XYZ endPoint = new XYZ(endX, endY, 0);

                // Create the line
                Line line = Line.CreateBound(startPoint, endPoint);
                CurveElement curveElement = doc.FamilyCreate.NewDetailCurve(doc.ActiveView, line);

                // Stop the stopwatch for this line creation
                lineStopwatch.Stop();
                // Store the elapsed time in milliseconds
                lineCreationTimes[i] = lineStopwatch.Elapsed.TotalMilliseconds;

                // Append the details for this line to the StringBuilder
                details.AppendLine($"Line {i + 1}: {lineCreationTimes[i]} ms");
            }

            trans.Commit();
        }

        // Stop the overall stopwatch
        totalStopwatch.Stop();
        double totalElapsedTime = totalStopwatch.Elapsed.TotalMilliseconds;

        // Append the total time to the StringBuilder
        details.AppendLine($"Total time to create 1500 lines: {totalElapsedTime} ms");

        // Display the detailed results to the user
        TaskDialog.Show("Line Creation Times", details.ToString());

        return Result.Succeeded;
    }
}
