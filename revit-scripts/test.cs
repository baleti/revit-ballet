using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

[Transaction(TransactionMode.Manual)]
public class MoveDimensionText : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        // Get the current Revit document
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;

        TaskDialog.Show("Debug", "Starting transaction...");

        // Start a transaction
        using (Transaction trans = new Transaction(doc, "Move Dimension Text"))
        {
            trans.Start();

            TaskDialog.Show("Debug", "Transaction started.");

            // Filter to get only dimensions in the document
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(Dimension));

            TaskDialog.Show("Debug", "Filtered dimensions.");

            // Define the offset value in Revit internal units (feet)
            double offset = 1000;

            TaskDialog.Show("Debug", $"Offset value in feet: {offset}");

            foreach (Dimension dimension in collector)
            {
                TaskDialog.Show("Debug", $"Processing dimension ID: {dimension.Id}");

                // Check if the dimension has an overridden text position
                if (dimension.TextPosition != null)
                {
                    // Get the current text position
                    XYZ currentTextPosition = dimension.TextPosition;

                    TaskDialog.Show("Debug", $"Current text position: {currentTextPosition}");

                    // Calculate the new text position
                    XYZ newTextPosition = new XYZ(currentTextPosition.X, currentTextPosition.Y + offset, currentTextPosition.Z);

                    TaskDialog.Show("Debug", $"New text position: {newTextPosition}");

                    // Set the new text position
                    dimension.TextPosition = newTextPosition;

                    TaskDialog.Show("Debug", "Text position updated.");
                }
                else
                {
                    TaskDialog.Show("Debug", "Dimension does not have a text position.");
                }
            }

            // Commit the transaction
            trans.Commit();

            TaskDialog.Show("Debug", "Transaction committed.");
        }

        TaskDialog.Show("Debug", "Finished processing.");

        return Result.Succeeded;
    }
}
