using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class testCutWithIntersectingElements : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            Selection selection = uiDoc.Selection;

            // Get all currently selected elements
            ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
            if (selectedIds.Count == 0)
            {
                message = "Please select at least one element.";
                return Result.Failed;
            }

            // Collect all walls in the document
            FilteredElementCollector wallCollector =
                new FilteredElementCollector(doc)
                .OfClass(typeof(Wall));

            // Prepare a list to store walls that intersect
            List<ElementId> intersectingWallIds = new List<ElementId>();

            // 1. Identify which walls intersect with the selected elements (via bounding-box overlap).
            foreach (ElementId selId in selectedIds)
            {
                Element selectedElement = doc.GetElement(selId);
                BoundingBoxXYZ selBb = selectedElement?.get_BoundingBox(null);
                if (selBb == null)
                    continue;

                // Compare bounding boxes with each wall
                foreach (Element w in wallCollector)
                {
                    BoundingBoxXYZ wallBb = w.get_BoundingBox(null);
                    if (wallBb == null)
                        continue;

                    // Check if bounding boxes intersect
                    if (BoundingBoxesIntersect(selBb, wallBb))
                    {
                        if (!intersectingWallIds.Contains(w.Id))
                        {
                            intersectingWallIds.Add(w.Id);
                        }
                    }
                }
            }

            if (intersectingWallIds.Count == 0)
            {
                TaskDialog.Show("Cut Walls", "No intersecting walls found.");
                return Result.Succeeded;
            }

            // 2. Perform the "cut" using SolidSolidCutUtils in a transaction
            using (Transaction tx = new Transaction(doc, "Cut Walls With Selection"))
            {
                tx.Start();

                // Attempt to cut each intersecting wall with each selected element
                foreach (ElementId wallId in intersectingWallIds)
                {
                    Element wall = doc.GetElement(wallId);
                    
                    foreach (ElementId selId in selectedIds)
                    {
                        Element cutter = doc.GetElement(selId);

                        try
                        {
                            // SolidSolidCutUtils: first param is the element to be cut (the wall),
                            // second param is the element that does the cutting (the selected element).
                            SolidSolidCutUtils.AddCutBetweenSolids(doc, wall, cutter);
                        }
                        catch (Exception ex)
                        {
                            // If the cut fails, handle/log or ignore
                            // e.g. "Selected element not a valid cutter" or geometry issues
                            // TaskDialog.Show("Cut Error", ex.Message);
                        }
                    }
                }

                tx.Commit();
            }

            // Optionally, select the walls to show which were processed
            uiDoc.SetSelectionIds(intersectingWallIds);

            return Result.Succeeded;
        }

        /// <summary>
        /// Simple axis-aligned bounding-box intersection test.
        /// </summary>
        private bool BoundingBoxesIntersect(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
        {
            if (bb1.Max.X < bb2.Min.X || bb1.Min.X > bb2.Max.X) return false;
            if (bb1.Max.Y < bb2.Min.Y || bb1.Min.Y > bb2.Max.Y) return false;
            if (bb1.Max.Z < bb2.Min.Z || bb1.Min.Z > bb2.Max.Z) return false;

            return true;
        }
    }
}
