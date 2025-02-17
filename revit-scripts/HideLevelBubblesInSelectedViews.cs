using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HideLevelBubbles
{
    [Transaction(TransactionMode.Manual)]
    public class HideLevelBubblesInSelectedViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the active UIDocument and Document.
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            if (uiDoc == null)
            {
                message = "No active document.";
                return Result.Failed;
            }
            Document doc = uiDoc.Document;

            // Retrieve the currently selected views.
            // (Ensure you select one or more views in the Project Browser before running this command.)
            ICollection<ElementId> selIds = uiDoc.Selection.GetElementIds();
            List<View> selectedViews = new List<View>();
            foreach (ElementId id in selIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is View view)
                {
                    selectedViews.Add(view);
                }
            }
            if (selectedViews.Count == 0)
            {
                message = "Please select one or more views.";
                return Result.Failed;
            }

            // Collect all Level elements.
            // Levels are datum planes that support HideBubbleInView().
            IList<Element> levels = new FilteredElementCollector(doc)
                                        .OfClass(typeof(Level))
                                        .ToElements();

            using (Transaction trans = new Transaction(doc, "Hide Level Bubbles"))
            {
                trans.Start();

                foreach (View view in selectedViews)
                {

                    foreach (Element levelElem in levels)
                    {
                        DatumPlane dp = levelElem as DatumPlane;
                        if (dp != null)
                        {
                            try
                            {
                                // Use valid DatumEnds (End1 and End2).
                                dp.HideBubbleInView(DatumEnds.End0, view);
                                dp.HideBubbleInView(DatumEnds.End1, view);
                            }
                            catch (Autodesk.Revit.Exceptions.ArgumentException)
                            {
                                // The datum plane isn't visible in this view.
                                // Log or ignore this exception as appropriate.
                            }
                        }
                    }
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
}
