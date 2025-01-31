using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class SplitFilledRegions : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Selection sel = uidoc.Selection;

            // Get selected filled regions
            ICollection<ElementId> selectedIds = sel.GetElementIds();
            List<FilledRegion> selectedRegions = new List<FilledRegion>();

            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is FilledRegion region)
                {
                    selectedRegions.Add(region);
                }
            }

            if (selectedRegions.Count == 0)
            {
                TaskDialog.Show("Error", "Please select at least one filled region.");
                return Result.Failed;
            }

            using (Transaction trans = new Transaction(doc, "Split Filled Regions"))
            {
                trans.Start();

                foreach (FilledRegion region in selectedRegions)
                {
                    // Get all boundary loops
                    IList<CurveLoop> loops = region.GetBoundaries();

                    // Skip splitting if only one loop
                    if (loops.Count <= 1) continue;

                    // Capture the original boundary line style (subcategory).
                    GraphicsStyle boundaryStyle = GetBoundaryLineStyle(doc, region);

                    // Create new filled region for each loop
                    foreach (CurveLoop loop in loops)
                    {
                        List<CurveLoop> singleLoop = new List<CurveLoop> { loop };

                        // In recent Revit API versions, FilledRegion.Create(...) returns a FilledRegion directly
                        FilledRegion newRegion = FilledRegion.Create(
                            doc, 
                            region.GetTypeId(), 
                            region.OwnerViewId, 
                            singleLoop
                        );

                        // Re-apply the original boundary style to the newly created region's lines
                        if (boundaryStyle != null && newRegion != null)
                        {
                            SetBoundaryLineStyle(doc, newRegion, boundaryStyle);
                        }
                    }

                    // Delete the original filled region
                    doc.Delete(region.Id);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Retrieves the line style (GraphicsStyle) of the first boundary line
        /// used by the given FilledRegion, by examining its dependent elements.
        /// </summary>
        private GraphicsStyle GetBoundaryLineStyle(Document doc, FilledRegion region)
        {
            var dependentIds = region.GetDependentElements(null); // Gets any dependent elements, including boundary lines
            foreach (ElementId depId in dependentIds)
            {
                Element e = doc.GetElement(depId);
                if (e is CurveElement curveElem)
                {
                    // The 'LineStyle' property is actually a GraphicsStyle on CurveElement
                    return curveElem.LineStyle as GraphicsStyle;
                }
            }
            return null;
        }

        /// <summary>
        /// Sets the boundary line style (GraphicsStyle) for all boundary lines
        /// of the newly created FilledRegion (again using its dependent elements).
        /// </summary>
        private void SetBoundaryLineStyle(Document doc, FilledRegion newRegion, GraphicsStyle style)
        {
            var dependentIds = newRegion.GetDependentElements(null);
            foreach (ElementId depId in dependentIds)
            {
                Element e = doc.GetElement(depId);
                if (e is CurveElement curveElem)
                {
                    curveElem.LineStyle = style;
                }
            }
        }
    }
}
