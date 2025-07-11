using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class CombineSelectedFilledRegions : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Selection sel = uidoc.Selection;

            // Get selected filled regions
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            List<FilledRegion> selectedRegions = new List<FilledRegion>();

            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is FilledRegion region)
                {
                    selectedRegions.Add(region);
                }
            }

            if (selectedRegions.Count < 2)
            {
                TaskDialog.Show("Error", "Please select at least two filled regions to combine.");
                return Result.Failed;
            }

            // Verify all regions are in the same view
            ElementId viewId = selectedRegions[0].OwnerViewId;
            if (!selectedRegions.All(r => r.OwnerViewId.IntegerValue == viewId.IntegerValue))
            {
                TaskDialog.Show("Error", "All selected filled regions must be in the same view.");
                return Result.Failed;
            }

            using (Transaction trans = new Transaction(doc, "Combine Filled Regions"))
            {
                trans.Start();

                // Get the type and boundary style from the first region
                FilledRegion firstRegion = selectedRegions[0];
                ElementId typeId = firstRegion.GetTypeId();
                GraphicsStyle boundaryStyle = GetBoundaryLineStyle(doc, firstRegion);

                // Collect all boundary loops from all selected regions
                List<CurveLoop> allLoops = new List<CurveLoop>();
                foreach (FilledRegion region in selectedRegions)
                {
                    IList<CurveLoop> loops = region.GetBoundaries();
                    allLoops.AddRange(loops);
                }

                // Create a new filled region with all loops combined
                FilledRegion combinedRegion = FilledRegion.Create(
                    doc,
                    typeId,
                    viewId,
                    allLoops
                );

                // Apply the boundary style from the first region to the new combined region
                if (boundaryStyle != null && combinedRegion != null)
                {
                    SetBoundaryLineStyle(doc, combinedRegion, boundaryStyle);
                }

                // Delete all original filled regions
                foreach (FilledRegion region in selectedRegions)
                {
                    doc.Delete(region.Id);
                }

                // Select the newly created combined region
                sel.SetElementIds(new List<ElementId> { combinedRegion.Id });

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
