using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    [CommandMeta("Filled Region")]
    public class SplitFilledRegions : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Selection sel = uidoc.Selection;

            // Get selected filled regions
            List<FilledRegion> selectedRegions = new List<FilledRegion>();
            foreach (ElementId id in uidoc.GetSelectionIds())
            {
                if (doc.GetElement(id) is FilledRegion region)
                    selectedRegions.Add(region);
            }

            if (selectedRegions.Count == 0)
            {
                CustomGUIs.SetCurrentUIDocument(uidoc);
                var allRegions = new FilteredElementCollector(doc)
                    .OfClass(typeof(FilledRegion)).Cast<FilledRegion>().ToList();
                var gridData = CustomGUIs.ConvertToDataGridFormat(allRegions, new List<string> { "Name" });
                var chosen = CustomGUIs.DataGrid(gridData, new List<string> { "Name" }, false);
                if (chosen == null) return Result.Failed;
                selectedRegions = CustomGUIs.ExtractOriginalObjects<FilledRegion>(chosen) ?? new List<FilledRegion>();
                if (selectedRegions.Count == 0) return Result.Succeeded;
            }

            List<ElementId> newRegionIds = new List<ElementId>();

            using (Transaction trans = new Transaction(doc, "Split Filled Regions"))
            {
                trans.Start();

                foreach (FilledRegion region in selectedRegions)
                {
                    // Get all boundary loops
                    IList<CurveLoop> loops = region.GetBoundaries();

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

                        newRegionIds.Add(newRegion.Id);
                    }

                    // Delete the original filled region
                    doc.Delete(region.Id);
                }

                trans.Commit();
            }

            // Set the new regions as the current selection
            uidoc.SetSelectionIds(newRegionIds);

            return Result.Succeeded;
        }

        /// <summary>
        /// Retrieves the line style (GraphicsStyle) of the first boundary line
        /// used by the given FilledRegion, by examining its dependent elements.
        /// </summary>
        private GraphicsStyle GetBoundaryLineStyle(Document doc, FilledRegion region)
        {
#if REVIT2017
            // GetDependentElements not available in Revit 2017 - skip line style preservation
            return null;
#else
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
#endif
        }

        /// <summary>
        /// Sets the boundary line style (GraphicsStyle) for all boundary lines
        /// of the newly created FilledRegion (again using its dependent elements).
        /// </summary>
        private void SetBoundaryLineStyle(Document doc, FilledRegion newRegion, GraphicsStyle style)
        {
#if REVIT2017
            // GetDependentElements not available in Revit 2017 - skip line style setting
            return;
#else
            var dependentIds = newRegion.GetDependentElements(null);
            foreach (ElementId depId in dependentIds)
            {
                Element e = doc.GetElement(depId);
                if (e is CurveElement curveElem)
                {
                    curveElem.LineStyle = style;
                }
            }
#endif
        }
    }
}
