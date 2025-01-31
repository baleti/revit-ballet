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
                if (elem is FilledRegion)
                {
                    selectedRegions.Add(elem as FilledRegion);
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
                    // Get the view where the filled region is placed
                    View view = doc.GetElement(region.OwnerViewId) as View;

                    // Get filled region type
                    FilledRegionType regionType = doc.GetElement(region.GetTypeId()) as FilledRegionType;

                    // Get all boundary loops
                    IList<CurveLoop> loops = region.GetBoundaries();

                    // Skip if only one loop
                    if (loops.Count <= 1) continue;

                    // Create new filled region for each loop
                    foreach (CurveLoop loop in loops)
                    {
                        List<CurveLoop> singleLoop = new List<CurveLoop> { loop };
                        FilledRegion.Create(doc, region.GetTypeId(), region.OwnerViewId, singleLoop);
                    }

                    // Delete original filled region
                    doc.Delete(region.Id);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}
