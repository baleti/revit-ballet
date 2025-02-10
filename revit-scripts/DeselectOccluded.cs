using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class DeselectOccluded : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            try
            {
                // Check if view is 3D
                if (!(activeView is View3D))
                {
                    TaskDialog.Show("Error", "This command works only in 3D views.");
                    return Result.Failed;
                }

                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (!selectedIds.Any())
                {
                    TaskDialog.Show("Warning", "No elements selected.");
                    return Result.Succeeded;
                }

                ReferenceIntersector refIntersector = new ReferenceIntersector(activeView as View3D);
                refIntersector.FindReferencesInRevitLinks = false;
                refIntersector.TargetType = FindReferenceTarget.Element;

                List<ElementId> toDeselect = new List<ElementId>();

                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem == null) continue;

                    BoundingBoxXYZ bbox = elem.get_BoundingBox(activeView);
                    if (bbox == null) continue;

                    XYZ rayStart = new XYZ(
                        (bbox.Min.X + bbox.Max.X) / 2,
                        (bbox.Min.Y + bbox.Max.Y) / 2,
                        (bbox.Min.Z + bbox.Max.Z) / 2
                    );

                    XYZ viewDirection = activeView.ViewDirection;
                    XYZ rayDirection = -viewDirection;

                    IList<ReferenceWithContext> references = refIntersector.Find(rayStart, rayDirection);

                    bool isOccluded = false;
                    foreach (ReferenceWithContext refContext in references)
                    {
                        ElementId hitId = refContext.GetReference().ElementId;
                        if (hitId != id && selectedIds.Contains(hitId))
                        {
                            isOccluded = true;
                            break;
                        }
                    }

                    if (isOccluded)
                    {
                        toDeselect.Add(id);
                    }
                }

                if (toDeselect.Any())
                {
                    uidoc.Selection.SetElementIds(selectedIds.Except(toDeselect).ToList());
                }

                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
