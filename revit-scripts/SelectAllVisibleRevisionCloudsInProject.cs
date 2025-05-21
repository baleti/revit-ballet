using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MyRevitCommands
{
    /// <summary>
    /// Adds every *visible* revision cloud in the project to the current selection.
    /// A revision cloud is considered visible when its parent Revision’s
    /// Sheet Issues/Revisions “Show” setting is not “None”
    /// (i.e. RevisionVisibility.CloudAndTagVisible).
    /// Target framework: .NET 4.8 · C# 7.3 · Revit API 2024.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class SelectAllVisibleRevisionCloudsInProject : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            if (commandData == null)
                throw new ArgumentNullException(nameof(commandData));

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uidoc.Document;

            try
            {
                // 1. Collect only those clouds whose Revision is set to show clouds.
                IList<ElementId> visibleCloudIds = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(RevisionCloud))
                    .Cast<RevisionCloud>()
                    .Where(cloud =>
                    {
                        ElementId revId = cloud.RevisionId;
                        if (revId == ElementId.InvalidElementId) return false;

                        Revision rev = doc.GetElement(revId) as Revision;
                        return rev != null
                               && rev.Visibility == RevisionVisibility.CloudAndTagVisible;
                    })
                    .Select(c => c.Id)
                    .ToList();

                if (!visibleCloudIds.Any())
                {
                    TaskDialog.Show("Select Visible Revision Clouds",
                                    "No visible revision clouds were found in this project.");
                    return Result.Succeeded;
                }

                // 2. Merge with the user’s current selection.
                ICollection<ElementId> current = uidoc.Selection.GetElementIds();
                HashSet<ElementId> merged      = new HashSet<ElementId>(current);
                merged.UnionWith(visibleCloudIds);

                // 3. Set the new selection.
                uidoc.Selection.SetElementIds(merged.ToList());

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
