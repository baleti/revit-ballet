using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MyRevitCommands
{
    /// <summary>
    /// External command that adds every revision cloud in the entire project (including those whose
    /// corresponding revision is set to "Show: None" in the Sheets Issues/Revision dialog) to the
    /// current user selection.  Target framework: .NET 4.8; Language version: C# 7.3; Revit API 2024.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class SelectAllRevisionCloudsInProject : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (commandData == null) throw new ArgumentNullException(nameof(commandData));

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc      = uidoc.Document;

            try
            {
                // 1. Collect every RevisionCloud element from the whole model (views + sheets).
                IList<ElementId> cloudIds = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(RevisionCloud))
                    .Select(e => e.Id)
                    .ToList();

                if (!cloudIds.Any())
                {
                    TaskDialog.Show("Select Revision Clouds", "No revision clouds were found in this project.");
                    return Result.Succeeded;
                }

                // 2. Merge with whatever the user currently has selected.
                ICollection<ElementId> currentSelection = uidoc.Selection.GetElementIds();
                HashSet<ElementId> mergedSelection       = new HashSet<ElementId>(currentSelection);
                mergedSelection.UnionWith(cloudIds);

                // 3. Apply the new selection set.
                uidoc.Selection.SetElementIds(mergedSelection.ToList());

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
