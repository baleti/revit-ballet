using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitBallet.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SelectArrayOfSelectedElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Get current selection using SelectionModeManager
                ICollection<ElementId> selection = uidoc.GetSelectionIds();

                if (selection.Count == 0)
                {
                    TaskDialog.Show("Select Array", "No elements selected. Please select group members that are part of an array.");
                    return Result.Cancelled;
                }

                // Collect all arrays in the document
                var allLinearArrays = new FilteredElementCollector(doc)
                    .OfClass(typeof(LinearArray))
                    .Cast<LinearArray>()
                    .ToList();

                var allRadialArrays = new FilteredElementCollector(doc)
                    .OfClass(typeof(RadialArray))
                    .Cast<RadialArray>()
                    .ToList();

                // Find array elements for selected members
                var arrayIds = new HashSet<ElementId>();

                foreach (var elemId in selection)
                {
                    Element elem = doc.GetElement(elemId);
                    if (elem == null) continue;

                    // Check if the element itself is already an array
                    if (elem is LinearArray || elem is RadialArray)
                    {
                        arrayIds.Add(elemId);
                        continue;
                    }

                    // Search linear arrays
                    foreach (var array in allLinearArrays)
                    {
                        var originalMembers = array.GetOriginalMemberIds();
                        var copiedMembers = array.GetCopiedMemberIds();

                        if (originalMembers.Contains(elemId) || copiedMembers.Contains(elemId))
                        {
                            arrayIds.Add(array.Id);
                            break;
                        }
                    }

                    // Search radial arrays
                    foreach (var array in allRadialArrays)
                    {
                        var originalMembers = array.GetOriginalMemberIds();
                        var copiedMembers = array.GetCopiedMemberIds();

                        if (originalMembers.Contains(elemId) || copiedMembers.Contains(elemId))
                        {
                            arrayIds.Add(array.Id);
                            break;
                        }
                    }
                }

                if (arrayIds.Count == 0)
                {
                    TaskDialog.Show("Select Array", "None of the selected elements are part of an array.");
                    return Result.Cancelled;
                }

                // Set selection to array elements using SelectionModeManager
                uidoc.SetSelectionIds(arrayIds);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
