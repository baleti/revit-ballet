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
    public class SelectArrayMembersOfSelectedElements : IExternalCommand
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
                    TaskDialog.Show("Select Array Members", "No elements selected. Please select array elements or array members.");
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

                // Collect all member IDs from arrays related to selection
                var allMemberIds = new HashSet<ElementId>();
                var arrayCount = 0;

                foreach (var elemId in selection)
                {
                    Element elem = doc.GetElement(elemId);
                    if (elem == null) continue;

                    // If element is an array, get its members directly
                    if (elem is LinearArray linearArray)
                    {
                        allMemberIds.UnionWith(linearArray.GetOriginalMemberIds());
                        allMemberIds.UnionWith(linearArray.GetCopiedMemberIds());
                        arrayCount++;
                        continue;
                    }

                    if (elem is RadialArray radialArray)
                    {
                        allMemberIds.UnionWith(radialArray.GetOriginalMemberIds());
                        allMemberIds.UnionWith(radialArray.GetCopiedMemberIds());
                        arrayCount++;
                        continue;
                    }

                    // Otherwise, search for arrays containing this element
                    bool foundInArray = false;

                    // Search linear arrays
                    foreach (var array in allLinearArrays)
                    {
                        var originalMembers = array.GetOriginalMemberIds();
                        var copiedMembers = array.GetCopiedMemberIds();

                        if (originalMembers.Contains(elemId) || copiedMembers.Contains(elemId))
                        {
                            allMemberIds.UnionWith(originalMembers);
                            allMemberIds.UnionWith(copiedMembers);
                            arrayCount++;
                            foundInArray = true;
                            break;
                        }
                    }

                    if (foundInArray) continue;

                    // Search radial arrays
                    foreach (var array in allRadialArrays)
                    {
                        var originalMembers = array.GetOriginalMemberIds();
                        var copiedMembers = array.GetCopiedMemberIds();

                        if (originalMembers.Contains(elemId) || copiedMembers.Contains(elemId))
                        {
                            allMemberIds.UnionWith(originalMembers);
                            allMemberIds.UnionWith(copiedMembers);
                            arrayCount++;
                            foundInArray = true;
                            break;
                        }
                    }
                }

                if (allMemberIds.Count == 0)
                {
                    TaskDialog.Show("Select Array Members", "None of the selected elements are arrays or array members.");
                    return Result.Cancelled;
                }

                // Add to existing selection using SelectionModeManager
                uidoc.AddToSelection(allMemberIds);

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
