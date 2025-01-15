using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace RevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class SelectAssociatedTags : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Get the currently selected elements
            ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

            // If no elements are selected, warn the user
            if (!selectedIds.Any())
            {
                message = "No elements selected. Please select one or more elements first.";
                return Result.Failed;
            }

            // Collect all IndependentTag elements from the document
            FilteredElementCollector tagCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(IndependentTag));

            List<ElementId> tagsToSelect = new List<ElementId>();

            // For each tag, check whether it references any of the selected elements in the *current* document
            foreach (IndependentTag tag in tagCollector)
            {
                // Get the LinkElementId objects that this tag references
                ICollection<LinkElementId> linkElementIds = tag.GetTaggedElementIds();
                
                // Check if any of these LinkElementIds match the local selected IDs
                bool tagReferencesSelection = linkElementIds.Any(linkElementId =>
                {
                    // If LinkInstanceId is invalid, the element is local
                    if (linkElementId.LinkInstanceId == ElementId.InvalidElementId)
                    {
                        ElementId localId = linkElementId.HostElementId;
                        return selectedIds.Contains(localId);
                    }
                    else
                    {
                        // If you need to handle linked elements, add logic here 
                        // to match the linked document element Id if desired.
                        return false;
                    }
                });

                if (tagReferencesSelection)
                {
                    tagsToSelect.Add(tag.Id);
                }
            }

            // Update the selection to include just the associated tags
            uiDoc.Selection.SetElementIds(tagsToSelect);

            return Result.Succeeded;
        }
    }
}
