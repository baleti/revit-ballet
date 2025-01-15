using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace RevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class SelectElementsHostedBySelectedTags : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Get the currently selected elements (could be tags, or possibly other categories)
            ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

            if (!selectedIds.Any())
            {
                message = "No tags are selected. Please select one or more tags first.";
                return Result.Failed;
            }

            // We'll track the actual tags out of the selection:
            List<IndependentTag> selectedTags = new List<IndependentTag>();

            // Filter the selection to ensure we're only dealing with IndependentTag elements
            foreach (ElementId elemId in selectedIds)
            {
                Element elem = doc.GetElement(elemId);
                if (elem is IndependentTag tag)
                {
                    selectedTags.Add(tag);
                }
            }

            if (!selectedTags.Any())
            {
                message = "The current selection contains no tags. Please select one or more tags.";
                return Result.Failed;
            }

            // We'll gather the IDs of all local (non-linked) elements that these tags reference
            HashSet<ElementId> elementsToSelect = new HashSet<ElementId>();

            // Loop through each selected tag to find what elements it references
            foreach (IndependentTag tag in selectedTags)
            {
                // GetTaggedElementIds() returns a collection of LinkElementId, 
                // which might represent either local or linked elements
                ICollection<LinkElementId> linkElementIds = tag.GetTaggedElementIds();

                // For each LinkElementId, check if it's local or linked
                foreach (LinkElementId linkElementId in linkElementIds)
                {
                    // If 'LinkInstanceId' is invalid, it's a local element
                    if (linkElementId.LinkInstanceId == ElementId.InvalidElementId)
                    {
                        // Add the local HostElementId to our selection set
                        elementsToSelect.Add(linkElementId.HostElementId);
                    }
                    else
                    {
                        // This indicates the element is in a linked document
                        // If you need to handle that, you'd retrieve the linked doc 
                        // and locate the corresponding element there.
                        // For now, we'll ignore linked elements in this example.
                    }
                }
            }

            // If we found any local elements to select, update the UI selection
            if (elementsToSelect.Any())
            {
                uiDoc.Selection.SetElementIds(elementsToSelect);
            }
            else
            {
                message = "No local elements were found that are referenced by the selected tags.";
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }
}
