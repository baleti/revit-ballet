using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class IsolateElementsInView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = doc.ActiveView;
        
        try
        {
            // Get selected elements - try references first for linked elements support
            IList<Reference> selectedReferences = uidoc.GetReferences();
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            
            // Check if we have any selection
            bool hasReferences = selectedReferences != null && selectedReferences.Count > 0;
            bool hasElementIds = selectedIds != null && selectedIds.Count > 0;
            
            if (!hasReferences && !hasElementIds)
            {
                TaskDialog.Show("No Selection", "Please select elements to isolate.");
                return Result.Cancelled;
            }
            
            // Collections to track what we want to keep visible
            HashSet<ElementId> elementsToKeepVisible = new HashSet<ElementId>();
            HashSet<ElementId> linkInstancesWithSelection = new HashSet<ElementId>();
            bool hasLinkedElements = false;
            
            // Process references (handles both regular and linked elements)
            if (hasReferences)
            {
                foreach (Reference reference in selectedReferences)
                {
                    if (reference.LinkedElementId != ElementId.InvalidElementId)
                    {
                        // This is a linked element
                        ElementId linkInstanceId = reference.ElementId;
                        
                        // Add the link instance to keep visible
                        elementsToKeepVisible.Add(linkInstanceId);
                        linkInstancesWithSelection.Add(linkInstanceId);
                        hasLinkedElements = true;
                    }
                    else
                    {
                        // Regular element in current document
                        elementsToKeepVisible.Add(reference.ElementId);
                    }
                }
            }
            
            // Also add any directly selected element IDs
            if (hasElementIds)
            {
                foreach (ElementId id in selectedIds)
                {
                    elementsToKeepVisible.Add(id);
                }
            }
            
            // Get all visible elements in the view
            FilteredElementCollector collector = new FilteredElementCollector(doc, activeView.Id)
                .WhereElementIsNotElementType();
            
            // Build list of elements to hide
            List<ElementId> elementsToHide = new List<ElementId>();
            
            foreach (ElementId id in collector.ToElementIds())
            {
                if (!elementsToKeepVisible.Contains(id))
                {
                    Element elem = doc.GetElement(id);
                    if (elem != null && elem.CanBeHidden(activeView))
                    {
                        elementsToHide.Add(id);
                    }
                }
            }
            
            // Start a transaction to hide elements
            using (Transaction trans = new Transaction(doc, "Isolate Selected Elements"))
            {
                trans.Start();
                
                // Hide elements in the current document
                if (elementsToHide.Count > 0)
                {
                    activeView.HideElements(elementsToHide);
                }
                
                trans.Commit();
            }
            
            // Report results
            string resultMessage;
            if (hasLinkedElements)
            {
                resultMessage = $"Isolated {elementsToKeepVisible.Count} element(s) in the view.\n\n" +
                               "Note: Individual elements within linked models cannot be hidden using the API. " +
                               "Only entire link instances were kept visible. To isolate specific linked elements, " +
                               "use Revit's UI isolation tools or consider using view filters.";
            }
            else
            {
                resultMessage = $"Isolated {elementsToKeepVisible.Count} element(s) by hiding {elementsToHide.Count} other element(s) in the view.";
            }
            
            TaskDialog.Show("Isolation Complete", resultMessage);
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}
