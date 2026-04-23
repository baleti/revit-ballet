using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class IsolateElementsInViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        try
        {
            // Determine target views (selected views or active view)
            ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();
            List<View> targetViews = new List<View>();

            bool hasSelectedViews = currentSelection.Any(id =>
            {
                Element elem = doc.GetElement(id);
                return elem is View;
            });

            if (hasSelectedViews)
            {
                // Use selected views
                targetViews = currentSelection
                    .Select(id => doc.GetElement(id))
                    .OfType<View>()
                    .ToList();
            }
            else
            {
                // Use current view only
                targetViews.Add(doc.ActiveView);
            }

            // Get selected elements - try references first for linked elements support
            IList<Reference> selectedReferences = uidoc.GetReferences();
            ICollection<ElementId> selectedIds = currentSelection;
            
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
            
            // Also add any directly selected element IDs (excluding views)
            if (hasElementIds)
            {
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (!(elem is View))
                    {
                        elementsToKeepVisible.Add(id);
                    }
                }
            }

            // Process each target view
            int totalHiddenCount = 0;
            List<string> viewsProcessed = new List<string>();

            using (Transaction trans = new Transaction(doc, "Isolate Selected Elements in Views"))
            {
                trans.Start();

                foreach (View view in targetViews)
                {
                    // Get all visible elements in this view
                    FilteredElementCollector collector = new FilteredElementCollector(doc, view.Id)
                        .WhereElementIsNotElementType();

                    // Build list of elements to hide in this view
                    List<ElementId> elementsToHide = new List<ElementId>();

                    foreach (ElementId id in collector.ToElementIds())
                    {
                        if (!elementsToKeepVisible.Contains(id))
                        {
                            Element elem = doc.GetElement(id);
                            if (elem != null && elem.CanBeHidden(view))
                            {
                                elementsToHide.Add(id);
                            }
                        }
                    }

                    // Hide elements in this view
                    if (elementsToHide.Count > 0)
                    {
                        view.HideElements(elementsToHide);
                        totalHiddenCount += elementsToHide.Count;
                    }

                    viewsProcessed.Add(view.Name);
                }

                trans.Commit();
            }

            // Report results
            string viewText = hasSelectedViews
                ? $"{targetViews.Count} view(s): {string.Join(", ", viewsProcessed)}"
                : "the active view";

            string resultMessage;
            if (hasLinkedElements)
            {
                resultMessage = $"Isolated {elementsToKeepVisible.Count} element(s) in {viewText}.\n\n" +
                               "Note: Individual elements within linked models cannot be hidden using the API. " +
                               "Only entire link instances were kept visible. To isolate specific linked elements, " +
                               "use Revit's UI isolation tools or consider using view filters.";
            }
            else
            {
                resultMessage = $"Isolated {elementsToKeepVisible.Count} element(s) by hiding {totalHiddenCount} other element(s) in {viewText}.";
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
