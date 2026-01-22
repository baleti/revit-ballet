using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class HideSelectedElementsInViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;

        try
        {
            // Get currently selected elements using SelectionModeManager
            ICollection<ElementId> selectedElementIds = uidoc.GetSelectionIds();

            // If no elements are selected, do nothing
            if (selectedElementIds.Count == 0)
            {
                return Result.Succeeded;
            }

            // Separate selected elements into views/viewports and regular elements
            List<View> targetViews = new List<View>();
            List<ElementId> elementsToHide = new List<ElementId>();

            foreach (ElementId id in selectedElementIds)
            {
                Element elem = doc.GetElement(id);
                if (elem == null) continue;

                if (elem is View view)
                {
                    // Don't hide in sheets or schedules
                    if (!(view is ViewSheet || view is ViewSchedule))
                    {
                        targetViews.Add(view);
                    }
                }
                else if (elem is Viewport viewport)
                {
                    // Get the view from the viewport
                    View viewFromViewport = doc.GetElement(viewport.ViewId) as View;
                    if (viewFromViewport != null &&
                        !(viewFromViewport is ViewSheet || viewFromViewport is ViewSchedule))
                    {
                        targetViews.Add(viewFromViewport);
                    }
                }
                else
                {
                    // Regular element - add to elements to hide
                    // Expand groups
                    if (elem is Group group)
                    {
                        ICollection<ElementId> memberIds = group.GetMemberIds();
                        foreach (ElementId memberId in memberIds)
                        {
                            Element memberElem = doc.GetElement(memberId);
                            if (memberElem != null)
                            {
                                elementsToHide.Add(memberId);
                            }
                        }
                    }
                    else
                    {
                        elementsToHide.Add(id);
                    }
                }
            }

            // If no views/viewports selected, use current active view
            if (targetViews.Count == 0)
            {
                // Check if active view supports element hiding
                if (activeView is ViewSheet || activeView is ViewSchedule)
                {
                    TaskDialog.Show("Error",
                        "Cannot hide elements in this view type.\n" +
                        "Element hiding is not supported in sheets or schedules.");
                    return Result.Failed;
                }
                targetViews.Add(activeView);
            }

            // If no elements to hide, do nothing
            if (elementsToHide.Count == 0)
            {
                return Result.Succeeded;
            }

            // Hide elements in each target view
            using (Transaction trans = new Transaction(doc, "Hide Selected Elements in Views"))
            {
                trans.Start();

                foreach (View targetView in targetViews)
                {
                    // Filter elements that can be hidden in this view
                    List<ElementId> validElementsToHide = new List<ElementId>();
                    foreach (ElementId id in elementsToHide)
                    {
                        Element elem = doc.GetElement(id);
                        if (elem != null && elem.CanBeHidden(targetView))
                        {
                            validElementsToHide.Add(id);
                        }
                    }

                    if (validElementsToHide.Count == 0)
                        continue;

                    // Get all currently visible elements in this view
                    HashSet<ElementId> currentlyVisibleIds = new HashSet<ElementId>();

                    FilteredElementCollector collector = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType();

                    foreach (Element elem in collector)
                    {
                        ElementId elemId = elem.Id;

                        // Skip if element cannot be hidden in this view
                        if (!elem.CanBeHidden(targetView))
                            continue;

                        // Check if element is visible (not permanently hidden and not temporarily hidden)
                        bool isPermanentlyHidden = elem.IsHidden(targetView);

                        bool isTemporarilyHidden = false;
                        if (targetView.IsInTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate))
                        {
                            isTemporarilyHidden = targetView.IsElementVisibleInTemporaryViewMode(
                                TemporaryViewMode.TemporaryHideIsolate, elemId) == false;
                        }

                        // If element is currently visible (not hidden by either method)
                        if (!isPermanentlyHidden && !isTemporarilyHidden)
                        {
                            currentlyVisibleIds.Add(elemId);
                        }
                    }

                    // Remove elements to hide from the visible set
                    foreach (ElementId id in validElementsToHide)
                    {
                        currentlyVisibleIds.Remove(id);
                    }

                    // Disable any existing temporary mode first
                    if (targetView.IsInTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate))
                    {
                        targetView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                    }

                    // Apply temporary isolation (showing only non-hidden elements)
                    if (currentlyVisibleIds.Count > 0)
                    {
                        targetView.IsolateElementsTemporary(currentlyVisibleIds.ToList());
                    }
                }

                trans.Commit();
            }

            // Refresh the view to show changes
            uidoc.RefreshActiveView();

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            TaskDialog.Show("Error", $"An error occurred:\n{ex.Message}");
            return Result.Failed;
        }
    }
}
