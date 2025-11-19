using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class UnhideSelectedElementsToTemporaryHideIsolate : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;

        try
        {
            // Check if temporary hide/isolate is active
            if (!activeView.IsInTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate))
            {
                TaskDialog.Show("Info", 
                    "No temporary hide/isolate mode is active.\n\n" +
                    "This command adds selected elements to an existing temporary isolation.");
                return Result.Succeeded;
            }

            // Get currently selected elements
            ICollection<ElementId> selectedElementIds = uidoc.GetSelectionIds();

            if (selectedElementIds.Count == 0)
            {
                TaskDialog.Show("Info", "No elements are currently selected.");
                return Result.Succeeded;
            }

            // Process selected elements and expand groups
            List<ElementId> validSelectedIds = new List<ElementId>();
            foreach (ElementId id in selectedElementIds)
            {
                Element elem = doc.GetElement(id);
                if (elem == null) continue;

                // Check if this is a model group
                if (elem is Group group)
                {
                    // For groups, add all member elements that can be hidden
                    ICollection<ElementId> memberIds = group.GetMemberIds();
                    foreach (ElementId memberId in memberIds)
                    {
                        Element memberElem = doc.GetElement(memberId);
                        if (memberElem != null && memberElem.CanBeHidden(activeView))
                        {
                            validSelectedIds.Add(memberId);
                        }
                    }
                }
                else if (elem.CanBeHidden(activeView))
                {
                    // For non-group elements, add directly if they can be hidden
                    validSelectedIds.Add(id);
                }
            }

            if (validSelectedIds.Count == 0)
            {
                TaskDialog.Show("Info", "None of the selected elements can be hidden/unhidden in this view.");
                return Result.Succeeded;
            }

            using (Transaction trans = new Transaction(doc, "Add Selected Elements to Temporary Hide/Isolate"))
            {
                trans.Start();

                // Strategy: Collect currently visible elements and add selected elements to them
                HashSet<ElementId> currentlyVisibleIds = new HashSet<ElementId>();
                
                // Important: Use document-wide collector to avoid section box filtering
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType();

                // Check each element's visibility in the current view
                foreach (Element elem in collector)
                {
                    ElementId elemId = elem.Id;
                    
                    // Skip if element cannot be hidden in this view
                    if (!elem.CanBeHidden(activeView))
                        continue;
                    
                    // Check if element is visible (not permanently hidden and not temporarily hidden)
                    bool isPermanentlyHidden = elem.IsHidden(activeView);
                    bool isTemporarilyHidden = activeView.IsElementVisibleInTemporaryViewMode(
                        TemporaryViewMode.TemporaryHideIsolate, elemId) == false;
                    
                    // If element is currently visible (not hidden by either method)
                    if (!isPermanentlyHidden && !isTemporarilyHidden)
                    {
                        currentlyVisibleIds.Add(elemId);
                    }
                }

                // Add selected elements (including expanded group members) to the visible set
                foreach (ElementId id in validSelectedIds)
                {
                    currentlyVisibleIds.Add(id);
                }

                // Reset current temporary mode and apply new isolation
                activeView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                
                if (currentlyVisibleIds.Count > 0)
                {
                    activeView.IsolateElementsTemporary(currentlyVisibleIds.ToList());
                }

                trans.Commit();
            }

            uidoc.RefreshActiveView();
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}
