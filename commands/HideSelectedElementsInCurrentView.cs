using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class HideSelectedElementsInCurrentView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;

        try
        {
            // Check if the active view supports element hiding
            if (activeView is ViewSheet || activeView is ViewSchedule)
            {
                TaskDialog.Show("Info", 
                    "Cannot hide elements in this view type.\n" +
                    "Element hiding is not supported in sheets or schedules.");
                return Result.Cancelled;
            }

            // Get currently selected elements using SelectionModeManager
            ICollection<ElementId> selectedElementIds = uidoc.GetSelectionIds();

            if (selectedElementIds.Count == 0)
            {
                TaskDialog.Show("Info", "No elements are currently selected.");
                return Result.Succeeded;
            }

            // Process selected elements and expand groups
            List<ElementId> elementsToHide = new List<ElementId>();
            List<string> warningMessages = new List<string>();
            int groupsExpanded = 0;

            foreach (ElementId id in selectedElementIds)
            {
                Element elem = doc.GetElement(id);
                if (elem == null) continue;

                // Check if this is a model group
                if (elem is Group group)
                {
                    groupsExpanded++;
                    
                    // For groups, add all member elements that can be hidden
                    ICollection<ElementId> memberIds = group.GetMemberIds();
                    foreach (ElementId memberId in memberIds)
                    {
                        Element memberElem = doc.GetElement(memberId);
                        if (memberElem != null && memberElem.CanBeHidden(activeView))
                        {
                            // Check if element is already hidden
                            if (!memberElem.IsHidden(activeView))
                            {
                                elementsToHide.Add(memberId);
                            }
                        }
                    }
                }
                else if (elem.CanBeHidden(activeView))
                {
                    // For non-group elements, add directly if they can be hidden
                    // Check if element is already hidden
                    if (!elem.IsHidden(activeView))
                    {
                        elementsToHide.Add(id);
                    }
                    else
                    {
                        // Element is already hidden
                        var category = elem.Category?.Name ?? "Unknown";
                        warningMessages.Add($"• {category}: {elem.Name ?? elem.Id.ToString()} (already hidden)");
                    }
                }
                else
                {
                    // Element cannot be hidden in this view
                    var category = elem.Category?.Name ?? "Unknown";
                    warningMessages.Add($"• {category}: {elem.Name ?? elem.Id.ToString()} (cannot be hidden in this view)");
                }
            }

            if (elementsToHide.Count == 0)
            {
                string infoMessage = "No elements to hide.\n\n";
                
                if (warningMessages.Count > 0)
                {
                    infoMessage += "Details:\n";
                    // Limit the number of warning messages shown
                    int maxWarnings = 10;
                    for (int i = 0; i < Math.Min(warningMessages.Count, maxWarnings); i++)
                    {
                        infoMessage += warningMessages[i] + "\n";
                    }
                    if (warningMessages.Count > maxWarnings)
                    {
                        infoMessage += $"... and {warningMessages.Count - maxWarnings} more items";
                    }
                }
                else
                {
                    infoMessage += "All selected elements are either already hidden or cannot be hidden in this view.";
                }
                
                TaskDialog.Show("Info", infoMessage);
                return Result.Succeeded;
            }

            using (Transaction trans = new Transaction(doc, "Hide Selected Elements in View"))
            {
                trans.Start();

                try
                {
                    // Hide the elements permanently in the current view
                    activeView.HideElements(elementsToHide);
                    
                    trans.Commit();
                    
                    // Build success message
                    string successMessage = $"Successfully hid {elementsToHide.Count} element(s) in {activeView.Name}.";
                    
                    if (groupsExpanded > 0)
                    {
                        successMessage += $"\n\nNote: {groupsExpanded} group(s) were expanded to hide their member elements.";
                    }
                    
                    if (warningMessages.Count > 0)
                    {
                        successMessage += $"\n\n{warningMessages.Count} element(s) were skipped (already hidden or cannot be hidden).";
                    }
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    throw new Exception($"Failed to hide elements: {ex.Message}", ex);
                }
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
