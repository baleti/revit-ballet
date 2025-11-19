using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class IsolateSelectedElementsToTemporaryHideIsolate : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;

        try
        {
            // No need to check if temporary hide/isolate is active
            // We'll create one if it doesn't exist

            // Get currently selected elements
            ICollection<ElementId> selectedElementIds = uidoc.GetSelectionIds();

            if (selectedElementIds.Count == 0)
            {
                TaskDialog.Show("Info", "No elements are currently selected.");
                return Result.Succeeded;
            }

            // Process selected elements and expand groups
            List<ElementId> elementsToIsolate = new List<ElementId>();
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
                            elementsToIsolate.Add(memberId);
                        }
                    }
                }
                else if (elem.CanBeHidden(activeView))
                {
                    // For non-group elements, add directly if they can be hidden
                    elementsToIsolate.Add(id);
                }
            }

            if (elementsToIsolate.Count == 0)
            {
                TaskDialog.Show("Info", "None of the selected elements can be hidden/unhidden in this view.");
                return Result.Succeeded;
            }

            using (Transaction trans = new Transaction(doc, "Isolate Selected Elements in Temporary Hide/Isolate"))
            {
                trans.Start();

                // If temporary hide/isolate is active, disable it first
                if (activeView.IsInTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate))
                {
                    activeView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                }
                
                // Isolate only the selected elements (and expanded group members)
                activeView.IsolateElementsTemporary(elementsToIsolate);

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
