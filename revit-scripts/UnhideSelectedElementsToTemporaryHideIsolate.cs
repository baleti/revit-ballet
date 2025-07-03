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

            // Filter selected elements to only those that can be hidden in the view
            List<ElementId> validSelectedIds = new List<ElementId>();
            foreach (ElementId id in selectedElementIds)
            {
                Element elem = doc.GetElement(id);
                if (elem != null && elem.CanBeHidden(activeView))
                {
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

                // Strategy: Collect ALL visible elements (regardless of category) and add the selected elements
                List<ElementId> elementsToIsolate = new List<ElementId>();
                
                // Collect ALL elements in the view that can be hidden
                FilteredElementCollector collector = new FilteredElementCollector(doc, activeView.Id)
                    .WhereElementIsNotElementType();

                // Add ALL elements that can be hidden and are not permanently hidden
                foreach (Element elem in collector)
                {
                    if (elem.CanBeHidden(activeView) && !elem.IsHidden(activeView))
                    {
                        elementsToIsolate.Add(elem.Id);
                    }
                }

                // Ensure selected elements are in the isolation list
                foreach (ElementId id in validSelectedIds)
                {
                    if (!elementsToIsolate.Contains(id))
                    {
                        elementsToIsolate.Add(id);
                    }
                }

                // Reset current temporary mode and apply new isolation
                activeView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                
                if (elementsToIsolate.Count > 0)
                {
                    activeView.IsolateElementsTemporary(elementsToIsolate);
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
