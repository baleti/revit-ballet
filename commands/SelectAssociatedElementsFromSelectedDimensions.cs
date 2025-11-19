using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectAssociatedElementsFromSelectedDimensions : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the active document and selection.
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;
        Selection selection = uiDoc.Selection;
        IList<ElementId> currentSelectionIds = uiDoc.GetSelectionIds().ToList();

        if (currentSelectionIds.Count == 0)
        {
            TaskDialog.Show("Error", "Please select one or more dimensions.");
            return Result.Cancelled;
        }

        // Use a HashSet to collect associated element IDs (avoiding duplicates)
        HashSet<ElementId> associatedElementIds = new HashSet<ElementId>();

        // Process each selected element and filter for dimensions
        foreach (ElementId id in currentSelectionIds)
        {
            Element elem = doc.GetElement(id);
            Dimension dim = elem as Dimension;
            if (dim != null)
            {
                // Loop through each reference in the dimension
                ReferenceArray refs = dim.References;
                if (refs != null)
                {
                    foreach (Reference r in refs)
                    {
                        associatedElementIds.Add(r.ElementId);
                    }
                }
            }
        }

        // Create a union of the original selection and the associated elements
        HashSet<ElementId> newSelection = new HashSet<ElementId>(currentSelectionIds);
        newSelection.UnionWith(associatedElementIds);

        if (newSelection.Count > 0)
        {
            // Update the selection to include both dimensions and their associated elements
            uiDoc.SetSelectionIds(newSelection.ToList());
        }
        else
        {
            TaskDialog.Show("Result", "No associated elements found for the selected dimensions.");
        }

        return Result.Succeeded;
    }
}
