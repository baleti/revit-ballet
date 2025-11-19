using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Linq;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class HighlightDoubleSwitchedSocketOutletInLink : IExternalCommand, ISelectionFilter
{
    private RevitLinkInstance Instance = null;

    public bool AllowElement(Element elem)
    {
        // Check if the element is a RevitLinkInstance
        Instance = elem as RevitLinkInstance;
        if (Instance != null) return true;

        return false;
    }

    public bool AllowReference(Reference reference, XYZ position)
    {
        // Ensure we are working with a RevitLinkInstance
        if (Instance == null)
        {
            return false;
        }

        // Get the linked document
        Document linkDoc = Instance.GetLinkDocument();
        if (linkDoc == null) return false;

        // Get the element in the linked document
        Element linkedElement = linkDoc.GetElement(reference.LinkedElementId);

        // Filter by family name "Double Switched Socket Outlet"
        if (linkedElement is FamilyInstance familyInstance 
            && familyInstance.Symbol.Family.Name == "Double Switched Socket Outlet")
        {
            return true; // Allow highlighting
        }

        return false; // Ignore other elements
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
        // Get the current document
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Collect all the linked models
        FilteredElementCollector linkCollector = new FilteredElementCollector(doc)
            .OfClass(typeof(RevitLinkInstance));

        RevitLinkInstance targetLink = linkCollector
            .Cast<RevitLinkInstance>()
            .FirstOrDefault(link => link.Name.Contains("21630-WDA-A-ZZZ-M3-MEP-0001"));

        if (targetLink == null)
        {
            TaskDialog.Show("Error", "Could not find the linked model.");
            return Result.Failed;
        }

        // Select the linked RevitLinkInstance
        Instance = targetLink;

        // Ask the user to select elements in the linked model using the filter
        try
        {
            IList<Reference> selectedReferences = uiDoc.Selection.PickObjects(ObjectType.LinkedElement, this, "Select Double Switched Socket Outlet instances");

            if (selectedReferences.Count > 0)
            {
                TaskDialog.Show("Info", $"Selected {selectedReferences.Count} 'Double Switched Socket Outlet' instances.");
            }
            else
            {
                TaskDialog.Show("Info", "No instances were selected.");
            }

            return Result.Succeeded;
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
            return Result.Cancelled;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}
