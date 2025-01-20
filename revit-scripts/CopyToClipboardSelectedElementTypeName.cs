using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms; // Needed for Clipboard.SetText

[Transaction(TransactionMode.Manual)]
public class CopyToClipboardTypeNameOfElementInLinkedModel : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document hostDoc = uiDoc.Document;
        Selection sel = uiDoc.Selection;

        try
        {
            // 1) Force user to pick an element from a linked model
            Reference pickedRef = sel.PickObject(ObjectType.LinkedElement, "Select an element from a linked model.");

            // 2) In the host doc, the 'ElementId' of pickedRef is the RevitLinkInstance
            RevitLinkInstance linkInstance = hostDoc.GetElement(pickedRef.ElementId) as RevitLinkInstance;
            if (linkInstance == null)
            {
                TaskDialog.Show("Error", "Reference did not point to a RevitLinkInstance.");
                return Result.Failed;
            }

            // 3) Get the linked document
            Document linkedDoc = linkInstance.GetLinkDocument();
            if (linkedDoc == null)
            {
                TaskDialog.Show("Error", "Could not retrieve the linked document.");
                return Result.Failed;
            }

            // 4) Retrieve the actual element inside the link
            Element linkedElement = linkedDoc.GetElement(pickedRef.LinkedElementId);
            if (linkedElement == null)
            {
                TaskDialog.Show("Error", "Could not retrieve the element in the linked document.");
                return Result.Failed;
            }

            // 5) Get the element's type
            ElementId typeId = linkedElement.GetTypeId();
            if (typeId == ElementId.InvalidElementId)
            {
                TaskDialog.Show("Error", "The selected linked element does not have a valid type.");
                return Result.Failed;
            }

            Element linkedType = linkedDoc.GetElement(typeId);
            if (linkedType == null)
            {
                TaskDialog.Show("Error", "Could not retrieve the type of the linked element.");
                return Result.Failed;
            }

            // 6) Copy the type name to clipboard
            string typeName = linkedType.Name ?? "<No Type Name>";
            Clipboard.SetText(typeName);
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
            // User canceled
            return Result.Cancelled;
        }
        catch (System.Exception ex)
        {
            TaskDialog.Show("Error", ex.Message);
            return Result.Failed;
        }

        return Result.Succeeded;
    }
}
