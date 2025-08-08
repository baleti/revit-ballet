using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectFamilyTypesOfSelectedElements : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            var uiApp = commandData.Application;
            var uiDoc = uiApp.ActiveUIDocument;
            var doc = uiDoc.Document;

            // 1) Prefer references (to capture linked selections). If none, fall back to element ids.
            IList<Reference> selectedRefs = uiDoc.GetReferences();
            var resultRefs = new List<Reference>();

            if (selectedRefs != null && selectedRefs.Count > 0)
            {
                foreach (var r in selectedRefs)
                {
                    TryAddTypeReferenceFromReference(uiDoc, r, resultRefs);
                }
            }
            else
            {
                // Fallback for plain element-id selection (host doc only)
                var ids = uiDoc.GetSelectionIds();
                foreach (var id in ids)
                {
                    var e = doc.GetElement(id);
                    TryAddTypeReferenceFromHostElement(e, resultRefs);
                }
            }

            // If we found nothing, bail politely
            if (resultRefs.Count == 0)
            {
                message = "No family instances or groups found in the selection.";
                return Result.Failed;
            }

            // 2) Set selection using references so we can include both host + linked type elements
            uiDoc.SetReferences(resultRefs);

            return Result.Succeeded;
        }
        catch (OperationCanceledException)
        {
            return Result.Cancelled;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }

    /// <summary>
    /// Given a selection Reference (which may be linked), resolve the underlying element,
    /// find its type element, and add a reference to that type element (host or linked) to resultRefs.
    /// </summary>
    private static void TryAddTypeReferenceFromReference(UIDocument uiDoc, Reference reference, List<Reference> resultRefs)
    {
        var hostDoc = uiDoc.Document;

        if (reference == null)
            return;

        // Linked case: LinkedElementId is valid, ElementId is the RevitLinkInstance in the host doc
        bool isLinked = reference.LinkedElementId != ElementId.InvalidElementId;

        if (isLinked)
        {
            var linkInst = hostDoc.GetElement(reference.ElementId) as RevitLinkInstance;
            if (linkInst == null)
                return;

            var linkedDoc = linkInst.GetLinkDocument();
            if (linkedDoc == null)
                return;

            var linkedElem = linkedDoc.GetElement(reference.LinkedElementId);
            if (linkedElem == null)
                return;

            // Resolve the type element in the linked doc
            var linkedTypeElem = GetTypeElement(linkedElem);
            if (linkedTypeElem == null)
                return;

            // Create a reference to the linked type element and convert it to a link reference
            var typeRef = new Reference(linkedTypeElem);
            var linkRef = typeRef.CreateLinkReference(linkInst);
            if (linkRef != null)
                resultRefs.Add(linkRef);
        }
        else
        {
            // Host element
            var elem = hostDoc.GetElement(reference.ElementId);
            TryAddTypeReferenceFromHostElement(elem, resultRefs);
        }
    }

    /// <summary>
    /// For a host (active-doc) element, resolve its type element and add its Reference.
    /// </summary>
    private static void TryAddTypeReferenceFromHostElement(Element elem, List<Reference> resultRefs)
    {
        if (elem == null)
            return;

        var typeElem = GetTypeElement(elem);
        if (typeElem == null)
            return;

        resultRefs.Add(new Reference(typeElem));
    }

    /// <summary>
    /// Returns the type element for supported selections:
    /// - FamilyInstance -> FamilySymbol
    /// - Group (model/detail) -> GroupType
    /// - ElementType (already a type) -> itself
    /// Otherwise returns null.
    /// </summary>
    private static Element GetTypeElement(Element elem)
    {
        if (elem == null)
            return null;

        var doc = elem.Document;

        // If the user already selected a type element, just keep it
        if (elem is ElementType)
            return elem;

        // Family instance -> its symbol (type)
        if (elem is FamilyInstance fi)
            return fi.Symbol;

        // Groups (model/detail) -> GroupType
        if (elem is Group g)
            return doc.GetElement(g.GetTypeId()) as GroupType;

        // Generic fallback via GetTypeId (works for many instances that have a valid type)
        ElementId typeId = elem.GetTypeId();
        if (typeId != ElementId.InvalidElementId)
            return doc.GetElement(typeId) as ElementType;

        return null;
    }
}
