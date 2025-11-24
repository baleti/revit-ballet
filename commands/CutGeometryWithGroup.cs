using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class CutGeometryWithGroup : IExternalCommand
{
  public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
  {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get element to cut
        Reference pickedRef = uidoc.Selection.PickObject(ObjectType.Element, "Select an element to cut");
        Element selectedElement = doc.GetElement(pickedRef);
        if (selectedElement == null)
        {
          message = "Please select an element.";
          return Result.Failed;
        }

        // Get group
        Reference pickedGroupRef = uidoc.Selection.PickObject(ObjectType.Element, "Select a model group");
        if (pickedGroupRef == null)
        {
          message = "Please select a model group.";
          return Result.Failed;
        }
        Group pickedGroup = doc.GetElement(pickedGroupRef.ElementId) as Group;
#if REVIT2017
        // In Revit 2017, GetDependentElements doesn't exist - use GetMemberIds and filter manually
        IList<ElementId> allMemberIds = pickedGroup.GetMemberIds();
        List<ElementId> dependentIds = new List<ElementId>();
        foreach (ElementId id in allMemberIds)
        {
            Element elem = doc.GetElement(id);
            if (elem is FamilyInstance)
                dependentIds.Add(id);
        }
#else
        ElementClassFilter filter = new ElementClassFilter(typeof(FamilyInstance));
        IList<ElementId> dependentIds = pickedGroup.GetDependentElements(filter);
#endif

        // Use a transaction to group cutting operations
        Transaction tx = new Transaction(doc);
        tx.Start("Cut with Group");
        try
        {
            foreach (ElementId id in dependentIds)
            {
                Element depElem = doc.GetElement(id);
                SolidSolidCutUtils.AddCutBetweenSolids(doc, selectedElement, depElem);
            }

            tx.Commit();
            message = "Element cut successfully with group members.";
        }
        catch (System.Exception ex)
        {
          tx.RollBack();
          message = $"An error occurred: {ex.Message}";
          return Result.Failed;
        }

        return Result.Succeeded;
  }
}
