using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

[Transaction(TransactionMode.Manual)]
public class SelectAssociatedTagsOfSelectedElements : IExternalCommand
{
    public Result Execute(ExternalCommandData cData, ref string msg, ElementSet els)
    {
        UIDocument  uiDoc = cData.Application.ActiveUIDocument;
        Document       doc = uiDoc.Document;
        ICollection<ElementId> picked = uiDoc.GetSelectionIds();

        if (!picked.Any())
        {
            msg = "No elements selected.";
            return Result.Failed;
        }

        /* 1. Collect ALL tag types we care about in one shot */
        var tagFilter = new LogicalOrFilter(
            new ElementClassFilter(typeof(IndependentTag)),
            new ElementClassFilter(typeof(SpatialElementTag)));

        List<ElementId> tagIds = new FilteredElementCollector(doc)
                                 .WhereElementIsNotElementType()
                                 .WherePasses(tagFilter)
                                 .Where(t => TagReferencesSelection(t, picked))
                                 .Select(t => t.Id)
                                 .ToList();

        /* 2. Push the result back to Revit */
        if (tagIds.Count == 0)
        {
            TaskDialog.Show("No Tags Found", "No tags were found for the selected elements.");
            return Result.Succeeded;
        }

        uiDoc.SetSelectionIds(tagIds);
        return Result.Succeeded;
    }

    /// <summary>Does this tag point to *any* of the user-selected element IDs?</summary>
    private static bool TagReferencesSelection(Element tag, ICollection<ElementId> picked)
    {
        switch (tag)
        {
            /* ---------- IndependentTag ---------- */
            case IndependentTag it:
                return it.GetTaggedElementIds()
                         .Any(leId => leId.LinkInstanceId == ElementId.InvalidElementId &&
                                      picked.Contains(leId.HostElementId));

            /* ---------- Room & Area tags ---------- */
            case RoomTag rt when rt.Room != null:
                return picked.Contains(rt.Room.Id);

            case AreaTag at when at.Area != null:
                return picked.Contains(at.Area.Id);

            /* ---------- SpaceTag (MEP) ---------- */
            default:
                // Avoid hard reference to MEP assembly – reflection once, cheap.
                PropertyInfo spaceProp = tag.GetType().GetProperty("Space");
                if (spaceProp?.GetValue(tag) is Element sp)
                    return picked.Contains(sp.Id);

                return false;
        }
    }
}
