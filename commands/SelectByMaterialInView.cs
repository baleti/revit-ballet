using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectByMaterialInView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Determine views to process
        ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();
        List<View> viewsToProcess = new List<View>();

        foreach (ElementId id in currentSelection)
        {
            Element elem = doc.GetElement(id);
            if (elem is View view)
                viewsToProcess.Add(view);
            else if (elem is Viewport vp)
            {
                View viewFromVp = doc.GetElement(vp.ViewId) as View;
                if (viewFromVp != null) viewsToProcess.Add(viewFromVp);
            }
        }

        if (viewsToProcess.Count == 0)
            viewsToProcess.Add(uidoc.ActiveView);

        // Build material -> element IDs mapping across all target views
        Dictionary<ElementId, List<ElementId>> materialToElements = new Dictionary<ElementId, List<ElementId>>();

        foreach (View view in viewsToProcess)
        {
            if (view == null) continue;
            var collector = new FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType();
            foreach (Element elem in collector)
                CollectMaterials(elem, materialToElements);
        }

        if (materialToElements.Count == 0)
        {
            TaskDialog.Show("No Materials", "No materials found in the target view(s).");
            return Result.Cancelled;
        }

        // Build DataGrid rows
        List<Dictionary<string, object>> materialList = BuildMaterialList(doc, materialToElements);

        CustomGUIs.SetCurrentUIDocument(uidoc);
        var propertyNames = new List<string> { "Name", "Material Class", "Count" };
        List<Dictionary<string, object>> selected = CustomGUIs.DataGrid(materialList, propertyNames, false);

        if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
            CustomGUIs.ApplyCellEditsToEntities();

        if (selected.Count == 0)
            return Result.Cancelled;

        List<ElementId> elementIds = GatherElementIds(selected);
        uidoc.SetSelectionIds(elementIds.Distinct().ToList());

        return Result.Succeeded;
    }

    private static void CollectMaterials(Element elem, Dictionary<ElementId, List<ElementId>> map)
    {
        try
        {
            foreach (ElementId matId in elem.GetMaterialIds(false))
            {
                if (!map.ContainsKey(matId)) map[matId] = new List<ElementId>();
                if (!map[matId].Contains(elem.Id)) map[matId].Add(elem.Id);
            }
            foreach (ElementId matId in elem.GetMaterialIds(true))
            {
                if (!map.ContainsKey(matId)) map[matId] = new List<ElementId>();
                if (!map[matId].Contains(elem.Id)) map[matId].Add(elem.Id);
            }
        }
        catch { }
    }

    private static List<Dictionary<string, object>> BuildMaterialList(Document doc, Dictionary<ElementId, List<ElementId>> materialToElements)
    {
        var list = new List<Dictionary<string, object>>();
        foreach (var kvp in materialToElements)
        {
            Material mat = doc.GetElement(kvp.Key) as Material;
            if (mat == null) continue;
            list.Add(new Dictionary<string, object>
            {
                { "Name", mat.Name },
                { "Material Class", mat.MaterialClass ?? "" },
                { "Count", kvp.Value.Count },
                { "ElementIdObject", kvp.Key },
                { "_ElementIds", kvp.Value }
            });
        }
        return list.OrderBy(e => e["Name"].ToString()).ToList();
    }

    private static List<ElementId> GatherElementIds(List<Dictionary<string, object>> selected)
    {
        var ids = new List<ElementId>();
        foreach (var row in selected)
        {
            if (row.ContainsKey("_ElementIds"))
                ids.AddRange((List<ElementId>)row["_ElementIds"]);
        }
        return ids;
    }
}
