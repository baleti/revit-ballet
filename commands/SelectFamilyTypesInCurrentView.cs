using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Architecture;

[Transaction(TransactionMode.Manual)]
public class SelectFamilyTypesInCurrentView : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
        Dictionary<string, Element> typeElementMap = new Dictionary<string, Element>();

        ElementId currentViewId = doc.ActiveView.Id;

        // Collect ALL elements in the current view without category restrictions
        var elementsInView = new FilteredElementCollector(doc, currentViewId)
            .WhereElementIsNotElementType()
            .Where(e => e.Category != null && IsElementVisibleInView(e, doc.ActiveView))
            .ToList();

        // Process different types of elements
        foreach (var element in elementsInView)
        {
            Element typeElement = null;
            string typeName = "";
            string familyName = "";

            if (element is FamilyInstance familyInstance)
            {
                typeElement = familyInstance.Symbol;
                typeName = familyInstance.Symbol.Name;
                familyName = familyInstance.Symbol.FamilyName;
            }
            else if (element is MEPCurve mepCurve)
            {
                typeElement = doc.GetElement(mepCurve.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    if (element is Pipe)
                        familyName = "Pipe";
                    else if (element is Duct)
                        familyName = "Duct";
                    else if (element is Conduit)
                        familyName = "Conduit";
                    else if (element is CableTray)
                        familyName = "Cable Tray";
                    else
                        familyName = "MEP Curve";
                }
            }
            else if (element is Wall wall)
            {
                typeElement = doc.GetElement(wall.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    familyName = GetSystemFamilyName(typeElement) ?? "Wall";
                }
            }
            else if (element is Floor floor)
            {
                typeElement = doc.GetElement(floor.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    familyName = GetSystemFamilyName(typeElement) ?? "Floor";
                }
            }
            else if (element is RoofBase roof)
            {
                typeElement = doc.GetElement(roof.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    familyName = GetSystemFamilyName(typeElement) ?? "Roof";
                }
            }
            else if (element is Ceiling ceiling)
            {
                typeElement = doc.GetElement(ceiling.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    familyName = GetSystemFamilyName(typeElement) ?? "Ceiling";
                }
            }
            else if (element is Grid grid)
            {
                typeElement = doc.GetElement(grid.GetTypeId());
                if (typeElement != null)
                {
                    typeName = typeElement.Name;
                    familyName = "Grid";
                }
            }
            else if (element is Level level)
            {
                typeName = level.Name;
                familyName = "Level";
                typeElement = element;
            }
            else if (element is ReferencePlane refPlane)
            {
                typeName = refPlane.Name;
                familyName = "Reference Plane";
                typeElement = element;
            }
            else if (element is Room room)
            {
                typeName = room.Name;
                familyName = "Room";
                typeElement = element;
            }
            else if (element is Area area)
            {
                typeName = area.Name;
                familyName = "Area";
                typeElement = element;
            }
            else if (element is Space space)
            {
                typeName = space.Name;
                familyName = "Space";
                typeElement = element;
            }
            else
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    typeElement = doc.GetElement(typeId);
                    if (typeElement != null)
                    {
                        typeName = typeElement.Name;
                        if (typeElement is FamilySymbol fs)
                        {
                            familyName = fs.FamilyName;
                        }
                        else
                        {
                            familyName = GetSystemFamilyName(typeElement) ?? element.GetType().Name;
                        }
                    }
                }
                else
                {
                    typeName = element.Name;
                    familyName = element.GetType().Name;
                    typeElement = element;
                }
            }

            if (typeElement != null && !string.IsNullOrEmpty(typeName))
            {
                string uniqueKey = $"{familyName}:{typeName}:{element.Category.Name}";
                if (!typeElementMap.ContainsKey(uniqueKey))
                {
                    var entry = new Dictionary<string, object>
                    {
                        { "Type Name", typeName },
                        { "Family", familyName },
                        { "Category", element.Category.Name }
                    };

                    typeElementMap[uniqueKey] = typeElement;
                    typeEntries.Add(entry);
                }
            }
        }

        // Sort by Category, then Family, then Type Name
        typeEntries = typeEntries
            .OrderBy(e => e["Category"].ToString())
            .ThenBy(e => e["Family"].ToString())
            .ThenBy(e => e["Type Name"].ToString())
            .ToList();

        // Display the list of unique types
        var propertyNames = new List<string> { "Category", "Family", "Type Name" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, false);

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled;
        }

        // Collect ElementIds of the selected types
        List<ElementId> selectedTypeIds = selectedEntries
            .Select(entry =>
            {
                string uniqueKey = $"{entry["Family"]}:{entry["Type Name"]}:{entry["Category"]}";
                return typeElementMap[uniqueKey].Id;
            })
            .ToList();

        // Set the selection to the selected types
        uidoc.SetSelectionIds(selectedTypeIds);

        return Result.Succeeded;
    }

    private string GetSystemFamilyName(Element typeElement)
    {
        Parameter familyParam = typeElement.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
        return (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
            ? familyParam.AsString()
            : null;
    }

    private bool IsElementVisibleInView(Element element, View view)
    {
        BoundingBoxXYZ bbox = element.get_BoundingBox(view);

        if (bbox == null && element.IsHidden(view))
        {
            return false;
        }

        return bbox != null || !element.IsHidden(view);
    }
}
