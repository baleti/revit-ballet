using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class DirectShapeBoundingBoxFromSelectedElements : IExternalCommand
{
    public Result Execute(ExternalCommandData data, ref string message, ElementSet elems)
    {
        UIDocument uidoc = data.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;

        // -------------------------------------------------- 1. collect selected elements
        IList<Element> elements = GetSelectedElements(uidoc);
        if (elements.Count == 0)
        {
            message = "No elements selected. Please select one or more elements.";
            return Result.Failed;
        }

        // -------------------------------------------------- 2. create DirectShapes
        List<ElementId> createdIds = new List<ElementId>();
        
        using (Transaction t = new Transaction(doc, "Create DirectShape Bounding Boxes"))
        {
            t.Start();
            foreach (Element elem in elements)
            {
                ElementId dsId = TryCreateBoundingBoxDirectShape(doc, elem);
                if (dsId != null && dsId != ElementId.InvalidElementId)
                    createdIds.Add(dsId);
            }

            if (createdIds.Count == 0)
            {
                message = "No DirectShapes could be created (elements may not have valid geometry).";
                t.RollBack();
                return Result.Failed;
            }
            t.Commit();
        }
        
        // Add newly created DirectShapes to current selection
        var currentSelection = uidoc.GetSelectionIds().ToList();
        currentSelection.AddRange(createdIds);
        uidoc.SetSelectionIds(currentSelection);
        
        return Result.Succeeded;
    }

    // --------------------------------------------------------------------------
    IList<Element> GetSelectedElements(UIDocument uidoc)
    {
        // Get currently selected elements
        var selection = uidoc.GetSelectionIds();
        
        // Get all selected elements that have a valid bounding box
        var elements = selection
            .Select(id => uidoc.Document.GetElement(id))
            .Where(e => e != null && e.get_BoundingBox(null) != null)
            .ToList();

        return elements;
    }

    // --------------------------------------------------------------------------
    ElementId TryCreateBoundingBoxDirectShape(Document doc, Element sourceElement)
    {
        // Get the bounding box of the element
        BoundingBoxXYZ bbox = sourceElement.get_BoundingBox(null);
        if (bbox == null || !bbox.Enabled)
            return ElementId.InvalidElementId;

        // Get min and max points
        XYZ min = bbox.Min;
        XYZ max = bbox.Max;

        // Check if bounding box has valid dimensions
        double width = max.X - min.X;
        double depth = max.Y - min.Y;
        double height = max.Z - min.Z;
        
        if (width <= 0 || depth <= 0 || height <= 0)
            return ElementId.InvalidElementId;

        // Create a solid box from the bounding box
        Solid boxSolid = null;
        try
        {
            // Create the bottom face profile (rectangle at min Z)
            XYZ p0 = new XYZ(min.X, min.Y, min.Z);
            XYZ p1 = new XYZ(max.X, min.Y, min.Z);
            XYZ p2 = new XYZ(max.X, max.Y, min.Z);
            XYZ p3 = new XYZ(min.X, max.Y, min.Z);

            // Create curves for the rectangle
            Line line0 = Line.CreateBound(p0, p1);
            Line line1 = Line.CreateBound(p1, p2);
            Line line2 = Line.CreateBound(p2, p3);
            Line line3 = Line.CreateBound(p3, p0);

            // Create curve loop
            List<Curve> curves = new List<Curve> { line0, line1, line2, line3 };
            CurveLoop curveLoop = CurveLoop.Create(curves);

            // Create solid by extrusion
            boxSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                new[] { curveLoop }, 
                XYZ.BasisZ, 
                height);
        }
        catch
        {
            return ElementId.InvalidElementId;
        }

        if (boxSolid == null)
            return ElementId.InvalidElementId;

        // Create DirectShape
        DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        if (ds == null)
            return ElementId.InvalidElementId;

        try 
        { 
            ds.SetShape(new GeometryObject[] { boxSolid }); 
        }
        catch 
        { 
            doc.Delete(ds.Id); 
            return ElementId.InvalidElementId; 
        }

        // Set name based on source element
        string elementName = GetElementName(sourceElement);
        ds.Name = $"BBox_{elementName}";

        // Set Comments with Type and Id information as key-value pairs
        string elementTypeName = GetElementTypeName(doc, sourceElement);
        string comments = $"Type: {elementTypeName}, Id: {sourceElement.Id.IntegerValue}";
        
        Parameter commentsParam = ds.LookupParameter("Comments") ??
                                 ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (commentsParam != null && commentsParam.StorageType == StorageType.String)
        {
            commentsParam.Set(comments);
        }

        return ds.Id;
    }

    // --------------------------------------------------------------------------
    string GetElementName(Element elem)
    {
        // Try to get element name
        string name = elem.Name;
        
        // If Name is empty, try to get type name
        if (string.IsNullOrWhiteSpace(name))
        {
            Parameter typeParam = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
            if (typeParam != null)
            {
                ElementId typeId = typeParam.AsElementId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    Element elemType = elem.Document.GetElement(typeId);
                    if (elemType != null)
                        name = elemType.Name;
                }
            }
        }

        // If still no name, use category name
        if (string.IsNullOrWhiteSpace(name))
        {
            if (elem.Category != null)
                name = elem.Category.Name;
            else
                name = "Element";
        }

        // Sanitize the name for DirectShape naming
        return Sanitize(name);
    }

    // --------------------------------------------------------------------------
    string GetElementTypeName(Document doc, Element elem)
    {
        string typeName = elem.GetType().Name;
        
        // Try to get the Revit element type name
        ElementId typeId = elem.GetTypeId();
        if (typeId != null && typeId != ElementId.InvalidElementId)
        {
            Element elemType = doc.GetElement(typeId);
            if (elemType != null && !string.IsNullOrWhiteSpace(elemType.Name))
            {
                typeName = $"{elem.Category?.Name ?? typeName}: {elemType.Name}";
            }
        }
        else if (elem.Category != null)
        {
            typeName = elem.Category.Name;
        }

        return typeName;
    }

    // --------------------------------------------------------------------------
    static string Sanitize(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return "Unnamed";
            
        // Remove illegal characters for Revit names
        System.Text.RegularExpressions.Regex illegal = 
            new System.Text.RegularExpressions.Regex(@"[<>:{}|;?*\\/\[\]]");
        string clean = illegal.Replace(raw, "_").Trim();
        
        // Limit length
        return clean.Length > 250 ? clean.Substring(0, 250) : clean;
    }
}
