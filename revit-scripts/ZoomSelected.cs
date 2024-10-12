using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ZoomSelected : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;
        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

        if (selectedIds.Count == 0)
        {
            message = "No elements selected.";
            return Result.Failed;
        }

        BoundingBoxXYZ boundingBox = CalculateBoundingBox(doc, selectedIds);
        if (boundingBox == null)
        {
            message = "No valid bounding box.";
            return Result.Failed;
        }

        if (!ZoomToBoundingBox(uiDoc, doc, boundingBox))
        {
            message = "Failed to find UIView for active view.";
            return Result.Failed;
        }

        return Result.Succeeded;
    }

    private BoundingBoxXYZ CalculateBoundingBox(Document doc, ICollection<ElementId> selectedIds)
    {
        BoundingBoxXYZ boundingBox = null;
        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            BoundingBoxXYZ elemBox = elem?.get_BoundingBox(doc.ActiveView);
            if (elemBox != null)
            {
                elemBox = AdjustBoundingBox(elemBox);
                boundingBox = boundingBox == null ? elemBox : UnionBoundingBox(boundingBox, elemBox);
            }
        }
        return boundingBox;
    }

    private bool ZoomToBoundingBox(UIDocument uiDoc, Document doc, BoundingBoxXYZ boundingBox)
    {
        UIView uiview = FindUIView(uiDoc, doc.ActiveView.Id);
        if (uiview == null)
        {
            return false;
        }

        XYZ center = GetCenterPoint(boundingBox);
        XYZ viewDirection = doc.ActiveView.ViewDirection;
        XYZ upDirection = doc.ActiveView.UpDirection;
        XYZ rightDirection = viewDirection.CrossProduct(upDirection);

        // Adjust scale factors based on bounding box size
        double width = boundingBox.Max.X - boundingBox.Min.X;
        double height = boundingBox.Max.Y - boundingBox.Min.Y;
        double scaleFactor = Math.Max(width, height) * 0.6; // Adjust this factor as necessary

        XYZ corner1 = center - rightDirection * scaleFactor - upDirection * scaleFactor;
        XYZ corner2 = center + rightDirection * scaleFactor + upDirection * scaleFactor;

        uiview.ZoomAndCenterRectangle(corner1, corner2);
        return true;
    }

    private UIView FindUIView(UIDocument uiDoc, ElementId viewId)
    {
        foreach (UIView uiview in uiDoc.GetOpenUIViews())
        {
            if (uiview.ViewId.Equals(viewId))
            {
                return uiview;
            }
        }
        return null;
    }

    private BoundingBoxXYZ UnionBoundingBox(BoundingBoxXYZ a, BoundingBoxXYZ b)
    {
        XYZ min = new XYZ(Math.Min(a.Min.X, b.Min.X), Math.Min(a.Min.Y, b.Min.Y), Math.Min(a.Min.Z, b.Min.Z));
        XYZ max = new XYZ(Math.Max(a.Max.X, b.Max.X), Math.Max(a.Max.Y, b.Max.Y), Math.Max(a.Max.Z, b.Max.Z));
        return new BoundingBoxXYZ { Min = min, Max = max };
    }

    private BoundingBoxXYZ AdjustBoundingBox(BoundingBoxXYZ box)
    {
        // This function now does nothing, but you can adjust logic here if needed
        return box;
    }

    private XYZ GetCenterPoint(BoundingBoxXYZ boundingBox)
    {
        return (boundingBox.Min + boundingBox.Max) * 0.5;
    }
}
