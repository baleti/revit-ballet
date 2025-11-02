using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ZoomSelected : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;
        
        // Get references which includes linked elements
        IList<Reference> references = uiDoc.GetReferences();
        
        // Also get regular selection IDs
        ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
        
        if ((references == null || references.Count == 0) && selectedIds.Count == 0)
        {
            message = "No elements selected.";
            return Result.Failed;
        }
        
        BoundingBoxXYZ boundingBox = CalculateBoundingBox(doc, selectedIds, references);
        
        if (boundingBox == null)
        {
            message = "No valid bounding box.";
            return Result.Failed;
        }
        
        // Check if current view is 3D
        View3D view3D = doc.ActiveView as View3D;
        if (view3D != null)
        {
            if (!ZoomToBoundingBox3D(uiDoc, view3D, boundingBox))
            {
                message = "Failed to zoom to bounding box in 3D view.";
                return Result.Failed;
            }
        }
        else
        {
            if (!ZoomToBoundingBox2D(uiDoc, doc, boundingBox))
            {
                message = "Failed to find UIView for active view.";
                return Result.Failed;
            }
        }
        
        return Result.Succeeded;
    }
    
    private BoundingBoxXYZ CalculateBoundingBox(Document doc, ICollection<ElementId> selectedIds, IList<Reference> references)
    {
        BoundingBoxXYZ boundingBox = null;
        
        // First handle regular selected elements
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
        
        // Handle references (which may include linked elements)
        if (references != null)
        {
            foreach (Reference reference in references)
            {
                if (reference.LinkedElementId != ElementId.InvalidElementId)
                {
                    // This is a linked element
                    RevitLinkInstance linkInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkInstance != null)
                    {
                        Document linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            Element linkedElement = linkedDoc.GetElement(reference.LinkedElementId);
                            if (linkedElement != null)
                            {
                                // Get bounding box in linked document's view
                                View linkedView = FindCorrespondingView(linkedDoc, doc.ActiveView);
                                BoundingBoxXYZ linkedBox = linkedElement.get_BoundingBox(linkedView ?? linkedDoc.ActiveView);
                                
                                if (linkedBox != null)
                                {
                                    // Transform the bounding box to host coordinates
                                    Transform transform = linkInstance.GetTotalTransform();
                                    BoundingBoxXYZ transformedBox = TransformBoundingBox(linkedBox, transform);
                                    
                                    boundingBox = boundingBox == null ? transformedBox : UnionBoundingBox(boundingBox, transformedBox);
                                }
                            }
                        }
                    }
                }
                else if (!selectedIds.Contains(reference.ElementId))
                {
                    // Regular reference not already handled
                    Element elem = doc.GetElement(reference.ElementId);
                    BoundingBoxXYZ elemBox = elem?.get_BoundingBox(doc.ActiveView);
                    if (elemBox != null)
                    {
                        elemBox = AdjustBoundingBox(elemBox);
                        boundingBox = boundingBox == null ? elemBox : UnionBoundingBox(boundingBox, elemBox);
                    }
                }
            }
        }
        
        return boundingBox;
    }
    
    private bool ZoomToBoundingBox3D(UIDocument uiDoc, View3D view3D, BoundingBoxXYZ boundingBox)
    {
        UIView uiview = FindUIView(uiDoc, view3D.Id);
        if (uiview == null)
        {
            return false;
        }
        
        // Get the center of the bounding box - this will be our target/orbit point
        XYZ targetPoint = GetCenterPoint(boundingBox);
        
        // Get current view orientation
        ViewOrientation3D currentOrientation = view3D.GetOrientation();
        XYZ currentEyePosition = currentOrientation.EyePosition;
        XYZ currentUpDirection = currentOrientation.UpDirection;
        XYZ currentForwardDirection = currentOrientation.ForwardDirection;
        
        // Calculate the size of the bounding box
        double width = boundingBox.Max.X - boundingBox.Min.X;
        double height = boundingBox.Max.Y - boundingBox.Min.Y;
        double depth = boundingBox.Max.Z - boundingBox.Min.Z;
        double maxDimension = Math.Max(Math.Max(width, height), depth);
        
        if (view3D.IsPerspective)
        {
            // For perspective views, we need to position the camera to look at the target
            // while maintaining a good viewing distance
            
            // Calculate appropriate viewing distance based on field of view
            // Revit's default FOV is about 50 degrees
            double fieldOfView = 50.0 * Math.PI / 180.0; // Convert to radians
            double halfFOV = fieldOfView / 2.0;
            
            // Use a more conservative viewing distance to avoid panning issues
            // The multiplier affects how much of the view is filled
            double viewingDistance = (maxDimension * 1.2) / Math.Tan(halfFOV);
            
            // Ensure minimum viewing distance to avoid being too close
            viewingDistance = Math.Max(viewingDistance, maxDimension * 2.0);
            
            // Position the camera at the calculated distance from the target
            XYZ newEyePosition = targetPoint - currentForwardDirection * viewingDistance;
            
            // Create new orientation looking at the target
            ViewOrientation3D newOrientation = new ViewOrientation3D(newEyePosition, currentUpDirection, currentForwardDirection);
            
            using (Transaction t = new Transaction(view3D.Document, "Set 3D View Orientation"))
            {
                t.Start();
                view3D.SetOrientation(newOrientation);
                t.Commit();
            }
            
            // Force viewport refresh
            uiDoc.RefreshActiveView();
            
            // Additional refresh for perspective views to ensure proper update
            if (view3D.IsPerspective)
            {
                // This helps reset the internal navigation speed
                uiDoc.UpdateAllOpenViews();
            }
        }
        else
        {
            // For orthographic views, use the zoom corners approach
            XYZ rightDirection = currentForwardDirection.CrossProduct(currentUpDirection);
            
            // Calculate the scale based on view size
            double scale = maxDimension * 0.6;
            
            // Calculate corners for zoom rectangle
            XYZ corner1 = targetPoint + currentUpDirection * scale - rightDirection * scale;
            XYZ corner2 = targetPoint - currentUpDirection * scale + rightDirection * scale;
            
            // First set the view orientation to center on target
            using (Transaction t = new Transaction(view3D.Document, "Center 3D View"))
            {
                t.Start();
                
                // Move the eye position to look at the center
                // For orthographic, we maintain the viewing direction but shift the position
                XYZ offset = targetPoint - GetViewCenter(uiview, view3D);
                XYZ newEyePosition = currentEyePosition + offset;
                
                ViewOrientation3D newOrientation = new ViewOrientation3D(
                    newEyePosition, 
                    currentUpDirection, 
                    currentForwardDirection
                );
                
                view3D.SetOrientation(newOrientation);
                t.Commit();
            }
            
            // Then zoom to the appropriate size
            uiview.ZoomAndCenterRectangle(corner1, corner2);
            
            // Refresh the view
            uiDoc.RefreshActiveView();
        }
        
        return true;
    }
    
    private XYZ GetViewCenter(UIView uiView, View3D view3D)
    {
        // Get the center of the current view
        IList<XYZ> corners = uiView.GetZoomCorners();
        if (corners.Count >= 2)
        {
            return (corners[0] + corners[1]) * 0.5;
        }
        
        // Fallback to view orientation
        ViewOrientation3D orientation = view3D.GetOrientation();
        return orientation.EyePosition + orientation.ForwardDirection * 10; // Arbitrary distance
    }
    
    private bool ZoomToBoundingBox2D(UIDocument uiDoc, Document doc, BoundingBoxXYZ boundingBox)
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
        double scaleFactor = Math.Max(width, height) * 0.6;
        
        XYZ corner1 = center - rightDirection * scaleFactor - upDirection * scaleFactor;
        XYZ corner2 = center + rightDirection * scaleFactor + upDirection * scaleFactor;
        
        uiview.ZoomAndCenterRectangle(corner1, corner2);
        return true;
    }
    
    private View FindCorrespondingView(Document linkedDoc, View hostView)
    {
        // Try to find a view in the linked document with the same name or type
        FilteredElementCollector collector = new FilteredElementCollector(linkedDoc)
            .OfClass(typeof(View));
        
        foreach (View view in collector)
        {
            if (view.Name == hostView.Name && view.GetType() == hostView.GetType())
            {
                return view;
            }
        }
        
        // If no matching view found, try to find any 3D view
        if (hostView is View3D)
        {
            return collector.OfType<View3D>().FirstOrDefault(v => !v.IsTemplate);
        }
        
        return null;
    }
    
    private BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ box, Transform transform)
    {
        // Transform all 8 corners of the bounding box
        XYZ[] corners = new XYZ[]
        {
            new XYZ(box.Min.X, box.Min.Y, box.Min.Z),
            new XYZ(box.Max.X, box.Min.Y, box.Min.Z),
            new XYZ(box.Min.X, box.Max.Y, box.Min.Z),
            new XYZ(box.Max.X, box.Max.Y, box.Min.Z),
            new XYZ(box.Min.X, box.Min.Y, box.Max.Z),
            new XYZ(box.Max.X, box.Min.Y, box.Max.Z),
            new XYZ(box.Min.X, box.Max.Y, box.Max.Z),
            new XYZ(box.Max.X, box.Max.Y, box.Max.Z)
        };
        
        // Transform each corner
        XYZ[] transformedCorners = corners.Select(c => transform.OfPoint(c)).ToArray();
        
        // Find the min and max of transformed corners
        double minX = transformedCorners.Min(p => p.X);
        double minY = transformedCorners.Min(p => p.Y);
        double minZ = transformedCorners.Min(p => p.Z);
        double maxX = transformedCorners.Max(p => p.X);
        double maxY = transformedCorners.Max(p => p.Y);
        double maxZ = transformedCorners.Max(p => p.Z);
        
        BoundingBoxXYZ transformedBox = new BoundingBoxXYZ();
        transformedBox.Min = new XYZ(minX, minY, minZ);
        transformedBox.Max = new XYZ(maxX, maxY, maxZ);
        
        return transformedBox;
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
        XYZ min = new XYZ(
            Math.Min(a.Min.X, b.Min.X), 
            Math.Min(a.Min.Y, b.Min.Y), 
            Math.Min(a.Min.Z, b.Min.Z)
        );
        XYZ max = new XYZ(
            Math.Max(a.Max.X, b.Max.X), 
            Math.Max(a.Max.Y, b.Max.Y), 
            Math.Max(a.Max.Z, b.Max.Z)
        );
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
