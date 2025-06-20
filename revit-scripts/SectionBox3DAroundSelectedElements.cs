using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SectionBox3DAroundSelectedElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;
                View3D view3D = doc.ActiveView as View3D;

                // Check if active view is 3D view
                if (view3D == null)
                {
                    TaskDialog.Show("Error", "Please switch to a 3D view to use this command.");
                    return Result.Failed;
                }

                // Check if view is a template
                if (view3D.IsTemplate)
                {
                    TaskDialog.Show("Error", "Cannot apply section box to a view template.");
                    return Result.Failed;
                }

                // Get current selection
                ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
                
                // For linked elements selected with TAB, we need to check the Selection's PickedBoxes
                // This handles elements that don't show up in GetElementIds()
                ICollection<Reference> selectedRefs = uiDoc.Selection.GetReferences();
                
                if (selectedIds.Count == 0 && (selectedRefs == null || selectedRefs.Count == 0))
                {
                    TaskDialog.Show("Error", "Please select elements before running this command.");
                    return Result.Failed;
                }

                // Collect all bounding boxes
                BoundingBoxXYZ combinedBBox = null;
                
                // Process regular selected elements
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem == null) continue;
                    
                    // Check if it's a RevitLinkInstance
                    if (elem is RevitLinkInstance linkInstance)
                    {
                        // User selected the entire linked model
                        BoundingBoxXYZ linkBBox = linkInstance.get_BoundingBox(view3D);
                        if (linkBBox != null && linkBBox.Enabled)
                        {
                            UpdateCombinedBoundingBox(ref combinedBBox, linkBBox);
                        }
                    }
                    else
                    {
                        // Regular element - get its bounding box
                        BoundingBoxXYZ elemBBox = elem.get_BoundingBox(view3D);
                        
                        // If no bounding box from view, try geometry
                        if (elemBBox == null || !elemBBox.Enabled)
                        {
                            Options geomOptions = new Options
                            {
                                ComputeReferences = false,
                                DetailLevel = ViewDetailLevel.Coarse,
                                IncludeNonVisibleObjects = false
                            };

                            GeometryElement geomElem = elem.get_Geometry(geomOptions);
                            if (geomElem != null)
                            {
                                elemBBox = geomElem.GetBoundingBox();
                            }
                        }

                        if (elemBBox != null && elemBBox.Enabled)
                        {
                            UpdateCombinedBoundingBox(ref combinedBBox, elemBBox);
                        }
                    }
                }
                
                // Process selected references (for linked elements selected with TAB)
                if (selectedRefs != null && selectedRefs.Count > 0)
                {
                    foreach (Reference reference in selectedRefs)
                    {
                        if (reference.LinkedElementId != ElementId.InvalidElementId)
                        {
                            // This is a linked element
                            RevitLinkInstance linkInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                            if (linkInstance != null)
                            {
                                Document linkDoc = linkInstance.GetLinkDocument();
                                if (linkDoc != null)
                                {
                                    Element linkedElem = linkDoc.GetElement(reference.LinkedElementId);
                                    if (linkedElem != null)
                                    {
                                        BoundingBoxXYZ linkedBBox = linkedElem.get_BoundingBox(null);
                                        if (linkedBBox != null && linkedBBox.Enabled)
                                        {
                                            // Transform the bounding box to host coordinates
                                            Transform totalTransform = linkInstance.GetTotalTransform();
                                            BoundingBoxXYZ transformedBBox = TransformBoundingBox(linkedBBox, totalTransform);
                                            
                                            if (transformedBBox != null)
                                            {
                                                UpdateCombinedBoundingBox(ref combinedBBox, transformedBBox);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if (combinedBBox == null)
                {
                    TaskDialog.Show("Error", "Could not determine bounding box for selected elements.");
                    return Result.Failed;
                }

                // Add a small offset to the bounding box for better visibility
                double offset = 1.0; // 1 foot offset
                combinedBBox.Min = new XYZ(
                    combinedBBox.Min.X - offset,
                    combinedBBox.Min.Y - offset,
                    combinedBBox.Min.Z - offset
                );
                
                combinedBBox.Max = new XYZ(
                    combinedBBox.Max.X + offset,
                    combinedBBox.Max.Y + offset,
                    combinedBBox.Max.Z + offset
                );

                // Apply section box to the 3D view
                using (Transaction trans = new Transaction(doc, "Set Section Box"))
                {
                    trans.Start();
                    
                    // Enable section box if not already enabled
                    if (!view3D.IsSectionBoxActive)
                    {
                        view3D.IsSectionBoxActive = true;
                    }
                    
                    // Set the section box
                    view3D.SetSectionBox(combinedBBox);
                    
                    trans.Commit();
                }

                // Refresh the view
                uiDoc.RefreshActiveView();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        /// <summary>
        /// Updates the combined bounding box with a new bounding box
        /// </summary>
        private void UpdateCombinedBoundingBox(ref BoundingBoxXYZ combinedBBox, BoundingBoxXYZ newBBox)
        {
            if (combinedBBox == null)
            {
                combinedBBox = new BoundingBoxXYZ
                {
                    Min = new XYZ(newBBox.Min.X, newBBox.Min.Y, newBBox.Min.Z),
                    Max = new XYZ(newBBox.Max.X, newBBox.Max.Y, newBBox.Max.Z)
                };
            }
            else
            {
                combinedBBox.Min = new XYZ(
                    Math.Min(combinedBBox.Min.X, newBBox.Min.X),
                    Math.Min(combinedBBox.Min.Y, newBBox.Min.Y),
                    Math.Min(combinedBBox.Min.Z, newBBox.Min.Z)
                );
                
                combinedBBox.Max = new XYZ(
                    Math.Max(combinedBBox.Max.X, newBBox.Max.X),
                    Math.Max(combinedBBox.Max.Y, newBBox.Max.Y),
                    Math.Max(combinedBBox.Max.Z, newBBox.Max.Z)
                );
            }
        }

        /// <summary>
        /// Transforms a bounding box using the given transform
        /// </summary>
        private BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bbox, Transform transform)
        {
            // Get all 8 corners of the bounding box
            XYZ[] corners = new XYZ[]
            {
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z)
            };

            // Transform all corners
            XYZ[] transformedCorners = corners.Select(c => transform.OfPoint(c)).ToArray();

            // Find min and max of transformed corners
            double minX = transformedCorners.Min(p => p.X);
            double minY = transformedCorners.Min(p => p.Y);
            double minZ = transformedCorners.Min(p => p.Z);
            double maxX = transformedCorners.Max(p => p.X);
            double maxY = transformedCorners.Max(p => p.Y);
            double maxZ = transformedCorners.Max(p => p.Z);

            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }
    }
}
