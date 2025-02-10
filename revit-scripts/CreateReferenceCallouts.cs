using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using System.Collections.Generic;
using System;

namespace RevitDetailLines
{
    [Transaction(TransactionMode.Manual)]
    public class CreateReferenceCalloutsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            View activeView = doc.ActiveView;

            // Get selected elements
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            if (!selectedIds.Any())
            {
                message = "Please select at least one element.";
                return Result.Failed;
            }

            try
            {
                // Get all sheets for view lookup
                var sheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .ToDictionary(sheet => sheet.Id);

                // Get all view ports to find sheet associations
                var viewPorts = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .ToList();

                // Dictionary to store view-to-sheet relationships
                var viewSheetMapping = new Dictionary<ElementId, ViewSheet>();
                foreach (var viewport in viewPorts)
                {
                    ElementId sheetId = viewport.SheetId;
                    ElementId viewId = viewport.ViewId;
                    if (sheets.ContainsKey(sheetId))
                    {
                        viewSheetMapping[viewId] = sheets[sheetId];
                    }
                }

                // Collect filtered views
                var allViews = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate && 
                               v.Id != activeView.Id &&
                               (v.ViewType == ViewType.DraftingView ||
                                v.ViewType == ViewType.Detail ||
                                v.ViewType == ViewType.Section ||
                                v.ViewType == ViewType.Elevation))
                    .OrderBy(v => v.ViewType.ToString())
                    .ThenBy(v => v.Name)
                    .ToList();

                // Prepare data for the DataGrid
                var viewEntries = allViews.Select(v =>
                {
                    ViewSheet sheet = null;
                    viewSheetMapping.TryGetValue(v.Id, out sheet);

                    return new Dictionary<string, object>
                    {
                        { "View Name", v.Name },
                        { "View Type", v.ViewType.ToString() },
                        { "Sheet Number", sheet?.SheetNumber ?? "" },
                        { "Sheet Name", sheet?.Name ?? "" },
                        { "Sheet Folder", sheet?.LookupParameter("Sheet Folder")?.AsString() ?? "" },
                        { "View Id", v.Id.IntegerValue.ToString() }
                    };
                }).ToList();

                var propertyNames = new List<string> 
                { 
                    "View Name",
                    "View Type",
                    "Sheet Number",
                    "Sheet Name",
                    "Sheet Folder",
                    "View Id"
                };

                // Show DataGrid to user
                var selection = CustomGUIs.DataGrid(viewEntries, propertyNames, false);

                // Check if user made a selection
                if (selection == null || !selection.Any())
                {
                    return Result.Cancelled;
                }

                // Get the selected view
                int selectedViewId = int.Parse(selection.First()["View Id"].ToString());
                View selectedView = doc.GetElement(new ElementId(selectedViewId)) as View;

                if (selectedView == null)
                {
                    message = "Selected view not found.";
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Create Reference Callouts"))
                {
                    trans.Start();

                    foreach (ElementId elementId in selectedIds)
                    {
                        Element elem = doc.GetElement(elementId);
                        if (elem == null) continue;

                        // Get element's bounding box in the view
                        BoundingBoxXYZ bbox = elem.get_BoundingBox(activeView);
                        if (bbox == null)
                        {
                            // Try getting the geometry bounding box if view bbox is not available
                            Options geomOptions = new Options
                            {
                                View = activeView,
                                ComputeReferences = true
                            };
                            GeometryElement geomElem = elem.get_Geometry(geomOptions);
                            if (geomElem != null)
                            {
                                bbox = geomElem.GetBoundingBox();
                            }
                        }

                        if (bbox != null)
                        {
                            // Calculate the size of the callout based on the element size
                            XYZ center = (bbox.Min + bbox.Max) * 0.5;
                            double width = bbox.Max.X - bbox.Min.X;
                            double height = bbox.Max.Y - bbox.Min.Y;
                            
                            // Add a small padding around the element
                            double padding = Math.Max(width, height) * 0.1;
                            XYZ point1 = new XYZ(bbox.Min.X - padding, bbox.Min.Y - padding, bbox.Min.Z);
                            XYZ point2 = new XYZ(bbox.Max.X + padding, bbox.Max.Y + padding, bbox.Max.Z);

                            // Create the reference callout
                            ViewSection.CreateReferenceCallout(
                                doc,
                                activeView.Id,
                                selectedView.Id,
                                point1,
                                point2
                            );
                        }
                    }

                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
