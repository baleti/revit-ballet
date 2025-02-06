using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System;

namespace RevitCommands
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ListBoundingBoxCoordinates : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get selected elements
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("Error", "Please select at least one element.");
                return Result.Failed;
            }

            List<Dictionary<string, object>> elementData = new List<Dictionary<string, object>>();
            
            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                BoundingBoxXYZ bbox = elem.get_BoundingBox(null);
                
                if (bbox == null) continue;

                // Get element's transform if it's a family instance
                Transform transform = Transform.Identity;
                if (elem is FamilyInstance familyInstance)
                {
                    transform = familyInstance.GetTransform();
                }
                
                // Transform coordinates to project coordinates
                XYZ minPoint = transform.OfPoint(bbox.Min);
                XYZ maxPoint = transform.OfPoint(bbox.Max);

                // Calculate dimensions
                double width = Math.Abs(maxPoint.X - minPoint.X);
                double depth = Math.Abs(maxPoint.Y - minPoint.Y);
                double height = Math.Abs(maxPoint.Z - minPoint.Z);

                Dictionary<string, object> data = new Dictionary<string, object>
                {
                    // Basic element parameters
                    {"ElementId", elem.Id.IntegerValue},
                    {"Name", elem.Name},
                    {"Category", elem.Category?.Name ?? "N/A"},
                    {"Family", (elem as FamilyInstance)?.Symbol?.Family?.Name ?? "N/A"},
                    {"Type", elem.GetType().Name},
                    
                    // Dimensions
                    {"Width", Math.Round(width * 304.8, 2)}, // Convert to mm
                    {"Depth", Math.Round(depth * 304.8, 2)},
                    {"Height", Math.Round(height * 304.8, 2)},
                    
                    // Min point coordinates
                    {"Min X", Math.Round(minPoint.X * 304.8, 2)},
                    {"Min Y", Math.Round(minPoint.Y * 304.8, 2)},
                    {"Min Z", Math.Round(minPoint.Z * 304.8, 2)},
                    
                    // Max point coordinates
                    {"Max X", Math.Round(maxPoint.X * 304.8, 2)},
                    {"Max Y", Math.Round(maxPoint.Y * 304.8, 2)},
                    {"Max Z", Math.Round(maxPoint.Z * 304.8, 2)},
                };

                // Add common parameters
                AddCommonParameters(elem, data);

                elementData.Add(data);
            }

            // Define column order
            List<string> propertyNames = new List<string>
            {
                "ElementId", "Name", "Category", "Family", "Type",
                "Width", "Depth", "Height",
                "Min X", "Min Y", "Min Z",
                "Max X", "Max Y", "Max Z",
                // Additional parameters will be added dynamically
            };

            // Add any additional parameter names that were found
            if (elementData.Count > 0)
            {
                var additionalParams = elementData[0].Keys
                    .Where(k => !propertyNames.Contains(k))
                    .OrderBy(k => k);
                propertyNames.AddRange(additionalParams);
            }

            // Show the DataGrid
            try
            {
                var results = CustomGUIs.DataGrid(elementData, propertyNames, false);
                
                // If user confirmed selection (didn't press Escape)
                if (results != null && results.Any())
                {
                    // Get the ElementIds from the selected rows
                    var selectedElementIds = new List<ElementId>();
                    foreach (var result in results)
                    {
                        if (result.ContainsKey("ElementId") && result["ElementId"] is int elementIdInt)
                        {
                            selectedElementIds.Add(new ElementId(elementIdInt));
                        }
                    }

                    // Update Revit's selection
                    using (Transaction trans = new Transaction(doc, "Update Selection"))
                    {
                        trans.Start();
                        uidoc.Selection.SetElementIds(selectedElementIds);
                        trans.Commit();
                    }
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void AddCommonParameters(Element elem, Dictionary<string, object> data)
        {
            // Get all parameters of the element
            foreach (Parameter param in elem.Parameters)
            {
                if (!param.HasValue) continue;

                string paramName = param.Definition.Name;
                if (data.ContainsKey(paramName)) continue; // Skip if already added

                // Get parameter value based on its storage type
                object paramValue = null;
                switch (param.StorageType)
                {
                    case StorageType.Double:
                        paramValue = Math.Round(param.AsDouble() * 304.8, 2); // Convert to mm
                        break;
                    case StorageType.Integer:
                        paramValue = param.AsInteger();
                        break;
                    case StorageType.String:
                        paramValue = param.AsString();
                        break;
                    case StorageType.ElementId:
                        paramValue = param.AsElementId().IntegerValue;
                        break;
                }

                if (paramValue != null)
                {
                    data[paramName] = paramValue;
                }
            }
        }
    }
}
