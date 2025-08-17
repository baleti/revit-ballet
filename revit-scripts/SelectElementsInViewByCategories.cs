using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectElementsInViewByCategories : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Document doc = uiapp.ActiveUIDocument.Document;
        View activeView = uiapp.ActiveUIDocument.ActiveView;
        UIDocument uiDoc = uiapp.ActiveUIDocument;
        
        // Collect elements visible in the current view.
        FilteredElementCollector collector = new FilteredElementCollector(doc, activeView.Id);
        
        // Build a unique set of category IDs and track Direct Shapes and regular elements separately.
        HashSet<ElementId> categoryIds = new HashSet<ElementId>();
        Dictionary<ElementId, List<DirectShape>> directShapesByCategory = new Dictionary<ElementId, List<DirectShape>>();
        Dictionary<ElementId, List<ElementId>> regularElementsByCategory = new Dictionary<ElementId, List<ElementId>>();
        
        foreach (Element elem in collector)
        {
            Category category = elem.Category;
            if (category != null)
            {
                categoryIds.Add(category.Id);
                
                // Check if this is a DirectShape
                if (elem is DirectShape directShape)
                {
                    if (!directShapesByCategory.ContainsKey(category.Id))
                    {
                        directShapesByCategory[category.Id] = new List<DirectShape>();
                    }
                    directShapesByCategory[category.Id].Add(directShape);
                }
                else
                {
                    // Store regular (non-DirectShape) element IDs
                    if (!regularElementsByCategory.ContainsKey(category.Id))
                    {
                        regularElementsByCategory[category.Id] = new List<ElementId>();
                    }
                    regularElementsByCategory[category.Id].Add(elem.Id);
                }
            }
        }
        
        // Additional logic for In-View mode to ensure Views (OST_Viewers) is included
        List<ElementId> elementsNotInGroups = new FilteredElementCollector(doc, activeView.Id)
            .WhereElementIsNotElementType()
            .Where(e => e.GroupId == ElementId.InvalidElementId)
            .Where(e => e.Category != null)
            .Where(e => e.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Cameras)
            .Select(e => e.Id)
            .ToList();
        
        // Check if any element belongs to the OST_Viewers category.
        var viewerElements = elementsNotInGroups
            .Select(id => doc.GetElement(id))
            .Where(e => e != null && e.Category != null &&
                      e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Viewers)
            .ToList();
        
        if (viewerElements.Count > 0)
        {
            ElementId viewersCatId = new ElementId((int)BuiltInCategory.OST_Viewers);
            categoryIds.Add(viewersCatId);
            
            foreach (var viewer in viewerElements)
            {
                if (viewer is DirectShape ds)
                {
                    if (!directShapesByCategory.ContainsKey(viewersCatId))
                    {
                        directShapesByCategory[viewersCatId] = new List<DirectShape>();
                    }
                    directShapesByCategory[viewersCatId].Add(ds);
                }
                else
                {
                    if (!regularElementsByCategory.ContainsKey(viewersCatId))
                    {
                        regularElementsByCategory[viewersCatId] = new List<ElementId>();
                    }
                    regularElementsByCategory[viewersCatId].Add(viewer.Id);
                }
            }
        }
        
        // Build a list of dictionaries for the DataGrid.
        List<Dictionary<string, object>> categoryList = new List<Dictionary<string, object>>();
        
        // Group viewer elements by their VIEW_FAMILY parameter for split categories
        Dictionary<string, List<ElementId>> viewersByFamily = new Dictionary<string, List<ElementId>>();
        bool hasViewerCategory = false;
        
        foreach (ElementId id in categoryIds)
        {
            // Handle OST_Viewers specially for in-view mode
            if (id.IntegerValue == (int)BuiltInCategory.OST_Viewers)
            {
                hasViewerCategory = true;
                
                // Group the viewer elements by their VIEW_FAMILY parameter
                var viewers = elementsNotInGroups
                    .Select(eid => doc.GetElement(eid))
                    .Where(e => e != null && e.Category != null &&
                                e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Viewers &&
                                !(e is DirectShape))
                    .ToList();
                
                foreach (var viewer in viewers)
                {
                    string familyName = viewer.get_Parameter(BuiltInParameter.VIEW_FAMILY)?.AsValueString() ?? "Unknown";
                    if (!viewersByFamily.ContainsKey(familyName))
                    {
                        viewersByFamily[familyName] = new List<ElementId>();
                    }
                    viewersByFamily[familyName].Add(viewer.Id);
                }
                
                // Handle Direct Shapes separately for OST_Viewers
                if (directShapesByCategory.ContainsKey(id))
                {
                    int directShapeCount = directShapesByCategory[id].Count;
                    string directShapeName = "Direct Shapes: Views (OST_Viewers)";
                    var entry = new Dictionary<string, object>
                    {
                        { "Name", directShapeName },
                        { "Count", directShapeCount },
                        { "CategoryId", id },
                        { "IsDirectShape", true },
                        { "DirectShapes", directShapesByCategory[id] }
                    };
                    categoryList.Add(entry);
                }
                
                continue; // Skip regular processing for OST_Viewers
            }
            
            Category cat = Category.GetCategory(doc, id);
            if (cat != null)
            {
                // Only add the regular category entry if it has non-DirectShape elements
                bool hasRegularElements = regularElementsByCategory.ContainsKey(id) && 
                                        regularElementsByCategory[id].Count > 0;
                
                if (hasRegularElements)
                {
                    var entry = new Dictionary<string, object>
                    {
                        { "Name", cat.Name },
                        { "Count", regularElementsByCategory[id].Count },
                        { "CategoryId", cat.Id },
                        { "IsDirectShape", false },
                        { "ElementIds", regularElementsByCategory[id] }
                    };
                    categoryList.Add(entry);
                }
                
                // If this category contains Direct Shapes, add a separate entry for them
                if (directShapesByCategory.ContainsKey(id))
                {
                    int directShapeCount = directShapesByCategory[id].Count;
                    string directShapeName = $"Direct Shapes: {cat.Name}";
                    var entry = new Dictionary<string, object>
                    {
                        { "Name", directShapeName },
                        { "Count", directShapeCount },
                        { "CategoryId", cat.Id },
                        { "IsDirectShape", true },
                        { "DirectShapes", directShapesByCategory[id] }
                    };
                    categoryList.Add(entry);
                }
            }
        }
        
        // Add split viewer categories based on VIEW_FAMILY
        foreach (var kvp in viewersByFamily)
        {
            var entry = new Dictionary<string, object>
            {
                { "Name", "Views: " + kvp.Key },
                { "Count", kvp.Value.Count },
                { "CategoryId", new ElementId((int)BuiltInCategory.OST_Viewers) },
                { "IsDirectShape", false },
                { "IsViewerFamily", true },
                { "ViewFamilyName", kvp.Key },
                { "ElementIds", kvp.Value }
            };
            categoryList.Add(entry);
        }
        
        // Sort the list to keep Direct Shapes grouped with their parent categories
        categoryList = categoryList.OrderBy(c => ((string)c["Name"]).Replace("Direct Shapes: ", "")).ToList();
        
        // Define properties to display.
        var propertyNames = new List<string> { "Name", "Count" };
        
        // Show the DataGrid to let the user select one or more categories.
        List<Dictionary<string, object>> selectedCategories = CustomGUIs.DataGrid(categoryList, propertyNames, false);
        if (selectedCategories.Count == 0)
            return Result.Cancelled;
        
        // Gather element IDs for each selected category.
        List<ElementId> elementIds = new List<ElementId>();
        foreach (var selectedCategory in selectedCategories)
        {
            bool isDirectShape = selectedCategory.ContainsKey("IsDirectShape") && (bool)selectedCategory["IsDirectShape"];
            
            if (isDirectShape)
            {
                // Get the Direct Shape IDs
                var directShapes = (List<DirectShape>)selectedCategory["DirectShapes"];
                elementIds.AddRange(directShapes.Select(ds => ds.Id));
            }
            else
            {
                // Get the regular element IDs
                var elementIdList = (List<ElementId>)selectedCategory["ElementIds"];
                elementIds.AddRange(elementIdList);
            }
        }
        
        // Merge with any currently selected elements.
        ICollection<ElementId> currentSelection = uiDoc.GetSelectionIds();
        elementIds.AddRange(currentSelection);
        
        // Update the selection (using Distinct() to remove duplicates).
        uiDoc.SetSelectionIds(elementIds.Distinct().ToList());
        
        return Result.Succeeded;
    }
}
