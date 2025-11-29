using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectByCategoriesInProject : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Document doc = uiapp.ActiveUIDocument.Document;
        UIDocument uiDoc = uiapp.ActiveUIDocument;
        
        // Collect elements from the entire document.
        FilteredElementCollector collector = new FilteredElementCollector(doc);
        collector.WhereElementIsNotElementType();

        // Build a unique set of category IDs and track Direct Shapes and regular elements separately.
        HashSet<ElementId> categoryIds = new HashSet<ElementId>();
        Dictionary<ElementId, List<DirectShape>> directShapesByCategory = new Dictionary<ElementId, List<DirectShape>>();
        Dictionary<ElementId, List<ElementId>> regularElementsByCategory = new Dictionary<ElementId, List<ElementId>>();

        // Separate tracking for views and view templates
        List<ElementId> viewIds = new List<ElementId>();
        List<ElementId> viewTemplateIds = new List<ElementId>();

        foreach (Element elem in collector)
        {
            // Special handling for ALL views to separate from view templates (regardless of category)
            if (elem is View view)
            {
                if (view.IsTemplate)
                {
                    viewTemplateIds.Add(elem.Id);
                }
                else
                {
                    viewIds.Add(elem.Id);
                }
                // Don't process views as regular category elements
                continue;
            }

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
                    // Store regular (non-DirectShape, non-View) element IDs
                    if (!regularElementsByCategory.ContainsKey(category.Id))
                    {
                        regularElementsByCategory[category.Id] = new List<ElementId>();
                    }
                    regularElementsByCategory[category.Id].Add(elem.Id);
                }
            }
        }
        
        // Build a list of dictionaries for the DataGrid.
        List<Dictionary<string, object>> categoryList = new List<Dictionary<string, object>>();

        foreach (ElementId id in categoryIds)
        {
            // Skip OST_Viewers - we'll handle it separately below
            if (id.AsLong() == (int)BuiltInCategory.OST_Viewers)
            {
                // Add direct shapes for OST_Viewers if any
                if (directShapesByCategory.ContainsKey(id))
                {
                    int directShapeCount = directShapesByCategory[id].Count;
                    string directShapeName = "Direct Shapes: Views";
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
                continue;
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

        // Add separate entries for Views and View Templates
        ElementId viewsCategoryId = ((long)BuiltInCategory.OST_Viewers).ToElementId();

        if (viewIds.Count > 0)
        {
            var viewsEntry = new Dictionary<string, object>
            {
                { "Name", "Views" },
                { "Count", viewIds.Count },
                { "CategoryId", viewsCategoryId },
                { "IsDirectShape", false },
                { "IsViewTemplate", false },
                { "ElementIds", viewIds }
            };
            categoryList.Add(viewsEntry);
        }

        if (viewTemplateIds.Count > 0)
        {
            var viewTemplatesEntry = new Dictionary<string, object>
            {
                { "Name", "View Templates" },
                { "Count", viewTemplateIds.Count },
                { "CategoryId", viewsCategoryId },
                { "IsDirectShape", false },
                { "IsViewTemplate", true },
                { "ElementIds", viewTemplateIds }
            };
            categoryList.Add(viewTemplatesEntry);
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
            bool isDirectShape = (bool)selectedCategory["IsDirectShape"];
            
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
