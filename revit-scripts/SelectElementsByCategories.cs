using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectElementsByCategories : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Document doc = uiapp.ActiveUIDocument.Document;
        UIDocument uiDoc = uiapp.ActiveUIDocument;
        
        // Use a dictionary to track unique categories by ID
        Dictionary<int, Dictionary<string, object>> uniqueCategories = new Dictionary<int, Dictionary<string, object>>();
        
        // Get all built-in categories
        foreach (BuiltInCategory bic in Enum.GetValues(typeof(BuiltInCategory)))
        {
            // Skip invalid categories
            if (bic == BuiltInCategory.INVALID)
                continue;
            
            try
            {
                int categoryId = (int)bic;
                
                // Skip if we already have this category
                if (uniqueCategories.ContainsKey(categoryId))
                    continue;
                
                Category cat = Category.GetCategory(doc, bic);
                if (cat != null)
                {
                    // Only add model categories (not annotation categories or subcategories)
                    if (cat.CategoryType == CategoryType.Model || cat.CategoryType == CategoryType.Annotation)
                    {
                        // Skip subcategories (they have a parent)
                        if (cat.Parent == null)
                        {
                            var entry = new Dictionary<string, object>
                            {
                                { "Name", cat.Name },
                                { "CategoryId", cat.Id }
                            };
                            uniqueCategories[categoryId] = entry;
                        }
                    }
                }
            }
            catch
            {
                // Some built-in categories might throw exceptions, skip them
                continue;
            }
        }
        
        // Also add main categories from the document settings (but not subcategories)
        Categories categories = doc.Settings.Categories;
        foreach (Category cat in categories)
        {
            // Skip subcategories (categories with parents)
            if (cat.Parent != null)
                continue;
            
            // Skip line styles and other non-element categories
            if (cat.CategoryType != CategoryType.Model && cat.CategoryType != CategoryType.Annotation)
                continue;
            
            int categoryId = cat.Id.IntegerValue;
            
            // Add only if not already in our dictionary
            if (!uniqueCategories.ContainsKey(categoryId))
            {
                var entry = new Dictionary<string, object>
                {
                    { "Name", cat.Name },
                    { "CategoryId", cat.Id }
                };
                uniqueCategories[categoryId] = entry;
            }
        }
        
        // Convert dictionary to list and sort alphabetically
        List<Dictionary<string, object>> categoryList = uniqueCategories.Values
            .OrderBy(c => (string)c["Name"])
            .ToList();
        
        // Define properties to display (only "Name" in this case).
        var propertyNames = new List<string> { "Name" };
        
        // Show the DataGrid to let the user select one or more categories.
        List<Dictionary<string, object>> selectedCategories = CustomGUIs.DataGrid(categoryList, propertyNames, false);
        if (selectedCategories.Count == 0)
            return Result.Cancelled;
        
        // Gather element IDs for each selected category.
        List<ElementId> elementIds = new List<ElementId>();
        foreach (var selectedCategory in selectedCategories)
        {
            ElementId catId = (ElementId)selectedCategory["CategoryId"];
            
            try
            {
                // Collect elements of this category from the entire document
                FilteredElementCollector categoryCollector = new FilteredElementCollector(doc);
                categoryCollector.WhereElementIsNotElementType();
                
                // Try using OfCategory with BuiltInCategory if it's a built-in category
                if (catId.IntegerValue < 0) // Built-in categories have negative IDs
                {
                    try
                    {
                        var builtInCat = (BuiltInCategory)catId.IntegerValue;
                        var elementIdsOfCategory = categoryCollector
                            .OfCategory(builtInCat)
                            .Select(e => e.Id)
                            .ToList();
                        elementIds.AddRange(elementIdsOfCategory);
                    }
                    catch
                    {
                        // If OfCategory fails, try OfCategoryId
                        categoryCollector = new FilteredElementCollector(doc);
                        var elementIdsOfCategory = categoryCollector
                            .WhereElementIsNotElementType()
                            .OfCategoryId(catId)
                            .Select(e => e.Id)
                            .ToList();
                        elementIds.AddRange(elementIdsOfCategory);
                    }
                }
                else
                {
                    // For custom categories, use OfCategoryId
                    var elementIdsOfCategory = categoryCollector
                        .OfCategoryId(catId)
                        .Select(e => e.Id)
                        .ToList();
                    elementIds.AddRange(elementIdsOfCategory);
                }
            }
            catch
            {
                // If we can't collect elements for this category, skip it
                continue;
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
