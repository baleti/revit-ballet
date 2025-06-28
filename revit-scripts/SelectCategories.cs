using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

// A simple wrapper class for category information.
public class CategoryWrapper
{
    public ElementId Id { get; set; }
    public string Name { get; set; }
    
    public CategoryWrapper(ElementId id, string name)
    {
        Id = id;
        Name = name;
    }
}

public abstract class SelectCategoriesBase : IExternalCommand
{
    public abstract bool selectInView { get; }
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Document doc = uiapp.ActiveUIDocument.Document;
        View activeView = uiapp.ActiveUIDocument.ActiveView;
        UIDocument uiDoc = uiapp.ActiveUIDocument;
        
        FilteredElementCollector collector;
        if (selectInView)
        {
            // Collect elements visible in the current view.
            collector = new FilteredElementCollector(doc, activeView.Id);
        }
        else
        {
            // Collect elements from the entire document.
            collector = new FilteredElementCollector(doc);
            collector.WhereElementIsNotElementType();
        }
        
        // Build a unique set of category IDs.
        HashSet<ElementId> categoryIds = new HashSet<ElementId>();
        foreach (Element elem in collector)
        {
            Category category = elem.Category;
            if (category != null)
            {
                categoryIds.Add(category.Id);
            }
        }
        
        // --- Additional logic for In-View mode to ensure Views (OST_Viewers) is included ---
        if (selectInView)
        {
            // Collect all elements in the current view that are not in groups,
            // excluding element types, null categories, and cameras.
            List<ElementId> elementsNotInGroups = new FilteredElementCollector(doc, activeView.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.GroupId == ElementId.InvalidElementId)
                .Where(e => e.Category != null)
                .Where(e => e.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Cameras)
                .Select(e => e.Id)
                .ToList();
            
            // Check if any element belongs to the OST_Viewers category.
            bool hasViewers = elementsNotInGroups
                .Select(id => doc.GetElement(id))
                .Any(e => e != null && e.Category != null &&
                          e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Viewers);
            
            if (hasViewers)
            {
                // Add the OST_Viewers category manually.
                ElementId viewersCatId = new ElementId((int)BuiltInCategory.OST_Viewers);
                categoryIds.Add(viewersCatId);
            }
        }
        // --- End additional logic ---
        
        // Build a list of CategoryWrapper objects for the DataGrid.
        List<CategoryWrapper> categoryList = new List<CategoryWrapper>();
        foreach (ElementId id in categoryIds)
        {
            // For OST_Viewers in in-view mode, we will add them later grouped by view family.
            if (id.IntegerValue == (int)BuiltInCategory.OST_Viewers && selectInView)
                continue;
            
            Category cat = Category.GetCategory(doc, id);
            if (cat != null)
            {
                categoryList.Add(new CategoryWrapper(cat.Id, cat.Name));
            }
            else if (id.IntegerValue == (int)BuiltInCategory.OST_Viewers)
            {
                categoryList.Add(new CategoryWrapper(id, "Views (OST_Viewers)"));
            }
        }
        
        // In in-view mode, add split viewer categories based on the VIEW_FAMILY built-in parameter.
        if (selectInView)
        {
            // Reuse the working elementsNotInGroups block.
            List<ElementId> elementsNotInGroups = new FilteredElementCollector(doc, activeView.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.GroupId == ElementId.InvalidElementId)
                .Where(e => e.Category != null)
                .Where(e => e.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Cameras)
                .Select(e => e.Id)
                .ToList();
            
            var viewerElements = elementsNotInGroups
                .Select(id => doc.GetElement(id))
                .Where(e => e != null && e.Category != null &&
                            e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Viewers)
                .ToList();
            
            // Group the viewer elements by their VIEW_FAMILY parameter.
            var grouped = viewerElements
                .GroupBy(e => e.get_Parameter(BuiltInParameter.VIEW_FAMILY)?.AsValueString() ?? "Unknown")
                .Select(g => g.Key)
                .ToList();
            
            foreach (string familyName in grouped)
            {
                // Create a separate entry for each group using the view family parameter value.
                categoryList.Add(new CategoryWrapper(new ElementId((int)BuiltInCategory.OST_Viewers), "Views: " + familyName));
            }
        }
        
        // Define properties to display (only "Name" in this example).
        var propertyNames = new List<string> { "Name" };
        
        // Show the DataGrid to let the user select one or more categories.
        List<CategoryWrapper> selectedCategories = CustomGUIs.DataGrid<CategoryWrapper>(categoryList, propertyNames);
        if (selectedCategories.Count == 0)
            return Result.Cancelled;
        
        // Gather element IDs for each selected category.
        List<ElementId> elementIds = new List<ElementId>();
        foreach (CategoryWrapper selectedCategory in selectedCategories)
        {
            FilteredElementCollector categoryCollector;
            if (selectInView)
            {
                categoryCollector = new FilteredElementCollector(doc, activeView.Id);
            }
            else
            {
                categoryCollector = new FilteredElementCollector(doc);
                categoryCollector.WhereElementIsNotElementType();
            }
            
            // If the selected category is a split Views category, filter further by the VIEW_FAMILY parameter.
            if (selectedCategory.Id.IntegerValue == (int)BuiltInCategory.OST_Viewers &&
                selectedCategory.Name.StartsWith("Views: "))
            {
                string familyName = selectedCategory.Name.Substring("Views: ".Length);
                var viewerIds = categoryCollector
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Viewers)
                    .Where(e => (e.get_Parameter(BuiltInParameter.VIEW_FAMILY)?.AsValueString() ?? "Unknown") == familyName)
                    .Select(e => e.Id);
                elementIds.AddRange(viewerIds);
            }
            else
            {
                var elementIdsOfCategory = categoryCollector
                    .WhereElementIsNotElementType()
                    .OfCategory((BuiltInCategory)selectedCategory.Id.IntegerValue)
                    .Select(e => e.Id)
                    .ToList();
                elementIds.AddRange(elementIdsOfCategory);
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

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectCategoriesInView : SelectCategoriesBase
{
    public override bool selectInView => true;
}

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectCategoriesInProject : SelectCategoriesBase
{
    public override bool selectInView => false;
}
