using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

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

        if(selectInView)
        {
            // Filter elements visible in the current view
            collector = new FilteredElementCollector(doc, activeView.Id);
        }
        else
        {
            // Filter elements visible in the entire doc
            collector = new FilteredElementCollector(doc);
            // Apply a generic filter, for example, excluding element types
            collector.WhereElementIsNotElementType();
        }

        // Create a set to store unique category IDs
        HashSet<ElementId> categoryIds = new HashSet<ElementId>();

        // Iterate through collected elements and add their category IDs to the set
        foreach (Element elem in collector)
        {
            Category category = elem.Category;
            if (category != null)
            {
                categoryIds.Add(category.Id);  // Add using category ID for uniqueness
            }
        }

        // Retrieve Category objects from their IDs
        List<Category> categoryList = categoryIds
            .Select(id => Category.GetCategory(doc, id))
            .Where(c => c != null)
            .ToList();

        // Define properties to display in the DataGrid
        var propertyNames = new List<string> { "Name" };

        // Get user-selected categories from the DataGrid
        List<Category> selectedCategories = CustomGUIs.DataGrid<Category>(categoryList, propertyNames);

        if (selectedCategories.Count == 0)
            return Result.Cancelled;

        List<ElementId> elementIds = new List<ElementId>();
        foreach (Category selectedCategory in selectedCategories)
        {
            // Note: Changed second 'collector' to a new variable to avoid variable redeclaration.
            FilteredElementCollector categoryCollector;
            if(selectInView)
            {
                // Filter elements visible in the current view
                categoryCollector = new FilteredElementCollector(doc, activeView.Id);
            }
            else
            {
                // Filter elements visible in the entire doc
                categoryCollector = new FilteredElementCollector(doc);
                // Apply a generic filter, for example, excluding element types
                categoryCollector.WhereElementIsNotElementType();
            }
            var elementIdsOfCategory = categoryCollector
                .WhereElementIsNotElementType()
                .OfCategory((BuiltInCategory)selectedCategory.Id.IntegerValue)
                .ToElementIds();
            
            elementIds.AddRange(elementIdsOfCategory);
        }

        // Retrieve the currently selected elements
        ICollection<ElementId> currentSelection = uiDoc.Selection.GetElementIds();

        // Add new elements to the current selection
        elementIds.AddRange(currentSelection);

        // Set the updated selection (without duplicates)
        uiDoc.Selection.SetElementIds(elementIds.Distinct().ToList());

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
