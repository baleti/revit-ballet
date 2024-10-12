using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.ComponentModel;

[Transaction(TransactionMode.Manual)]
public class TagElementsInSelectedViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the selected elements
        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
        if (selectedIds.Count == 0)
        {
            return Result.Failed;
        }

        List<Autodesk.Revit.DB.View> selectedViews = new List<Autodesk.Revit.DB.View>();
        foreach (ElementId id in selectedIds)
        {
            Element element = doc.GetElement(id);
            if (element is Viewport viewport)
            {
                // Get the view associated with the viewport
                ElementId viewId = viewport.ViewId;
                Autodesk.Revit.DB.View view = doc.GetElement(viewId) as Autodesk.Revit.DB.View;
                if (view != null && !(view is ViewSchedule) && view.ViewType != ViewType.DrawingSheet)
                {
                    selectedViews.Add(view);
                }
            }
            else if (element is Autodesk.Revit.DB.View view && !(view is ViewSchedule) && view.ViewType != ViewType.DrawingSheet)
            {
                selectedViews.Add(view);
            }
        }

        if (selectedViews.Count == 0)
        {
            return Result.Failed;
        }

        // Gather all family types present in the selected views
        Dictionary<ElementId, (string CategoryName, string FamilyName, string TypeName)> familyTypes = new Dictionary<ElementId, (string, string, string)>();
        foreach (Autodesk.Revit.DB.View view in selectedViews)
        {
            var elementsToTag = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && e.CanBeTagged(view))
                .ToList();

            foreach (Element element in elementsToTag)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId && !familyTypes.ContainsKey(typeId))
                {
                    ElementType elementType = doc.GetElement(typeId) as ElementType;
                    if (elementType != null)
                    {
                        familyTypes[typeId] = (element.Category.Name, elementType.FamilyName, elementType.Name);
                    }
                }
            }
        }

        // Prompt user to choose family types to tag
        List<Dictionary<string, object>> entries = familyTypes
            .Select(kv => new Dictionary<string, object> { { "Category", kv.Value.CategoryName }, { "Family", kv.Value.FamilyName }, { "Type", kv.Value.TypeName } })
            .ToList();
        List<string> propertyNames = new List<string> { "Category", "Family", "Type" };
        List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);
        if (selectedEntries.Count == 0)
        {
            TaskDialog.Show("Error", "No family types selected to tag.");
            return Result.Failed;
        }

        HashSet<ElementId> selectedTypeIds = new HashSet<ElementId>(
            selectedEntries.Select(e =>
                familyTypes.FirstOrDefault(kv => kv.Value.CategoryName == e["Category"].ToString() && kv.Value.FamilyName == e["Family"].ToString() && kv.Value.TypeName == e["Type"].ToString()).Key));

        using (Transaction trans = new Transaction(doc, "Place Tags"))
        {
            trans.Start();
            try
            {
                int tagNumber = 1;

                foreach (Autodesk.Revit.DB.View view in selectedViews)
                {
                    var elementsToTag = new FilteredElementCollector(doc, view.Id)
                        .WhereElementIsNotElementType()
                        .Where(e => e.Category != null && e.CanBeTagged(view) && selectedTypeIds.Contains(e.GetTypeId()))
                        .ToList();

                    foreach (Element element in elementsToTag)
                    {
                        LocationPoint location = element.Location as LocationPoint;
                        XYZ originalPosition = location?.Point;
                        BoundingBoxXYZ boundingBox = element.get_BoundingBox(null);

                        if (boundingBox != null && originalPosition != null)
                        {
                            // Get the current view scale
                            Autodesk.Revit.DB.View activeView = view;
                            int viewScale = activeView.Scale;

                            // Calculate the offset based on the view scale
                            double offsetY = 0.009 * viewScale; // Arbitrary scaling factor

                            // Calculate the tag position with offset
                            XYZ minPoint = boundingBox.Min;
                            XYZ tagPosition = new XYZ(originalPosition.X, minPoint.Y - offsetY, originalPosition.Z);

                            IndependentTag newTag = IndependentTag.Create(doc, activeView.Id, new Reference(element), false, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, tagPosition);
                            Parameter typeMarkParam = newTag.LookupParameter("Type Mark");

                            if (typeMarkParam != null && typeMarkParam.IsReadOnly == false)
                            {
                                typeMarkParam.Set(tagNumber.ToString());
                            }

                            tagNumber++;
                        }
                    }
                }

                trans.Commit();
            }
            catch (Exception ex)
            {
                message = ex.Message;
                trans.RollBack();
                return Result.Failed;
            }
        }

        return Result.Succeeded;
    }
}

// Extension method to check if an element can be tagged in a specific view
public static class ElementExtensions
{
    public static bool CanBeTagged(this Element element, Autodesk.Revit.DB.View view)
    {
        // Check if the element can be tagged in the given view
        // You can expand this logic as needed for your specific requirements
        return element.Category != null &&
               element.Category.HasMaterialQuantities &&
               element.Category.CanAddSubcategory &&
               !element.Category.IsTagCategory;
    }
}
