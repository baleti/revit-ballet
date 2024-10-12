using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class TagElementsInViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Gather all views in the project, excluding schedules and sheets
        List<Autodesk.Revit.DB.View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(Autodesk.Revit.DB.View))
            .Cast<Autodesk.Revit.DB.View>()
            .Where(v => !(v is ViewSchedule) && v.ViewType != ViewType.DrawingSheet)
            .ToList();

        // Prepare entries for the CustomGUI
        List<Dictionary<string, object>> viewEntries = allViews
            .Select(v => new Dictionary<string, object> { { "Title", v.Title } })
            .ToList();
        List<string> viewPropertyNames = new List<string> { "Title" };

        // Prompt user to choose views
        List<Dictionary<string, object>> selectedViewEntries = CustomGUIs.DataGrid(viewEntries, viewPropertyNames, spanAllScreens: false);
        if (selectedViewEntries.Count == 0)
        {
            return Result.Failed;
        }

        List<Autodesk.Revit.DB.View> selectedViews = new List<Autodesk.Revit.DB.View>();
        foreach (var entry in selectedViewEntries)
        {
            string viewTitle = entry["Title"].ToString();
            Autodesk.Revit.DB.View view = allViews.FirstOrDefault(v => v.Title == viewTitle);
            if (view != null)
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
