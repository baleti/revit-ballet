using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class Select3DSectionBoxInCurrentView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;
        
        // Check if the current view is a 3D view
        if (!(activeView is View3D view3D))
        {
            message = "Current view is not a 3D view.";
            return Result.Failed;
        }
        
        // Check if section box is active
        if (!view3D.IsSectionBoxActive)
        {
            message = "Section box is not active in the current 3D view. Enable it first via View Properties.";
            return Result.Failed;
        }
        
        // Get the current view's name
        string viewName = view3D.Name;
        
        // Collect all section boxes in the document
        FilteredElementCollector collector = new FilteredElementCollector(doc)
            .WhereElementIsNotElementType()
            .OfCategory(BuiltInCategory.OST_SectionBox);
        
        List<ElementId> sectionBoxIds = new List<ElementId>();
        
        foreach (Element elem in collector)
        {
            // Get the workset parameter
            Parameter worksetParam = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
            
            string worksetValue = worksetParam.AsString();
            
            // Check if the workset value matches the expected pattern for this view
            // Pattern is: View "3D View: [ViewName]"
            string expectedWorksetValue1 = $"View \"3D View: {viewName}\"";
            string expectedWorksetValue2 = $"View \"{viewName}\""; // Sometimes it might be without "3D View: "
            
            if (worksetValue == expectedWorksetValue1 || worksetValue == expectedWorksetValue2)
            {
                sectionBoxIds.Add(elem.Id);
            }
            
            // Alternative: Check if workset contains the view name (less strict matching)
            else if (worksetValue.Contains(viewName))
            {
                // Debug info - uncomment to see what workset values look like
                // TaskDialog.Show("Debug", $"Found potential match:\nView Name: {viewName}\nWorkset: {worksetValue}");
                sectionBoxIds.Add(elem.Id);
            }
        }
        
        if (sectionBoxIds.Count == 0)
        {
            // Try alternative approach: check view-specific elements in current view
            FilteredElementCollector viewCollector = new FilteredElementCollector(doc, view3D.Id)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_SectionBox);
            
            sectionBoxIds = viewCollector.ToElementIds().ToList();
            
            if (sectionBoxIds.Count == 0)
            {
                message = "No section box element found for the current 3D view.";
                return Result.Failed;
            }
        }

        // Clear current selection and select the section box
        uidoc.SetSelectionIds(sectionBoxIds);
        
        if (sectionBoxIds.Count == 1)
        {
            TaskDialog.Show("Success", "Selected the section box in the current view.");
        }
        else if (sectionBoxIds.Count > 1)
        {
            TaskDialog.Show("Warning", $"Found {sectionBoxIds.Count} section boxes for this view. All have been selected.");
        }
        
        return Result.Succeeded;
    }
}
