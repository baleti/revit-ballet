using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.ReadOnly)]
public class ShowViewRange : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        try
        {
            // Get the current document and active view
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            if (activeView == null)
            {
                TaskDialog.Show("Error", "No active view found.");
                return Result.Failed;
            }

            // Debug information about the view
            string viewInfo = $"View Name: {activeView.Name}\n" +
                            $"View Type: {activeView.ViewType}\n";

            // Get the View Template assigned to the active view
            ElementId templateId = activeView.ViewTemplateId;
            
            // Check if view template exists
            if (templateId != null && templateId != ElementId.InvalidElementId)
            {
                View viewTemplate = doc.GetElement(templateId) as View;
                if (viewTemplate != null)
                {
                    // Get view range information
                    string viewRangeInfo = GetViewRangeInfo(viewTemplate);
                    
                    // Display complete information
                    TaskDialog.Show("View Template Information", 
                        viewInfo + 
                        $"Template Name: {viewTemplate.Name}\n" +
                        $"Template Id: {templateId}\n" +
                        "View Range Parameters:\n" +
                        viewRangeInfo);
                }
                else
                {
                    TaskDialog.Show("Error", 
                        viewInfo + 
                        $"Template ID exists ({templateId}) but failed to retrieve template.");
                }
            }
            else
            {
                // Check if view has a non-locked template
                Parameter viewTemplateParam = activeView.get_Parameter(BuiltInParameter.VIEW_TEMPLATE);
                string templateName = viewTemplateParam?.AsValueString() ?? "None";

                TaskDialog.Show("Template Status", 
                    viewInfo +
                    $"View Template Applied: {templateName}\n" +
                    "Note: The template might be applied but not locked.\n" +
                    "To get view range information, the template needs to be locked.");
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"An error occurred:\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}");
            return Result.Failed;
        }
    }

    private string GetViewRangeInfo(View viewTemplate)
    {
        string GetParamValue(BuiltInParameter param)
        {
            Parameter p = viewTemplate.get_Parameter(param);
            return p != null && p.HasValue ? p.AsValueString() : "Not Set";
        }

        return string.Join("\n", new[]
        {
            $"Top: {GetParamValue(BuiltInParameter.PLAN_VIEW_CUT_PLANE_HEIGHT)}",
            $"Cut Plane: {GetParamValue(BuiltInParameter.PLAN_VIEW_RANGE)}",
            $"Bottom: {GetParamValue(BuiltInParameter.PLAN_VIEW_LEVEL)}"
        });
    }
}
