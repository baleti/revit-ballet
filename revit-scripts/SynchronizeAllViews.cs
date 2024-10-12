using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SynchronizeAllViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        Document doc = commandData.Application.ActiveUIDocument.Document;

        View activeView = doc.ActiveView;

        UIView uiView = GetUIView(commandData.Application.ActiveUIDocument, activeView);

        // Get the corners of the current UIView
        IList<XYZ> corners = uiView.GetZoomCorners();
        XYZ topLeft = corners[0];
        XYZ bottomRight = corners[1];

        // Loop through all other open UI views and move them to the same position
        foreach (UIView otherUIView in commandData.Application.ActiveUIDocument.GetOpenUIViews())
        {
            if (otherUIView.ViewId != activeView.Id)
            {
                if (!(doc.GetElement(otherUIView.ViewId) is ViewSheet))
                {
                otherUIView.ZoomAndCenterRectangle(topLeft, bottomRight);
                }
            }
        }

        return Result.Succeeded;
    }

    private UIView GetUIView(UIDocument uidoc, View view)
    {
        Document doc = uidoc.Document;
        UIView uiView = null;

        // Find the UIView for the given view
        foreach (UIView uiview in uidoc.GetOpenUIViews())
        {
            if (uiview.ViewId == view.Id)
            {
                uiView = uiview;
                break;
            }
        }

        return uiView;
    }
}
