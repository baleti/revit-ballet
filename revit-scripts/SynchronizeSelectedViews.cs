using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SynchronizeSelectedViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        Document doc = commandData.Application.ActiveUIDocument.Document;
        UIDocument uidoc = commandData.Application.ActiveUIDocument;

        Autodesk.Revit.DB.View activeView = doc.ActiveView;
        UIView activeuiView = uidoc?.GetOpenUIViews()?.FirstOrDefault(item => item.ViewId == uidoc.ActiveView.Id);

        // Get the corners of the current UIView
        IList<XYZ> corners = activeuiView.GetZoomCorners();
        XYZ topLeft = corners[0];
        XYZ bottomRight = corners[1];

        ListBox resultListBox = new ListBox();
        // Populate ListBox with the names of the opened UIViews
        foreach (var uiView in uidoc.GetOpenUIViews())
        {
            ElementId viewId = uiView.ViewId;
            var viewElement = doc.GetElement(viewId) as Autodesk.Revit.DB.View;
            resultListBox.Items.Add(viewElement.Name);
        }


        // Loop through all other open UI views and move them to the same position
        foreach (UIView openedView in commandData.Application.ActiveUIDocument.GetOpenUIViews())
        {
            if (openedView.ViewId != activeView.Id)
            {
                if (!(doc.GetElement(openedView.ViewId) is ViewSheet))
                {
                    //openedView.ZoomAndCenterRectangle(topLeft, bottomRight);
                }
            }
        }

        return Result.Succeeded;
    }
    private void FilterListBoxItems(ListBox listBox, string searchText)
    {
        List<string> filteredItems = new List<string>();
        foreach (string item in listBox.Items)
        {
            if (item.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                filteredItems.Add(item);
            }
        }
        listBox.BeginUpdate();
        listBox.Items.Clear();
        listBox.Items.AddRange(filteredItems.ToArray());
        listBox.EndUpdate();
    }
}
