using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CloseViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get SessionId for database operations
        string sessionId = RevitBallet.LogViewChanges.GetSessionId();

        IList<UIView> UIViews = uidoc.GetOpenUIViews();
        List<View> Views = new List<View>();
        foreach (UIView UIview in UIViews)
        {
            View view = doc.GetElement(UIview.ViewId) as View;
            Views.Add(view);
        }

        List<string> properties = new List<string> { "Title", "ViewType" };
        var viewDicts = CustomGUIs.ConvertToDataGridFormat(Views, properties);
        var selectedDicts = CustomGUIs.DataGrid(viewDicts, properties, false);
        List<View> selectedUIViews = CustomGUIs.ExtractOriginalObjects<View>(selectedDicts);

        if (selectedUIViews.Count == 0)
            return Result.Failed;

        foreach (View view in selectedUIViews)
        {
            foreach (UIView openedUIView in uidoc.GetOpenUIViews())
            {
                if (openedUIView.ViewId.Equals(view.Id))
                {
                    // Remove the closed view from history
                    try
                    {
                        LogViewChangesDatabase.RemoveViewFromHistory(sessionId, doc.Title, view.Title);
                    }
                    catch
                    {
                        // Silently fail - don't interrupt the close operation
                    }

                    openedUIView.Close();
                }
            }
        }

        return Result.Succeeded;
    }
}
