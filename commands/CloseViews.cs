using System;
using System.Collections.Generic;
using System.IO;
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
        string projectName = doc != null ? doc.Title : "UnknownProject";

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
                    RemoveViewFromLog(view.Title, projectName); // Remove the closed view from the log file
                    openedUIView.Close();
                }
            }

        }
        return Result.Succeeded;
    }
    private void RemoveViewFromLog(string viewTitle, string projectName)
    {
        string logFilePath = PathHelper.GetLogViewChangesPath(projectName);
        if (File.Exists(logFilePath))
        {
            List<string> logEntries = File.ReadAllLines(logFilePath).ToList();
            logEntries = logEntries
                .Where(entry => !entry.Contains($" {viewTitle}"))
                .ToList();
            File.WriteAllLines(logFilePath, logEntries);
        }
    }
}
