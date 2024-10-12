using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

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
        List<View> selectedUIViews = CustomGUIs.DataGrid<View>(Views, properties);

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
        string logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "revit-scripts", "LogViewChanges", $"{projectName}");
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
