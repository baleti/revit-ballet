using System;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CloseCurrentView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get SessionId for database operations
        string sessionId = RevitBallet.LogViewChanges.GetSessionId();

        // Remove the current view from history
        string activeViewTitle = uidoc.ActiveView.Title;
        try
        {
            LogViewChangesDatabase.RemoveViewFromHistory(sessionId, doc.Title, activeViewTitle);
        }
        catch
        {
            // Silently fail - don't interrupt the close operation
        }

        UIView activeUIView = uidoc.GetOpenUIViews().FirstOrDefault(u => u.ViewId == uidoc.ActiveView.Id);
        activeUIView.Close();

        return Result.Succeeded;
    }
}
