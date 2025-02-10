using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenViews : IExternalCommand
{
    public class ViewInfo
    {
        public string Title { get; set; }
        public string SheetFolder { get; set; }
        public View RevitView { get; set; }
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;

        List<View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => !v.IsTemplate && v.Title != "Project Browser" && v.Title != "System Browser")
            .OrderBy(v => v.Title)
            .ToList();

        List<ViewInfo> viewInfoList = allViews.Select(v =>
        {
            Parameter sheetFolderParam = v.LookupParameter("Sheet Folder");
            string sheetFolderValue = sheetFolderParam?.AsString() ?? string.Empty;
            return new ViewInfo
            {
                Title = v.Title,
                SheetFolder = sheetFolderValue,
                RevitView = v
            };
        }).ToList();

        int selectedIndex = -1;
        if (activeView is ViewSheet)
        {
            selectedIndex = viewInfoList.FindIndex(v => v.RevitView.Id == activeView.Id);
        }
        else
        {
            var viewports = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .Where(vp => vp.ViewId == activeView.Id)
                .ToList();

            if (viewports.Any())
            {
                ViewSheet containingSheet = doc.GetElement(viewports.First().SheetId) as ViewSheet;
                if (containingSheet != null)
                {
                    selectedIndex = viewInfoList.FindIndex(v => v.RevitView.Id == containingSheet.Id);
                }
            }
            else
            {
                selectedIndex = viewInfoList.FindIndex(v => v.RevitView.Id == activeView.Id);
            }
        }

        List<int> initialSelectionIndices = selectedIndex >= 0 ? new List<int> { selectedIndex } : new List<int>();
        List<string> properties = new List<string> { "Title", "SheetFolder" };
        
        var selectedItems = CustomGUIs.DataGrid<ViewInfo>(viewInfoList, properties, initialSelectionIndices);
        selectedItems.ForEach(vInfo => uidoc.RequestViewChange(vInfo.RevitView));
        
        return Result.Succeeded;
    }
}
