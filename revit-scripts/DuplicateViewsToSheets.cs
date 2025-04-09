using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class DuplicateViewsToSheets : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            // Dictionary mapping the original view's ElementId to the viewport box center (sheet coordinates).
            Dictionary<ElementId, XYZ> originalViewLocations = new Dictionary<ElementId, XYZ>();

            // Step 1. Retrieve all views that are placed on sheets.
            List<View> placedViews = new List<View>();
            FilteredElementCollector vpCollector = new FilteredElementCollector(doc).OfClass(typeof(Viewport));
            foreach (Viewport vp in vpCollector)
            {
                if (vp.SheetId != ElementId.InvalidElementId)
                {
                    View view = doc.GetElement(vp.ViewId) as View;
                    if (view != null && !placedViews.Contains(view))
                    {
                        placedViews.Add(view);
                        
                        // Store the viewport's box center coordinates (sheet coordinates)
                        XYZ boxCenter = vp.GetBoxCenter();
                        originalViewLocations[view.Id] = boxCenter;
                    }
                }
            }
            if (placedViews.Count == 0)
            {
                TaskDialog.Show("Duplicate Views To Sheets", "No views placed on sheets were found.");
                return Result.Cancelled;
            }

            // Build DataGrid entries for views. Now include a combined "Sheet" column and a "Sheet Folder" column.
            // Order: Title, Sheet (combined as SheetNumber - SheetName), Sheet Folder, then Id.
            List<Dictionary<string, object>> viewEntries = new List<Dictionary<string, object>>();
            foreach (View v in placedViews)
            {
                // Retrieve the first associated viewport for this view.
                Viewport vp = new FilteredElementCollector(doc).OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .FirstOrDefault(x => x.ViewId == v.Id && x.SheetId != ElementId.InvalidElementId);
                
                string combinedSheet = "";
                string sheetFolder = "";
                if (vp != null)
                {
                    ViewSheet sheet = doc.GetElement(vp.SheetId) as ViewSheet;
                    if (sheet != null)
                    {
                        // Combine sheet number and name into one string.
                        combinedSheet = sheet.SheetNumber + " - " + sheet.Name;
                        // Look up the sheet folder from the sheet.
                        Parameter sheetFolderParam = sheet.LookupParameter("Sheet Folder");
                        sheetFolder = sheetFolderParam?.AsString() ?? string.Empty;
                    }
                }
                
                viewEntries.Add(new Dictionary<string, object>
                {
                    { "Sheet", combinedSheet },
                    { "View Title", v.Name },
                    { "Sheet Folder", sheetFolder },
                    { "Id", v.Id.IntegerValue }
                });
            }

            // Define the order of columns for the first prompt.
            List<string> viewPropertyNames = new List<string> { "Sheet", "View Title", "Sheet Folder", "Id" };
            List<Dictionary<string, object>> selectedViewEntries = CustomGUIs.DataGrid(viewEntries, viewPropertyNames, spanAllScreens: false);
            if (selectedViewEntries == null || selectedViewEntries.Count == 0)
            {
                TaskDialog.Show("Duplicate Views To Sheets", "No views were selected.");
                return Result.Cancelled;
            }
            // Build a set of selected view IDs.
            HashSet<ElementId> selectedViewIds = new HashSet<ElementId>();
            foreach (var entry in selectedViewEntries)
            {
                if (entry.ContainsKey("Id"))
                {
                    int idValue = Convert.ToInt32(entry["Id"]);
                    selectedViewIds.Add(new ElementId(idValue));
                }
            }
            List<View> selectedViews = placedViews.Where(v => selectedViewIds.Contains(v.Id)).ToList();

            // Step 2. Prompt the user to select target sheets.
            // Build DataGrid entries for sheets with an extra column "Sheet Number" at the beginning.
            List<ViewSheet> allSheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();
            if (allSheets.Count == 0)
            {
                TaskDialog.Show("Duplicate Views To Sheets", "No sheets were found in the project.");
                return Result.Cancelled;
            }
            
            List<Dictionary<string, object>> sheetEntries = new List<Dictionary<string, object>>();
            foreach (ViewSheet sheet in allSheets)
            {
                string sheetFolder = "";
                Parameter sheetFolderParam = sheet.LookupParameter("Sheet Folder");
                sheetFolder = sheetFolderParam?.AsString() ?? string.Empty;
                
                sheetEntries.Add(new Dictionary<string, object>
                {
                    { "Sheet Number", sheet.SheetNumber },
                    { "Title", sheet.Name },
                    { "Sheet Folder", sheetFolder },
                    { "Id", sheet.Id.IntegerValue }
                });
            }
            List<string> sheetPropertyNames = new List<string> { "Sheet Number", "Title", "Sheet Folder", "Id" };
            List<Dictionary<string, object>> selectedSheetEntries = CustomGUIs.DataGrid(sheetEntries, sheetPropertyNames, spanAllScreens: false);
            if (selectedSheetEntries == null || selectedSheetEntries.Count == 0)
            {
                TaskDialog.Show("Duplicate Views To Sheets", "No sheets were selected.");
                return Result.Cancelled;
            }
            // Build a set of selected sheet IDs.
            HashSet<ElementId> selectedSheetIds = new HashSet<ElementId>();
            foreach (var entry in selectedSheetEntries)
            {
                if (entry.ContainsKey("Id"))
                {
                    int idValue = Convert.ToInt32(entry["Id"]);
                    selectedSheetIds.Add(new ElementId(idValue));
                }
            }
            List<ViewSheet> targetSheets = allSheets.Where(s => selectedSheetIds.Contains(s.Id)).ToList();

            // Step 3. Duplicate each selected view (with detailing) and place it on every selected target sheet.
            using (Transaction tx = new Transaction(doc, "Duplicate Views To Sheets"))
            {
                tx.Start();

                foreach (View origView in selectedViews)
                {
                    // Retrieve the original viewport box center position (in sheet coordinates)
                    XYZ origBoxCenter = originalViewLocations.ContainsKey(origView.Id) ?
                                        originalViewLocations[origView.Id] : new XYZ(0, 0, 0);

                    foreach (ViewSheet targetSheet in targetSheets)
                    {
                        // Duplicate the view using WithDetailing.
                        ElementId dupViewId = origView.Duplicate(ViewDuplicateOption.WithDetailing);
                        View dupView = doc.GetElement(dupViewId) as View;
                        if (dupView == null)
                            continue;

                        // Generate a unique view name for the duplicate.
                        string newName = GetUniqueViewName(doc, origView.Name);
                        dupView.Name = newName;

                        // Create the viewport
                        Viewport newViewport = Viewport.Create(doc, targetSheet.Id, dupView.Id, origBoxCenter);

                        // Optional: Update parameters on the duplicate.
                        Parameter viewTitleParam = dupView.LookupParameter("View Title");
                        if (viewTitleParam != null && !viewTitleParam.IsReadOnly)
                        {
                            viewTitleParam.Set(origView.Name + " - " + targetSheet.Name);
                        }
                        Parameter dupSheetFolderParam = dupView.LookupParameter("Sheet Folder");
                        if (dupSheetFolderParam != null && !dupSheetFolderParam.IsReadOnly)
                        {
                            // Get the target sheet's folder information.
                            Parameter targetSheetFolderParam = targetSheet.LookupParameter("Sheet Folder");
                            string targetFolder = targetSheetFolderParam?.AsString() ?? "";
                            dupSheetFolderParam.Set(targetFolder);
                        }
                    }
                }
                tx.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Helper method to generate a unique view name by appending a numeric suffix.
        /// </summary>
        private string GetUniqueViewName(Document doc, string baseName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(View));
            HashSet<string> existingNames = new HashSet<string>(collector.Cast<View>().Select(v => v.Name));
            string newName = baseName;
            int suffix = 2;
            while (existingNames.Contains(newName))
            {
                newName = baseName + " " + suffix;
                suffix++;
            }
            return newName;
        }
    }
}
