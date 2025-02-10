#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
#endregion

namespace MyRevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class CreateCalloutCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData, 
            ref string message, 
            ElementSet elements)
        {
            // Get the active document and view.
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null)
            {
                message = "No active document found.";
                return Result.Failed;
            }
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            // Collect available views (non-template and not the active view)
            IList<View> availableViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.Id != activeView.Id)
                .ToList();

            if (availableViews.Count == 0)
            {
                TaskDialog.Show("Reference Callout", "No available views found to reference.");
                return Result.Cancelled;
            }

            // Prepare list of views for the custom data grid UI.
            List<Dictionary<string, object>> viewEntries = new List<Dictionary<string, object>>();
            foreach (View view in availableViews)
            {
                var entry = new Dictionary<string, object>
                {
                    { "Id", view.Id.IntegerValue },
                    { "Name", view.Name }
                };
                viewEntries.Add(entry);
            }

            List<string> propertyNames = new List<string> { "Id", "Name" };

            // Display the UI (provided by your CustomGUIs.DataGrid method) to let the user choose a view.
            List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(viewEntries, propertyNames, false);
            if (selectedEntries == null || selectedEntries.Count == 0)
            {
                TaskDialog.Show("Reference Callout", "No view was selected.");
                return Result.Cancelled;
            }

            int selectedViewIdInt = Convert.ToInt32(selectedEntries[0]["Id"]);
            ElementId selectedViewId = new ElementId(selectedViewIdInt);

            // Get the selected view element.
            View selectedView = doc.GetElement(selectedViewId) as View;
            if (selectedView == null)
            {
                TaskDialog.Show("Reference Callout", "The selected element is not a view.");
                return Result.Failed;
            }

            // --- Validate the view type combination ---
            // In Revit 2024 the following names are used:
            //   DraftingView  (instead of Drafting)
            //   StructurePlan (instead of StructuralPlan)
            bool validCombination = false;
            if (selectedView.ViewType == ViewType.DraftingView)
            {
                validCombination = true;
            }
            // If the referenced view is a FloorPlan or CeilingPlan,
            // then the parent view must be one of these.
            else if (selectedView.ViewType == ViewType.FloorPlan ||
                     selectedView.ViewType == ViewType.CeilingPlan)
            {
                if (activeView.ViewType == ViewType.FloorPlan ||
                    activeView.ViewType == ViewType.CeilingPlan)
                {
                    validCombination = true;
                }
            }
            // For Section views: allowed if parent is a Section or DraftingView.
            else if (selectedView.ViewType == ViewType.Section)
            {
                if (activeView.ViewType == ViewType.Section ||
                    activeView.ViewType == ViewType.DraftingView)
                {
                    validCombination = true;
                }
            }
            // For Elevation views: allowed if parent is an Elevation or DraftingView.
            else if (selectedView.ViewType == ViewType.Elevation)
            {
                if (activeView.ViewType == ViewType.Elevation ||
                    activeView.ViewType == ViewType.DraftingView)
                {
                    validCombination = true;
                }
            }
            // For Detail views, assume it is valid unless the parent is one of the restricted types.
            else if (selectedView.ViewType == ViewType.Detail)
            {
                if (activeView.ViewType != ViewType.FloorPlan &&
                    activeView.ViewType != ViewType.CeilingPlan)
                {
                    validCombination = true;
                }
            }

            // Check the result of the validation.
            if (!validCombination)
            {
                TaskDialog.Show("Reference Callout",
                    $"The active view ({activeView.ViewType}) does not support reference callouts to a view of type ({selectedView.ViewType}).");
                return Result.Cancelled;
            }
            // --- End validation ---

            // Get the center of the active view’s crop box.
            BoundingBoxXYZ cropBox = activeView.CropBox;
            if (cropBox == null)
            {
                TaskDialog.Show("Reference Callout", "Active view does not have a crop box.");
                return Result.Failed;
            }
            XYZ centerPoint = (cropBox.Min + cropBox.Max) * 0.5;

            // Create two diagonally opposed points.
            // The CreateReferenceCallout method requires that the two points differ when projected 
            // onto a plane perpendicular to the view direction.
            // We use a small offset from the center so the two points aren’t identical.
            double offset = 5.0; // Adjust as needed (Revit internal units)
            XYZ point1 = new XYZ(centerPoint.X - offset, centerPoint.Y - offset, centerPoint.Z);
            XYZ point2 = new XYZ(centerPoint.X + offset, centerPoint.Y + offset, centerPoint.Z);

            // Create the reference callout annotation within a transaction.
            using (Transaction trans = new Transaction(doc, "Create Reference Callout"))
            {
                trans.Start();

                // CreateReferenceCallout is a static method on the ViewSection class.
                // Parameters:
                //    document          : The Revit document.
                //    parentViewId      : The view in which the callout symbol appears.
                //    viewIdToReference : The view that will be referenced.
                //    point1, point2    : Two diagonally opposed corners of the callout symbol.
                ViewSection.CreateReferenceCallout(doc, activeView.Id, selectedViewId, point1, point2);

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}
