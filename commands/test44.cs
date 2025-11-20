using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class CheckGroupExcludedMembers : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Reference pickedRef = null;
            try
            {
                pickedRef = uidoc.Selection.PickObject(ObjectType.Element, new ModelGroupFilter(), "Select a model group");
            }
            catch
            {
                return Result.Cancelled;
            }

            Group group = doc.GetElement(pickedRef) as Group;
            if (group == null)
            {
                TaskDialog.Show("Error", "Selected element is not a model group.");
                return Result.Failed;
            }

            bool hasExcluded = group.Name.Contains("(members excluded)");
            string resultMessage = hasExcluded
                ? "The selected model group has excluded members."
                : "The selected model group does not have excluded members.";

            TaskDialog.Show("Model Group Excluded Members", resultMessage);

            return Result.Succeeded;
        }

        private class ModelGroupFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Group && elem.Category != null && elem.Category.Id.AsLong() == (int)BuiltInCategory.OST_IOSModelGroups;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
    }
}
