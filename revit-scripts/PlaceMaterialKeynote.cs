using System;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

[Transaction(TransactionMode.Manual)]
public class PostMaterialKeynoteCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;
        UIApplication uiApp = commandData.Application;

        // Ensure we have selected elements
        var selectedIds = uiDoc.Selection.GetElementIds().ToList();
        if (!selectedIds.Any())
        {
            message = "No elements selected.";
            return Result.Cancelled;
        }

        // Find a material keynote tag type
        // In a real scenario, prompt user or filter to ensure this is a Material Keynote type
        var keynoteTagType = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_KeynoteTags)
                                .OfClass(typeof(FamilySymbol))
                                .Cast<FamilySymbol>()
                                .FirstOrDefault();
        if (keynoteTagType == null)
        {
            message = "No keynote tag type found.";
            return Result.Cancelled;
        }

        using (Transaction t = new Transaction(doc, "Set Active Material Keynote"))
        {
            t.Start();

            // Activate the keynote tag type if it's not active
            if(!keynoteTagType.IsActive)
                keynoteTagType.Activate();

            // Optionally, attempt to set this type as the active type in the UI:
            // The Type Selector is not directly controllable via API, 
            // but you can ensure the last used tag type is this one by placing a dummy tag and undoing.
            // However, just activating should often be enough that Revit picks it as default.

            t.Commit();
        }

        // Now post the Tag By Category command.
        // This will start the tag placement command with the currently active tag family/type.
        // The user must then click on the element’s face to create a material keynote tag.
        RevitCommandId tagByCategoryCmdId = RevitCommandId.LookupPostableCommandId(PostableCommand.TagByCategory);
        if (uiApp.CanPostCommand(tagByCategoryCmdId))
        {
            uiApp.PostCommand(tagByCategoryCmdId);
        }
        else
        {
            message = "Cannot post TagByCategory command.";
            return Result.Failed;
        }

        return Result.Succeeded;
    }
}
