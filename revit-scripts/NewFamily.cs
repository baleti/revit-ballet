using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

[Transaction(TransactionMode.Manual)]
public class NewFamily : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        RevitCommandId newFamilyCommandId = RevitCommandId.LookupPostableCommandId(PostableCommand.NewFamily);
        commandData.Application.PostCommand(newFamilyCommandId);

        return Result.Succeeded;
    }
}
