using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

#if REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
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

#endif
