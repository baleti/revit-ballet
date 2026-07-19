using Autodesk.Revit.DB;

namespace RevitBallet.Commands;

public class DuplicateTypesHandler : IDuplicateTypeNamesHandler
{
    public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
    {
      return DuplicateTypeAction.UseDestinationTypes;
    }
}
