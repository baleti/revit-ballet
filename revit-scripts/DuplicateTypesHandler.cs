using Autodesk.Revit.DB;

public class DuplicateTypesHandler : IDuplicateTypeNamesHandler
{
    public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
    {
      return DuplicateTypeAction.UseDestinationTypes;
    }
}
