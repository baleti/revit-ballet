using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

/// <summary>
/// Lists selected elements across all Revit sessions in the network.
/// Uses peer-to-peer network communication to aggregate selections from multiple sessions.
/// </summary>
[Transaction(TransactionMode.Manual)]
public class ListSelectedInNetwork : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        TaskDialog.Show("ListSelectedInNetwork", "To be implemented");
        return Result.Cancelled;
    }
}
