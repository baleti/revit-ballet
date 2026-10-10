namespace Autodesk.Revit.UI
{
    public class UIDocument { public Autodesk.Revit.DB.Document Document { get { return new Autodesk.Revit.DB.Document(); } } }
    public class UIApplication { public string VersionNumber { get { return "2026"; } } public UIDocument ActiveUIDocument { get { return new UIDocument(); } } }
}
