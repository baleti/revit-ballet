namespace Autodesk.Revit.DB
{
    public class Document { public string Title { get { return "stub document"; } } public override string ToString() { return "Document(stub)"; } }
    public class ElementId { public ElementId(int id) { Value = id; } public long Value { get; } }
}
