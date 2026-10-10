var els = new FilteredElementCollector(Doc).WhereElementIsNotElementType().ToElements().Take(2000).ToList();
int ok = 0;
using (var t = new Transaction(Doc, "bench")) {
    t.Start();
    foreach (var e in els) { var p = e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS); if (p != null && !p.IsReadOnly && p.Set("bench")) ok++; }
    t.RollBack();
}
return ok;
