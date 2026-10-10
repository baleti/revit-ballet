import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInParameter
els = list(FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements())[:2000]
ok = 0
t = Transaction(doc, "bench")
t.Start()
for e in els:
    p = e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if p is not None and not p.IsReadOnly and p.Set("bench"):
        ok += 1
t.RollBack()
result = ok
