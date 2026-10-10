var els = new FilteredElementCollector(Doc).WhereElementIsNotElementType().ToElements().Take(3000).ToList();
int n = 0;
foreach (var e in els) foreach (Parameter p in e.Parameters) { var s = p.AsValueString(); if (!string.IsNullOrEmpty(s)) n++; var d = p.Definition.Name; }
return els.Count + ":" + n;
