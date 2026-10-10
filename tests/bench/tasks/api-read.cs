var els = new FilteredElementCollector(Doc).WhereElementIsNotElementType().ToElements();
var cats = new HashSet<string>(); long ids = 0;
foreach (var e in els) { ids += e.Id.Value; var c = e.Category; if (c != null) cats.Add(c.Name); }
return els.Count + ":" + cats.Count;
