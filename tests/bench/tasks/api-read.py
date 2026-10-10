import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector
els = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
cats = set()
ids = 0
for e in els:
    ids += e.Id.Value
    c = e.Category
    if c is not None:
        cats.add(c.Name)
result = "%d:%d" % (len(els), len(cats))
