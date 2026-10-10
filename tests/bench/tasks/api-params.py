import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector
els = list(FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements())[:3000]
n = 0
for e in els:
    for p in e.Parameters:
        s = p.AsValueString()
        if s:
            n += 1
        d = p.Definition.Name
result = "%d:%d" % (len(els), n)
