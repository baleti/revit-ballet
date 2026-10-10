import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import XYZ
o = XYZ(0, 0, 0)
s = 0.0
for i in range(50000):
    s += XYZ(i, i % 7, 1).DistanceTo(o)
result = round(s)
