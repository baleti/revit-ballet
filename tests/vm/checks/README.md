# VM CI checks

Each `NN-name.cs` is a Roslyn script POSTed to the revit-ballet `/roslyn` endpoint of every Revit year under test
(globals `UIApp`, `UIDoc`, `Doc`; ASCII only). The matching `NN-name.expect` holds text that must appear in the
response `Output` field; `{YEAR}` is replaced by the Revit year. A check also fails if `Success` is false.

Add a check by dropping in a new pair of files. Checks run in name order inside one Revit session per year.
