using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;

[Transaction(TransactionMode.ReadOnly)]
public class CopyCrossSession : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;

        var sel = uidoc.GetSelectionIds().Select(id => doc.GetElement(id)).ToList();
        var viewDependentElements = sel.Where(e => e != null && e.ViewSpecific).ToList();

        if (!viewDependentElements.Any())
        {
            TaskDialog.Show("Copy", "No view-dependent elements selected.");
            return Result.Cancelled;
        }

        var elementDataList = new List<CopiedElementData>();
        var familiesToExport = new HashSet<Family>();

        foreach (var elem in viewDependentElements)
        {
            if (elem is FamilyInstance fi && IsDetailFamilyInstance(fi))
            {
                LocationPoint lp = fi.Location as LocationPoint;
                if (lp != null)
                {
                    var symbol = fi.Symbol;
                    var family = symbol.Family;
                    double rotation = lp.Rotation;
                    XYZ point = lp.Point;

                    // Collect flipping and size parameters
                    bool handFlipped = fi.HandFlipped;
                    bool facingFlipped = fi.FacingFlipped;

                    // Collect instance parameters that can influence size/shape
                    var instParams = GetFamilyInstanceParameters(fi);

                    elementDataList.Add(new CopiedElementData
                    {
                        ElementType = CopiedElementType.FamilyInstance,
                        FamilyName = family.Name,
                        SymbolName = symbol.Name,
                        Location = new XYZData(point.X, point.Y, point.Z),
                        Rotation = rotation,
                        HandFlipped = handFlipped,
                        FacingFlipped = facingFlipped,
                        InstanceParameters = instParams
                    });
                    familiesToExport.Add(family);
                }
            }
            else if (elem is DetailCurve dc)
            {
                Curve c = dc.GeometryCurve;
                if (c != null)
                {
                    elementDataList.Add(new CopiedElementData
                    {
                        ElementType = CopiedElementType.DetailCurve,
                        Curves = new List<CurveData> { CurveToCurveData(c) },
                        LineStyleName = GetElementLineStyleName(dc)
                    });
                }
            }
            else if (elem is FilledRegion fr)
            {
                var frType = doc.GetElement(fr.GetTypeId()) as FilledRegionType;
                if (frType != null)
                {
                    var boundaries = fr.GetBoundaries();
                    var boundaryData = new List<List<CurveData>>();
                    foreach (var loop in boundaries)
                    {
                        var loopData = new List<CurveData>();
                        foreach (var curve in loop)
                        {
                            loopData.Add(CurveToCurveData(curve));
                        }
                        boundaryData.Add(loopData);
                    }

                    elementDataList.Add(new CopiedElementData
                    {
                        ElementType = CopiedElementType.FilledRegion,
                        FilledRegionTypeName = frType.Name,
                        FilledRegionBoundaries = boundaryData
                    });
                }
            }
            else if (elem is TextNote textNote)
            {
                var tnType = doc.GetElement(textNote.GetTypeId()) as TextNoteType;
                XYZ point = textNote.Coord;
                elementDataList.Add(new CopiedElementData
                {
                    ElementType = CopiedElementType.TextNote,
                    Text = textNote.Text,
                    Location = new XYZData(point.X, point.Y, point.Z),
                    Rotation = 0, // rotation ignored
                    TextTypeName = tnType?.Name
                });
            }
            else if (elem is Dimension dim)
            {
                var dimType = doc.GetElement(dim.GetTypeId()) as DimensionType;
                elementDataList.Add(new CopiedElementData
                {
                    ElementType = CopiedElementType.Dimension,
                    DimensionTypeName = dimType?.Name
                });
            }
        }

        string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string folderPath = Path.Combine(appData, "revit-scripts");
        if (!Directory.Exists(folderPath))
        {
            Directory.CreateDirectory(folderPath);
        }

        string filePath = Path.Combine(folderPath, "selectedElements.json");
        string json = JsonConvert.SerializeObject(elementDataList, Formatting.Indented);
        File.WriteAllText(filePath, json);

        // Export families
        foreach (var fam in familiesToExport)
        {
            try
            {
                string familyFilePath = Path.Combine(folderPath, fam.Name + ".rfa");
                if (!File.Exists(familyFilePath))
                {
                    ExportFamily(doc, fam, familyFilePath);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Family", $"Could not export family {fam.Name}:\n{ex.Message}");
            }
        }

        TaskDialog.Show("Copy", "View-dependent elements copied. Ready to paste in another session.");
        return Result.Succeeded;
    }

    private bool IsDetailFamilyInstance(FamilyInstance fi)
    {
        return fi.ViewSpecific;
    }

    private Dictionary<string, string> GetFamilyInstanceParameters(FamilyInstance fi)
    {
        var result = new Dictionary<string, string>();
        foreach (Parameter p in fi.Parameters)
        {
            if (!p.IsReadOnly && p.Definition != null && p.StorageType != StorageType.None)
            {
                // Only store parameters that might influence geometry
                // This is heuristic: we can store all non-read-only parameters.
                // For example, dimension-like parameters are often double values.
                string paramName = p.Definition.Name;
                string value = ParameterToString(p);
                if (value != null)
                {
                    result[paramName] = value;
                }
            }
        }
        return result;
    }

    private string ParameterToString(Parameter p)
    {
        switch (p.StorageType)
        {
            case StorageType.Double:
                return p.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
            case StorageType.Integer:
                return p.AsInteger().ToString();
            case StorageType.String:
                return p.AsString();
            case StorageType.ElementId:
                // For cross-session copy, element IDs won't match. Skip ElementId parameters.
                return null;
            default:
                return null;
        }
    }

    private void ExportFamily(Document doc, Family family, string filePath)
    {
        Document famDoc = doc.EditFamily(family);
        SaveAsOptions sao = new SaveAsOptions();
        sao.OverwriteExistingFile = true;
        famDoc.SaveAs(filePath, sao);
        famDoc.Close(false);
    }

    private CurveData CurveToCurveData(Curve c)
    {
        Line line = c as Line;
        if (line != null)
        {
            return new CurveData
            {
                CurveType = "Line",
                Start = new XYZData(line.GetEndPoint(0).X, line.GetEndPoint(0).Y, line.GetEndPoint(0).Z),
                End = new XYZData(line.GetEndPoint(1).X, line.GetEndPoint(1).Y, line.GetEndPoint(1).Z)
            };
        }
        throw new NotImplementedException("Only lines are currently supported.");
    }

    private string GetElementLineStyleName(DetailCurve dc)
    {
        var gs = dc.LineStyle as GraphicsStyle;
        return gs?.Name;
    }

    private class CopiedElementData
    {
        public CopiedElementType ElementType { get; set; }

        public XYZData Location { get; set; }
        public double Rotation { get; set; }

        public string FamilyName { get; set; }
        public string SymbolName { get; set; }

        public bool HandFlipped { get; set; }
        public bool FacingFlipped { get; set; }
        public Dictionary<string, string> InstanceParameters { get; set; }

        public List<CurveData> Curves { get; set; }
        public string LineStyleName { get; set; }

        public string FilledRegionTypeName { get; set; }
        public List<List<CurveData>> FilledRegionBoundaries { get; set; }

        public string Text { get; set; }
        public string TextTypeName { get; set; }

        public string DimensionTypeName { get; set; }
    }

    private enum CopiedElementType
    {
        FamilyInstance,
        DetailCurve,
        FilledRegion,
        TextNote,
        Dimension
    }

    private class XYZData
    {
        public XYZData() { }
        public XYZData(double x, double y, double z) { X = x; Y = y; Z = z; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public XYZ ToXYZ()
        {
            return new XYZ(X, Y, Z);
        }
    }

    private class CurveData
    {
        public string CurveType { get; set; }
        public XYZData Start { get; set; }
        public XYZData End { get; set; }
    }
}
