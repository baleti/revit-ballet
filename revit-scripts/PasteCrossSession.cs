using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;

[Transaction(TransactionMode.Manual)]
public class PasteCrossSession : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;

        string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string folderPath = Path.Combine(appData, "revit-scripts");
        string filePath = Path.Combine(folderPath, "selectedElements.json");

        if (!File.Exists(filePath))
        {
            TaskDialog.Show("Paste", "No copied data found.");
            return Result.Cancelled;
        }

        string json = File.ReadAllText(filePath);
        var elementDataList = JsonConvert.DeserializeObject<List<CopiedElementData>>(json);

        using (Transaction t = new Transaction(doc, "Paste Cross Session Elements"))
        {
            t.Start();

            foreach (var data in elementDataList)
            {
                switch (data.ElementType)
                {
                    case CopiedElementType.FamilyInstance:
                        PasteFamilyInstance(doc, data, folderPath);
                        break;
                    case CopiedElementType.DetailCurve:
                        PasteDetailCurve(doc, data);
                        break;
                    case CopiedElementType.FilledRegion:
                        PasteFilledRegion(doc, data);
                        break;
                    case CopiedElementType.TextNote:
                        PasteTextNote(doc, data);
                        break;
                    case CopiedElementType.Dimension:
                        TaskDialog.Show("Paste", $"Skipping dimension '{data.DimensionTypeName}' - cannot recreate references.");
                        break;
                }
            }

            t.Commit();
        }

        TaskDialog.Show("Paste", "Elements pasted successfully (with some limitations).");
        return Result.Succeeded;
    }

    private void PasteFamilyInstance(Document doc, CopiedElementData data, string folderPath)
    {
        FamilySymbol symbol = FindFamilySymbol(doc, data.FamilyName, data.SymbolName);
        if (symbol == null)
        {
            string familyFilePath = Path.Combine(folderPath, data.FamilyName + ".rfa");
            if (File.Exists(familyFilePath))
            {
                Family loadedFamily = null;
                if (doc.LoadFamily(familyFilePath, out loadedFamily))
                {
                    symbol = FindFamilySymbol(doc, data.FamilyName, data.SymbolName);
                }
            }

            if (symbol == null)
            {
                TaskDialog.Show("Paste", $"Could not find or load family symbol {data.FamilyName}:{data.SymbolName}. Skipping.");
                return;
            }
        }

        if (!symbol.IsActive) symbol.Activate();

        XYZ location = data.Location.ToXYZ();
        FamilyInstance fi = doc.Create.NewFamilyInstance(location, symbol, doc.ActiveView);

        // Apply rotation
        double rotation = data.Rotation;
        if (Math.Abs(rotation) > 1e-9)
        {
            Line axis = Line.CreateBound(location, location.Add(XYZ.BasisZ));
            ElementTransformUtils.RotateElement(doc, fi.Id, axis, rotation);
        }

        bool desiredHandFlipped = data.HandFlipped;
        bool desiredFacingFlipped = data.FacingFlipped;

        // Check if we can flip hand
        if (fi.CanFlipHand && fi.HandFlipped != desiredHandFlipped)
        {
            fi.flipHand();
        }

        // Check if we can flip facing
        if (fi.CanFlipFacing && fi.FacingFlipped != desiredFacingFlipped)
        {
            fi.flipFacing();
        }

        // Apply instance parameters
        if (data.InstanceParameters != null)
        {
            foreach (var kvp in data.InstanceParameters)
            {
                SetFamilyInstanceParameter(fi, kvp.Key, kvp.Value);
            }
        }
    }

    private void SetFamilyInstanceParameter(FamilyInstance fi, string paramName, string val)
    {
        Parameter p = fi.LookupParameter(paramName);
        if (p == null || p.IsReadOnly) return;

        switch (p.StorageType)
        {
            case StorageType.Double:
                if (double.TryParse(val, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dVal))
                {
                    p.Set(dVal);
                }
                break;
            case StorageType.Integer:
                if (int.TryParse(val, out int iVal))
                {
                    p.Set(iVal);
                }
                break;
            case StorageType.String:
                p.Set(val);
                break;
            case StorageType.ElementId:
                // We skip ElementId parameters because they may not match across documents
                break;
        }
    }

    private FamilySymbol FindFamilySymbol(Document doc, string familyName, string symbolName)
    {
        return new FilteredElementCollector(doc)
            .OfClass(typeof(FamilySymbol))
            .Cast<FamilySymbol>()
            .FirstOrDefault(fs => fs.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) &&
                                  fs.Name.Equals(symbolName, StringComparison.OrdinalIgnoreCase));
    }

    private void PasteDetailCurve(Document doc, CopiedElementData data)
    {
        if (data.Curves != null)
        {
            foreach (var cData in data.Curves)
            {
                if (cData.CurveType == "Line")
                {
                    XYZ start = cData.Start.ToXYZ();
                    XYZ end = cData.End.ToXYZ();
                    Curve line = Line.CreateBound(start, end);
                    DetailCurve dc = doc.Create.NewDetailCurve(doc.ActiveView, line);

                    if (!string.IsNullOrEmpty(data.LineStyleName))
                    {
                        SetLineStyle(doc, dc, data.LineStyleName);
                    }
                }
                else
                {
                    TaskDialog.Show("Paste", "Only line detail curves are currently supported.");
                }
            }
        }
    }

    private void SetLineStyle(Document doc, DetailCurve dc, string styleName)
    {
        Category linesCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
        var subCats = linesCategory.SubCategories;
        Category targetCat = subCats.OfType<Category>().FirstOrDefault(c => c.Name.Equals(styleName, StringComparison.OrdinalIgnoreCase));
        if (targetCat != null)
        {
            var gs = targetCat.GetGraphicsStyle(GraphicsStyleType.Projection);
            if (gs != null)
            {
                dc.LineStyle = gs;
            }
        }
    }

    private void PasteFilledRegion(Document doc, CopiedElementData data)
    {
        FilledRegionType frType = new FilteredElementCollector(doc)
            .OfClass(typeof(FilledRegionType))
            .Cast<FilledRegionType>()
            .FirstOrDefault(t => t.Name.Equals(data.FilledRegionTypeName, StringComparison.OrdinalIgnoreCase));

        if (frType == null)
        {
            TaskDialog.Show("Paste", $"Could not find FilledRegionType '{data.FilledRegionTypeName}'. Skipping filled region.");
            return;
        }

        List<CurveLoop> loops = new List<CurveLoop>();
        foreach (var loopData in data.FilledRegionBoundaries)
        {
            CurveLoop loop = new CurveLoop();
            foreach (var cData in loopData)
            {
                if (cData.CurveType == "Line")
                {
                    XYZ start = cData.Start.ToXYZ();
                    XYZ end = cData.End.ToXYZ();
                    loop.Append(Line.CreateBound(start, end));
                }
                else
                {
                    TaskDialog.Show("Paste", "Only line curves supported for filled regions.");
                }
            }
            loops.Add(loop);
        }

        FilledRegion.Create(doc, frType.Id, doc.ActiveView.Id, loops);
    }

    private void PasteTextNote(Document doc, CopiedElementData data)
    {
        TextNoteType tnt = new FilteredElementCollector(doc)
            .OfClass(typeof(TextNoteType))
            .Cast<TextNoteType>()
            .FirstOrDefault(x => x.Name.Equals(data.TextTypeName, StringComparison.OrdinalIgnoreCase));

        if (tnt == null)
        {
            tnt = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault();
        }

        XYZ loc = data.Location.ToXYZ();
        TextNoteOptions opts = new TextNoteOptions(tnt.Id);
        TextNote.Create(doc, doc.ActiveView.Id, loc, data.Text, opts);
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
