// Graphics-bound work: switch to a 3D view, refresh it, export it as a 1920 px image. Run with a model open and a 3D view in it.
var doc = Doc; var uidoc = UIDoc;
var view = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate && v.Name == "{3D}")
        ?? new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().First(v => !v.IsTemplate);
var times = new List<double>(); var export = new List<double>(); var sizes = new List<long>();
uidoc.ActiveView = view;
for (int i = 0; i < 7; i++)
{
    var sw = System.Diagnostics.Stopwatch.StartNew();
    uidoc.RefreshActiveView();
    times.Add(sw.Elapsed.TotalMilliseconds);
    var dir = Path.Combine(Path.GetTempPath(), "rb-gfx"); Directory.CreateDirectory(dir);
    var file = Path.Combine(dir, "export" + i);
    var o = new ImageExportOptions { ZoomType = ZoomFitType.FitToPage, PixelSize = 1920, FilePath = file, FitDirection = FitDirectionType.Horizontal,
        HLRandWFViewsFileType = ImageFileType.PNG, ShadowViewsFileType = ImageFileType.PNG, ImageResolution = ImageResolution.DPI_150, ExportRange = ExportRange.SetOfViews };
    o.SetViewsAndSheets(new List<ElementId> { view.Id });
    sw.Restart();
    doc.ExportImage(o);
    export.Add(sw.Elapsed.TotalMilliseconds);
    var written = Directory.GetFiles(dir, "export" + i + "*").FirstOrDefault();
    sizes.Add(written == null ? -1 : new FileInfo(written).Length);
}
Func<List<double>, string> med = l => { var s = l.Skip(2).OrderBy(x => x).ToList(); return s[s.Count / 2].ToString("F0"); };
Console.WriteLine("view: " + view.Name + "  detail level: " + view.DetailLevel + "  style: " + view.DisplayStyle);
Console.WriteLine("RefreshActiveView median ms (5 timed runs): " + med(times) + "   all: " + string.Join(", ", times.Select(t => t.ToString("F0"))));
Console.WriteLine("ExportImage 1920 px median ms (5 timed runs): " + med(export) + "   all: " + string.Join(", ", export.Select(t => t.ToString("F0"))));
Console.WriteLine("exported file sizes (bytes): " + string.Join(", ", sizes));
return "ok";
