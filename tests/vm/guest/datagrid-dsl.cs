// Drives the REAL DataGrid dialog inside Revit: opens it with fixture rows, types queries into its search box from a helper thread, reads back the
// rows the grid shows (its working set), the visible columns in display order, and closes the form. Returns one JSON object: { "<query>": { rows, columns } }.
using System.Windows.Forms;
using System.Reflection;
// Which DataGrid is under test:
//  - the installed one when it already has the query engine (a full CI run installs the build under test first);
//  - otherwise, on .NET 8 (Revit 2025+), the hot-reloaded build loaded into an isolated load context (hot reload swaps only the agent runtime, and
//    the add-in's own RevitBallet.Core is already loaded in Revit's context, so the new one must not bind to it);
//  - otherwise (.NET Framework years in the quick tier) the test cannot isolate a second copy: report skipped.
Type guis = null; string how = null;
if (Type.GetType("RevitBallet.Core.Query.QueryEngine, RevitBallet.Core") != null) { guis = typeof(CustomGUIs); how = "installed"; }
else
{
    var tAlc = Type.GetType("System.Runtime.Loader.AssemblyLoadContext, System.Runtime.Loader");
    var hot = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "revit-ballet", "hot-reload", UIApp.Application.VersionNumber);
    if (tAlc == null) return "{\"skipped\":\"installed build has no query engine and this Revit cannot isolate a hot-reloaded copy (.NET Framework): run the full CI\"}";
    object alc = Activator.CreateInstance(tAlc, new object[] { "datagrid-under-test", true });
    var load = tAlc.GetMethod("LoadFromStream", new[] { typeof(System.IO.Stream) });
    using (var cs = System.IO.File.OpenRead(System.IO.Path.Combine(hot, "RevitBallet.Core.dll"))) load.Invoke(alc, new object[] { cs });
    Assembly mainAsm;
    using (var ms = System.IO.File.OpenRead(System.IO.Path.Combine(hot, "revit-ballet.dll"))) mainAsm = (Assembly)load.Invoke(alc, new object[] { ms });
    guis = mainAsm.GetType("RevitBallet.Commands.CustomGUIs"); how = "hot-reload (isolated context)";
}
var results = new System.Collections.Generic.Dictionary<string, object>();
Func<string, string, string, string, string, string, Dictionary<string, object>> row = (mark, type, level, area, count, comments) =>
    new Dictionary<string, object> { { "Mark", mark }, { "TypeName", type }, { "Level", level }, { "Area", area }, { "Count", count }, { "Comments", comments } };
var data = new List<Dictionary<string, object>> {
    row("D1", "Single Flush 900", "Level 1", "1.89", "4", "fire rated"),
    row("D10", "Double Glazed 1800", "Level 10", "3.24", "12", ""),
    row("D2", "Single Flush 900", "Level 2", "1.89", "2", "needs check"),
    row("W1", "Fixed 1200x1500", "Level 1", "1.80", "30", "fire rated glass"),
    row("W2", "Casement 600x900", "Level 2", "0.54", "7", null),
};
var columns = new List<string> { "Mark", "TypeName", "Level", "Area", "Count", "Comments" };
var queries = new[] { "", "single", "//level 2", "!single", "/ft mark /at level", "/s count desc", "/rv", "//count >5", "//\"type name\" flush",
                      "/ft comments //level 1", "//level 1 /s mark desc", "single !//level 2", "/ft instance", "//zzz foo", "" };
string error = null;

Func<Control, Func<Control, bool>, Control> find = null;
find = (root, pred) => { foreach (Control c in root.Controls) { if (pred(c)) return c; var d = find(c, pred); if (d != null) return d; } return null; };

var helper = System.Threading.Tasks.Task.Run(() =>
{
    Form form = null;
    try
    {
        for (int i = 0; i < 300 && form == null; i++)
        {
            System.Threading.Thread.Sleep(100);
            foreach (Form f in Application.OpenForms) { try { if (find(f, c => c is DataGridView) != null && find(f, c => c is TextBox) != null) { form = f; break; } } catch { } }
        }
        if (form == null) { error = "the DataGrid form never appeared"; return; }
        System.Threading.Thread.Sleep(500);
        var field = guis.GetField("_cachedFilteredData", BindingFlags.NonPublic | BindingFlags.Static);
        foreach (var q in queries)
        {
            DataGridView grid = null; TextBox box = null;
            form.Invoke(new Action(() => { grid = (DataGridView)find(form, c => c is DataGridView); box = (TextBox)find(form, c => c is TextBox); box.Text = q; }));
            System.Threading.Thread.Sleep(400);
            form.Invoke(new Action(() =>
            {
                var shown = (List<Dictionary<string, object>>)field.GetValue(null);
                var cols = grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToList();
                results[q + "\u0000" + results.Count] = new Dictionary<string, object> { { "query", q }, { "rows", shown.Select(r => r["Mark"].ToString()).ToList() }, { "columns", cols }, { "gridRows", grid.RowCount } };
            }));
        }
    }
    catch (Exception ex) { error = ex.GetType().Name + ": " + ex.Message; }
    finally { try { if (form != null) form.Invoke(new Action(() => form.Close())); } catch { } }
});

try { guis.GetMethod("DataGrid", BindingFlags.Public | BindingFlags.Static).Invoke(null, new object[] { data, columns, false, null, null, false, null, false }); }
catch (Exception ex) { error = "DataGrid: " + (ex.InnerException ?? ex).GetType().Name + ": " + (ex.InnerException ?? ex).Message; }
helper.Wait(30000);
Func<string, string> q = x => "\"" + (x ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
Func<IEnumerable<string>, string> arr = xs => "[" + string.Join(",", xs.Select(q)) + "]";
var sb = new System.Text.StringBuilder();
sb.Append("{\"how\":").Append(q(how)).Append(",\"error\":").Append(error == null ? "null" : q(error)).Append(",\"steps\":[");
sb.Append(string.Join(",", results.Values.Select(v => { var d = (Dictionary<string, object>)v;
    return "{\"query\":" + q((string)d["query"]) + ",\"rows\":" + arr((List<string>)d["rows"]) + ",\"columns\":" + arr((List<string>)d["columns"]) + ",\"gridRows\":" + d["gridRows"] + "}"; })));
sb.Append("]}");
return sb.ToString();
