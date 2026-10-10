using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using RevitBallet.Core;
using RevitBallet.Core.Query;

namespace RevitBallet.Commands;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Search box -> rows, columns and order
    //
    //  The search box speaks the picker query DSL (~/.config/docs/query-dsl.md is the canonical definition,
    //  docs/remote-commands-and-query-dsl.md has the migration table from the old `$col:value` syntax).
    //  The engine lives in RevitBallet.Core.Query (no WinForms, unit-tested on Linux); this file only applies its result to the grid.
    // ──────────────────────────────────────────────────────────────

    /// <summary>The selection sets an element belongs to are a computed field of every grid that has a document.</summary>
    private const string SelectionSetField = "Selection Set";

    private static QuerySchema _querySchema;
    private static string _lastColumnState = "";

    /// <summary>Row -> lowercased text of all its cells, built once with the search index (reference-keyed: rows are dictionaries).</summary>
    private static Dictionary<object, string> _haystackByRow;

    private sealed class ReferenceComparer : IEqualityComparer<object>
    {
        public static readonly ReferenceComparer Instance = new ReferenceComparer();
        public new bool Equals(object x, object y) { return ReferenceEquals(x, y); }
        public int GetHashCode(object obj) { return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj); }
    }

    /// <summary>Builds the query vocabulary for this grid: its columns (names matched ignoring spaces, so `type name` reaches "TypeName"),
    /// the cached free-text haystack, and the computed selection-set field when a document is available.</summary>
    private static QuerySchema BuildQuerySchema(List<string> propertyNames)
    {
        var schema = QuerySchema.FromColumns(propertyNames);
        schema.IgnoreSpacesInNames = true;
        var cache = _haystackByRow;
        schema.Haystack = row =>
        {
            string text;
            if (cache != null && cache.TryGetValue(row, out text)) return text;
            var sb = new StringBuilder();
            foreach (string prop in propertyNames)
            {
                object v;
                if (row.TryGetValue(prop, out v) && v != null) sb.Append(v.ToString().ToLowerInvariant()).Append(' ');
            }
            return sb.ToString();
        };
        var uidoc = CurrentUIDoc as Autodesk.Revit.UI.UIDocument;
        if (uidoc != null && uidoc.Document != null && !propertyNames.Contains(SelectionSetField))
            schema.AddVirtualField(SelectionSetField, SelectionSetsOfRow);
        return schema;
    }

    /// <summary>Applies the search text: returns the surviving rows (in the query's order when it has one) and shows/hides/reorders the grid's columns.</summary>
    private static List<Dictionary<string, object>> ApplyFilters(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        string searchText,
        DataGridView grid,
        out bool queryOrdersRows)
    {
        queryOrdersRows = false;
        if (string.IsNullOrWhiteSpace(searchText))
        {
            ApplyColumns(grid, null);
            return entries;
        }

        if (_querySchema == null) _querySchema = BuildQuerySchema(propertyNames);
        QueryResult result = QueryEngine.Run(entries, _querySchema, searchText);
        queryOrdersRows = result.OrdersRows;

        // a column name the grid does not have (the virtual selection-set field) has no column to show
        ApplyColumns(grid, result.Columns);
        return result.Rows.Cast<Dictionary<string, object>>().ToList();
    }

    /// <summary>Shows exactly these columns, in this order (null = every column in its original order). Does nothing when nothing changed.</summary>
    private static void ApplyColumns(DataGridView grid, List<string> columns)
    {
        string state = columns == null ? "*" : string.Join("\u0001", columns);
        if (state == _lastColumnState) return;
        _lastColumnState = state;

        grid.SuspendLayout();
        try
        {
            var byName = new Dictionary<string, DataGridViewColumn>(StringComparer.Ordinal);
            foreach (DataGridViewColumn col in grid.Columns) byName[col.Name] = col;

            var visible = columns == null
                ? grid.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.Index).ToList()
                : columns.Where(byName.ContainsKey).Select(n => byName[n]).ToList();

            foreach (DataGridViewColumn col in grid.Columns) col.Visible = visible.Contains(col);
            int display = 0;
            foreach (DataGridViewColumn col in visible) col.DisplayIndex = display++;
            _lastVisibleColumns = new HashSet<string>(visible.Select(c => c.Name));
        }
        finally { grid.ResumeLayout(); }
    }

    // ──────────────────────────────────────────────────────────────
    //  Selection sets as a computed field (replaces the old `#name` token: `//selection "my set"`)
    // ──────────────────────────────────────────────────────────────

    private static List<string> _selectionSetNamesCache;

    private static List<string> GetAllSelectionSetNames()
    {
        if (_selectionSetNamesCache != null) return _selectionSetNamesCache;
        var names = new List<string>();
        var uidoc = CurrentUIDoc as Autodesk.Revit.UI.UIDocument;
        if (uidoc != null && uidoc.Document != null)
        {
            try
            {
                names = new Autodesk.Revit.DB.FilteredElementCollector(uidoc.Document)
                    .OfClass(typeof(Autodesk.Revit.DB.SelectionFilterElement))
                    .Cast<Autodesk.Revit.DB.SelectionFilterElement>()
                    .Select(s => s.Name).Distinct().OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
            }
            catch { /* a document that cannot be queried has no selection sets as far as the grid is concerned */ }
        }
        _selectionSetNamesCache = names;
        return names;
    }

    /// <summary>"set a|set b" for the sets that contain the row's element (or its viewport), null when it is in none.</summary>
    private static string SelectionSetsOfRow(IDictionary<string, object> row)
    {
        long id = RowId(row, "ElementID");
        if (id == 0) id = RowId(row, "Id");
        long viewport = RowId(row, "ViewportID");
        if (id == 0 && viewport == 0) return null;
        var hits = new List<string>();
        foreach (string name in GetAllSelectionSetNames())
        {
            HashSet<long> members = GetSelectionSetElementIds(name);
            if ((id != 0 && members.Contains(id)) || (viewport != 0 && members.Contains(viewport))) hits.Add(name);
        }
        return hits.Count == 0 ? null : string.Join("|", hits);
    }

    private static long RowId(IDictionary<string, object> row, string column)
    {
        object o;
        if (!row.TryGetValue(column, out o) || o == null) return 0;
        if (o is long l) return l;
        if (o is int i) return i;
        if (o is Autodesk.Revit.DB.ElementId eid) return eid.AsLong();
        long parsed;
        return long.TryParse(o.ToString(), out parsed) ? parsed : 0;
    }
}
