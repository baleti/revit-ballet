using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitBallet.Core.Query
{
    /// <summary>Runs a parsed query over rows: filter rows, then work out the columns, then order.</summary>
    public static class QueryEngine
    {
        public static QueryResult Run(IList<IDictionary<string, object>> rows, QuerySchema schema, string query)
        {
            return Apply(rows, schema, QueryParser.Parse(query));
        }

        public static QueryResult Run(IEnumerable<Dictionary<string, object>> rows, QuerySchema schema, string query)
        {
            return Apply(rows.Select(r => (IDictionary<string, object>)r).ToList(), schema, QueryParser.Parse(query));
        }

        public static QueryResult Apply(IList<IDictionary<string, object>> rows, QuerySchema schema, QueryPlan plan)
        {
            var result = new QueryResult();

            // 1. rows: an order-independent AND of terms
            var keep = new List<int>();
            for (int i = 0; i < rows.Count; i++)
            {
                bool ok = true;
                foreach (var term in plan.RowTerms)
                    if (!TermMatches(rows[i], schema, term)) { ok = false; break; }
                if (ok) keep.Add(i);
            }

            // 2. order: stable; `/sort` replaces the default order, `/reverse` flips whatever is in effect
            string[] keys = ResolveSortKeys(schema, plan.Sort);
            if (keys != null)
            {
                bool desc = plan.Sort.Descending;
                var cache = new Dictionary<int, string[]>();
                foreach (int idx in keep)
                    cache[idx] = keys.Select(k => Cell(rows[idx], k)).ToArray();
                var sorted = new List<int>(keep);
                sorted.Sort((a, b) =>
                {
                    for (int k = 0; k < keys.Length; k++)
                    {
                        int c = FieldCompare.Compare(cache[a][k], cache[b][k], desc);
                        if (c != 0) return c;
                    }
                    return a.CompareTo(b);          // stable
                });
                keep = sorted;
            }
            if (plan.Reverse) keep.Reverse();

            foreach (int idx in keep) { result.Rows.Add(rows[idx]); result.RowIndices.Add(idx); }

            // 3. columns: a left-to-right pipeline, then the auto-shown filter fields
            result.Columns = ComputeColumns(result.Rows, schema, plan, result.FilterFields);
            return result;
        }

        // ---------------------------------------------------------------- rows

        /// <summary>The cell's text, or null when the row has no value for it.</summary>
        public static string Cell(IDictionary<string, object> row, string field)
        {
            object v;
            if (row == null || !row.TryGetValue(field, out v) || v == null) return null;
            string s = v.ToString();
            return s.Length == 0 ? null : s;
        }

        private static string Haystack(IDictionary<string, object> row, QuerySchema schema)
        {
            if (schema.Haystack != null) return schema.Haystack(row) ?? "";
            return string.Join(" ", row.Values.Where(v => v != null).Select(v => v.ToString()));
        }

        private static bool TermMatches(IDictionary<string, object> row, QuerySchema schema, RowTerm t)
        {
            bool inert;
            bool positive = Evaluate(row, schema, t, out inert);
            if (inert) return true;
            return t.Negated ? !positive : positive;
        }

        private static bool Evaluate(IDictionary<string, object> row, QuerySchema schema, RowTerm t, out bool inert)
        {
            inert = false;
            switch (t.Form)
            {
                case TermForm.Text:
                    return QuerySchema.Contains(Haystack(row, schema), t.Text);

                case TermForm.PathOnly:
                    {
                        var r = schema.Resolve(t.Path);
                        if (r.BareGroups.Count > 0) return GroupExists(row, schema, r.BareGroups);
                        return QuerySchema.Contains(Haystack(row, schema), t.Text);       // flat-only or unresolved: free text
                    }

                case TermForm.Scoped:
                    if (!t.HasValue) { inert = true; return true; }                        // `path:` still being typed
                    return ScopedMatch(row, schema, t.Path, t.Value);

                case TermForm.Via:
                    {
                        if (!t.HasValue)
                        {
                            if (string.IsNullOrEmpty(t.Path)) { inert = true; return true; }
                            var r = schema.Resolve(t.Path);
                            if (r.BareGroups.Count > 0) return GroupExists(row, schema, r.BareGroups);   // existence filter
                            inert = true; return true;                                                       // flat or nothing: waits for a value
                        }
                        return ScopedMatch(row, schema, t.Path, t.Value);
                    }
            }
            inert = true;
            return true;
        }

        private static bool GroupExists(IDictionary<string, object> row, QuerySchema schema, List<string> groups)
        {
            foreach (var g in groups)
                foreach (var sub in schema.Groups[g])
                    if (Cell(row, g + "." + sub) != null) return true;
            return false;
        }

        private static bool ScopedMatch(IDictionary<string, object> row, QuerySchema schema, string path, string value)
        {
            var r = schema.Resolve(path ?? "");
            if (!r.Any) return false;                         // a complete path that resolves to nothing narrows to nothing

            string op = null, operand = value;
            if (value.StartsWith(">=", StringComparison.Ordinal) || value.StartsWith("<=", StringComparison.Ordinal)) { op = value.Substring(0, 2); operand = value.Substring(2); }
            else if (value.StartsWith(">", StringComparison.Ordinal) || value.StartsWith("<", StringComparison.Ordinal)) { op = value.Substring(0, 1); operand = value.Substring(1); }

            foreach (var f in r.Fields)
            {
                string cell = Cell(row, f);
                if (op != null)
                {
                    bool m;
                    if (cell != null && FieldCompare.TryCompareFilter(cell, op, operand, out m) && m) return true;
                }
                else if (cell != null ? QuerySchema.Contains(cell, value) : value.Length == 0) return true;
            }
            return false;
        }

        // ---------------------------------------------------------------- order

        private static string[] ResolveSortKeys(QuerySchema schema, SortSpec sort)
        {
            if (sort == null || sort.Paths.Count == 0) return null;
            var keys = new List<string>();
            foreach (string p in sort.Paths)
            {
                string f = schema.ResolveSingle(p);
                if (f == null) return null;                    // one bad key makes the whole sort inert
                keys.Add(f);
            }
            return keys.ToArray();
        }

        // ---------------------------------------------------------------- columns

        private static List<string> ComputeColumns(List<IDictionary<string, object>> rows, QuerySchema schema, QueryPlan plan, List<string> filterFields)
        {
            var shown = new List<string>(schema.DefaultColumns);
            foreach (var op in plan.ColumnOps)
            {
                var r = schema.Resolve(op.Path);
                var fields = r.Fields;
                if (op.Verb == Verb.FilterType && op.Via && !r.Any && !string.IsNullOrEmpty(op.Path))
                {
                    // value pattern: show whichever columns actually hold a matching value in the current rows (a discovery reset)
                    shown = schema.AllFields().Where(f => rows.Any(row => ValueMatchesPattern(Cell(row, f), op.Path))).ToList();
                    continue;
                }
                switch (op.Verb)
                {
                    case Verb.FilterType: shown = shown.Where(fields.Contains).ToList(); break;
                    case Verb.AddType: foreach (var f in fields) if (!shown.Contains(f)) shown.Add(f); break;
                    case Verb.RemoveType: shown = shown.Where(f => !fields.Contains(f)).ToList(); break;
                }
            }

            // auto-shown filter fields: additive, even past an /ft or /rt
            foreach (var t in plan.RowTerms)
            {
                var f = FilterField(schema, t);
                foreach (var name in f)
                    if (!filterFields.Contains(name)) filterFields.Add(name);
            }
            foreach (var name in filterFields)
                if (!shown.Contains(name)) shown.Add(name);
            return shown;
        }

        private static List<string> FilterField(QuerySchema schema, RowTerm t)
        {
            var none = new List<string>();
            switch (t.Form)
            {
                case TermForm.Scoped:
                    {
                        var r = schema.Resolve(t.Path ?? "");
                        if (!t.HasValue) return r.Fields.Count == 1 ? r.Fields : (r.BareGroups.Count == 1 ? GroupDefault(schema, r) : none);
                        return r.BareGroups.Count > 0 ? GroupDefault(schema, r).Concat(r.Fields.Where(f => !IsGroupField(schema, f, r))).ToList() : r.Fields;
                    }
                case TermForm.Via:
                    {
                        if (string.IsNullOrEmpty(t.Path) && !t.HasValue) return none;
                        var r = schema.Resolve(t.Path ?? "");
                        if (!t.HasValue) return r.Fields.Count == 1 ? r.Fields : (r.BareGroups.Count == 1 ? GroupDefault(schema, r) : none);
                        return r.BareGroups.Count > 0 ? GroupDefault(schema, r).Concat(r.Fields.Where(f => !IsGroupField(schema, f, r))).ToList() : r.Fields;
                    }
                case TermForm.PathOnly:
                    {
                        var r = schema.Resolve(t.Path);
                        return r.BareGroups.Count > 0 ? GroupDefault(schema, r) : none;
                    }
            }
            return none;
        }

        private static bool IsGroupField(QuerySchema schema, string field, PathResolution r)
        {
            int dot = field.IndexOf('.');
            return dot > 0 && r.BareGroups.Contains(field.Substring(0, dot));
        }

        private static List<string> GroupDefault(QuerySchema schema, PathResolution r)
        {
            var l = new List<string>();
            foreach (var g in r.BareGroups)
            {
                string d = schema.GroupDefaultField(g);
                if (d != null) l.Add(d);
            }
            return l;
        }

        /// <summary>`/ft/64*`: substring without a `*`, anchored glob with one (`64*` starts with, `*64` ends with).</summary>
        private static bool ValueMatchesPattern(string cell, string pattern)
        {
            if (cell == null) return false;
            if (!pattern.Contains("*")) return QuerySchema.Contains(cell, pattern);
            return TextMatch.MatchesGlobPattern(cell, pattern);
        }
    }
}
