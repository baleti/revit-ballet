using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitBallet.Core.Query
{
    /// <summary>The result of resolving a type path against a schema.</summary>
    public sealed class PathResolution
    {
        /// <summary>Every field the path reaches (a bare group contributes all its subfields), in schema order.</summary>
        public List<string> Fields = new List<string>();
        /// <summary>The groups the path named without a subfield (existence filters apply to these).</summary>
        public List<string> BareGroups = new List<string>();
        /// <summary>True when at least one flat field or group matched.</summary>
        public bool Any { get { return Fields.Count > 0; } }
    }

    /// <summary>
    /// The vocabulary of one picker: flat fields, optional groups of subfields ("claude" -> "claude.title"), the default columns and the
    /// free-text haystack. Names are each host's own; only the grammar is shared.
    /// A schema without groups treats a dot as part of a field name, so "Ver. 1.0" is one field, not a path.
    /// </summary>
    public sealed class QuerySchema
    {
        public readonly List<string> Fields = new List<string>();
        public readonly Dictionary<string, List<string>> Groups = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, string> GroupDefaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public readonly List<string> DefaultColumns = new List<string>();

        /// <summary>Free-text haystack for bare words. Default: every field value, space separated.</summary>
        public Func<IDictionary<string, object>, string> Haystack;

        public static QuerySchema FromColumns(IEnumerable<string> columns)
        {
            var s = new QuerySchema();
            foreach (var c in columns)
            {
                if (c == null || s.Fields.Contains(c)) continue;
                s.Fields.Add(c);
                s.DefaultColumns.Add(c);
            }
            return s;
        }

        public QuerySchema AddField(string name, bool defaultColumn = true)
        {
            if (!Fields.Contains(name)) Fields.Add(name);
            if (defaultColumn && !DefaultColumns.Contains(name)) DefaultColumns.Add(name);
            return this;
        }

        /// <summary>Adds a group; each subfield becomes the field "group.sub". The first subfield is the default one (used by /sort).</summary>
        public QuerySchema AddGroup(string name, IEnumerable<string> subfields, string defaultSub = null, bool defaultColumns = false)
        {
            var subs = subfields.ToList();
            Groups[name] = subs;
            GroupDefaults[name] = defaultSub ?? (subs.Count > 0 ? subs[0] : null);
            foreach (var sub in subs)
            {
                string full = name + "." + sub;
                if (defaultColumns && !DefaultColumns.Contains(full)) DefaultColumns.Add(full);
            }
            return this;
        }

        /// <summary>All field names in schema order: flat fields first, then each group's subfields.</summary>
        public List<string> AllFields()
        {
            var all = new List<string>(Fields);
            foreach (var g in Groups) foreach (var sub in g.Value) all.Add(g.Key + "." + sub);
            return all;
        }

        public string GroupDefaultField(string group)
        {
            string sub;
            List<string> subs;
            if (!GroupDefaults.TryGetValue(group, out sub) || sub == null) return null;
            if (!Groups.TryGetValue(group, out subs)) return null;
            return group + "." + sub;
        }

        /// <summary>
        /// Resolves a type path: dot-separated segments, each substring-matched case-insensitively and unioned; `*` as the second
        /// segment is every subfield; an empty path matches everything. A path that resolves to nothing has no fields.
        /// </summary>
        public PathResolution Resolve(string path)
        {
            var r = new PathResolution();
            if (path == null) return r;
            string[] segs = Groups.Count == 0 ? new[] { path } : path.Split('.');
            string first = segs[0];
            var fields = new List<string>();

            foreach (var f in Fields)
                if (segs.Length == 1 && Contains(f, first)) fields.Add(f);

            foreach (var g in Groups)
            {
                if (!Contains(g.Key, first)) continue;
                if (segs.Length == 1)
                {
                    r.BareGroups.Add(g.Key);
                    foreach (var sub in g.Value) fields.Add(g.Key + "." + sub);
                }
                else if (segs.Length == 2)
                {
                    foreach (var sub in g.Value)
                        if (segs[1] == "*" || Contains(sub, segs[1])) fields.Add(g.Key + "." + sub);
                }
            }
            // schema order, no duplicates
            var order = AllFields();
            r.Fields = order.Where(fields.Contains).Distinct().ToList();
            return r;
        }

        /// <summary>
        /// Resolves a path that must name exactly one field (a sort key): a bare group means its default subfield, `*` and ambiguous
        /// paths are rejected (null).
        /// </summary>
        public string ResolveSingle(string path)
        {
            if (string.IsNullOrEmpty(path) || path.Contains("*")) return null;
            var r = Resolve(path);
            var picks = new List<string>();
            foreach (var f in r.Fields)
            {
                int dot = Groups.Count == 0 ? -1 : f.IndexOf('.');
                if (dot > 0 && r.BareGroups.Contains(f.Substring(0, dot)))
                {
                    string def = GroupDefaultField(f.Substring(0, dot));
                    if (def != null && !picks.Contains(def)) picks.Add(def);
                }
                else if (!picks.Contains(f)) picks.Add(f);
            }
            return picks.Count == 1 ? picks[0] : null;
        }

        internal static bool Contains(string haystack, string needle)
        {
            if (string.IsNullOrEmpty(needle)) return true;
            return haystack != null && haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
