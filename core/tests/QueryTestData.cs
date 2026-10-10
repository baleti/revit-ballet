using System.Collections.Generic;
using System.Linq;
using RevitBallet.Core.Query;

namespace RevitBallet.Core.Tests
{
    /// <summary>Shared fixtures for the query DSL tests.</summary>
    internal static class QueryTestData
    {
        public static IDictionary<string, object> Row(params object[] kv)
        {
            var d = new Dictionary<string, object>();
            for (int i = 0; i < kv.Length; i += 2) d[(string)kv[i]] = kv[i + 1];
            return d;
        }

        /// <summary>A DataGrid-like flat table: doors and windows.</summary>
        public static List<IDictionary<string, object>> Elements()
        {
            return new List<IDictionary<string, object>>
            {
                Row("Mark", "D1",  "Type Name", "Single Flush 900", "Level", "Level 1",  "Area", "1.89",  "Count", "4",  "Comments", "fire rated"),
                Row("Mark", "D10", "Type Name", "Double Glazed 1800", "Level", "Level 10", "Area", "3.24",  "Count", "12", "Comments", ""),
                Row("Mark", "D2",  "Type Name", "Single Flush 900", "Level", "Level 2",  "Area", "1.89",  "Count", "2",  "Comments", "needs check"),
                Row("Mark", "W1",  "Type Name", "Fixed 1200x1500", "Level", "Level 1",  "Area", "1.80",  "Count", "30", "Comments", "fire rated glass"),
                Row("Mark", "W2",  "Type Name", "Casement 600x900", "Level", "Level 2",  "Area", "0.54",  "Count", "7",  "Comments", null),
            };
        }

        public static QuerySchema ElementSchema()
        {
            return QuerySchema.FromColumns(new[] { "Mark", "Type Name", "Level", "Area", "Count", "Comments" });
        }

        public static string[] Marks(QueryResult r)
        {
            return r.Rows.Select(row => row["Mark"].ToString()).ToArray();
        }

        /// <summary>A picker table with a group, like winswitch: windows with optional tmux and claude metadata.</summary>
        public static List<IDictionary<string, object>> Windows()
        {
            return new List<IDictionary<string, object>>
            {
                Row("title", "Alacritty", "workspace", "1", "age", "30s", "tmux.session", "main",  "tmux.title", "build",  "claude.title", "ovh proposal",   "claude.contents", "write the OVH mail"),
                Row("title", "Firefox",   "workspace", "2", "age", "5m"),
                Row("title", "Emacs",     "workspace", "10", "age", "3h", "tmux.session", "notes", "tmux.title", "journal"),
                Row("title", "Alacritty", "workspace", "2", "age", "2d",  "claude.title", "dsl design",   "claude.contents", "ovh is mentioned here too"),
            };
        }

        public static QuerySchema WindowSchema()
        {
            return new QuerySchema()
                .AddField("title").AddField("workspace", false).AddField("age", false)
                .AddGroup("tmux", new[] { "session", "title" }, "session")
                .AddGroup("claude", new[] { "title", "contents" }, "contents");
        }

        public static string[] Titles(QueryResult r)
        {
            return r.Rows.Select(row => row["title"].ToString() + (row.ContainsKey("workspace") ? "@" + row["workspace"] : "")).ToArray();
        }
    }
}
