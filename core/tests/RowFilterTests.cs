using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Core;
using Xunit;

namespace RevitBallet.Core.Tests
{
    public class RowFilterTests
    {
        private static readonly List<string> Props = new List<string> { "Name", "Category", "Area" };

        private static Dictionary<string, object> Row(string name, string category, object area = null)
        {
            return new Dictionary<string, object>
            {
                ["Name"] = name,
                ["Category"] = category,
                ["Area"] = area
            };
        }

        private static List<Dictionary<string, object>> Filter(string query, params Dictionary<string, object>[] rows)
        {
            var ctx = new RowMatchContext { PropertyNames = Props };
            return RowFilter.Filter(rows.ToList(), FilterQueryParser.Parse(query), ctx);
        }

        [Fact]
        public void SubstringTerms_AreAndedAcrossAllColumns()
        {
            var kept = Filter("door 900",
                Row("Door-900x2100", "Doors"),
                Row("Door-800x2100", "Doors"),
                Row("Window-900", "Windows"));

            Assert.Single(kept);
            Assert.Equal("Door-900x2100", kept[0]["Name"]);
        }

        [Fact]
        public void OrGroups_MatchEitherSide()
        {
            var kept = Filter("door || window",
                Row("Door-900", "Doors"),
                Row("Window-600", "Windows"),
                Row("Wall-200", "Walls"));

            Assert.Equal(2, kept.Count);
        }

        [Fact]
        public void Exclusion_RemovesMatchingRows()
        {
            var kept = Filter("!door",
                Row("Door-900", "Doors"),
                Row("Window-600", "Windows"));

            Assert.Single(kept);
            Assert.Equal("Window-600", kept[0]["Name"]);
        }

        [Fact]
        public void MatchingIsCaseInsensitive()
        {
            var kept = Filter("DOOR", Row("door-900", "Doors"));
            Assert.Single(kept);
        }

        [Fact]
        public void ColumnValueFilter_OnlyChecksThatColumn()
        {
            var kept = Filter("$category:doors",
                Row("Door-900", "Doors"),
                Row("Doors are here", "Windows"));

            Assert.Single(kept);
            Assert.Equal("Door-900", kept[0]["Name"]);
        }

        [Fact]
        public void ExactValueFilter_RequiresFullCellMatch()
        {
            var kept = Filter("$category:e\"Door\"",
                Row("a", "Door"),
                Row("b", "Doors"));

            Assert.Single(kept);
            Assert.Equal("a", kept[0]["Name"]);
        }

        [Fact]
        public void GeneralExactFilter_MatchesWholeCellOnly()
        {
            var kept = Filter("e\"Doors\"",
                Row("x", "Doors"),
                Row("y", "Doors and more"));

            Assert.Single(kept);
            Assert.Equal("x", kept[0]["Name"]);
        }

        [Fact]
        public void GlobPattern_MatchesWildcards()
        {
            var kept = Filter("door*2100",
                Row("Door-900x2100", "Doors"),
                Row("Door-900x2000", "Doors"));

            Assert.Single(kept);
            Assert.Equal("Door-900x2100", kept[0]["Name"]);
        }

        [Fact]
        public void Comparison_AllColumns_MatchesNumericCells()
        {
            var kept = Filter(">100",
                Row("a", "Doors", 150),
                Row("b", "Doors", 50));

            Assert.Single(kept);
            Assert.Equal("a", kept[0]["Name"]);
        }

        [Fact]
        public void Comparison_OnColumn_IgnoresOtherColumns()
        {
            var kept = Filter("$area:<100",
                Row("999", "Doors", 50),
                Row("1", "Doors", 200));

            Assert.Single(kept);
            Assert.Equal("999", kept[0]["Name"]);
        }

        [Fact]
        public void SelectionSetFilter_UsesInjectedLookup()
        {
            var rows = new[]
            {
                new Dictionary<string, object> { ["Name"] = "in", ["Id"] = 42L },
                new Dictionary<string, object> { ["Name"] = "out", ["Id"] = 7L }
            };
            var ctx = new RowMatchContext
            {
                PropertyNames = new List<string> { "Name", "Id" },
                SelectionSetLookup = name => name == "temp" ? new HashSet<long> { 42L } : new HashSet<long>()
            };

            var kept = RowFilter.Filter(rows.ToList(), FilterQueryParser.Parse("#temp"), ctx);

            Assert.Single(kept);
            Assert.Equal("in", kept[0]["Name"]);
        }

        [Fact]
        public void SelectionSetFilter_ForeignIdConverter_IsUsedForOpaqueIds()
        {
            var rows = new[]
            {
                new Dictionary<string, object> { ["Name"] = "in", ["Id"] = new OpaqueId(42) }
            };
            var ctx = new RowMatchContext
            {
                PropertyNames = new List<string> { "Name", "Id" },
                SelectionSetLookup = _ => new HashSet<long> { 42L },
                ForeignIdConverter = o => o is OpaqueId oid ? oid.Value : 0
            };

            var kept = RowFilter.Filter(rows.ToList(), FilterQueryParser.Parse("#temp"), ctx);

            Assert.Single(kept);
        }

        [Fact]
        public void SearchIndex_WhenProvided_IsUsedInsteadOfRowValues()
        {
            // Index deliberately disagrees with the row to prove it is consulted
            var rows = new[] { Row("Door-900", "Doors") };
            var ctx = new RowMatchContext
            {
                PropertyNames = Props,
                SearchIndexAllColumns = new Dictionary<int, string> { [0] = "window" }
            };

            var kept = RowFilter.Filter(rows.ToList(), FilterQueryParser.Parse("window"), ctx);

            Assert.Single(kept);
        }

        private class OpaqueId
        {
            public long Value { get; }
            public OpaqueId(long value) { Value = value; }
            public override string ToString() => "opaque"; // not parseable as long
        }
    }
}
