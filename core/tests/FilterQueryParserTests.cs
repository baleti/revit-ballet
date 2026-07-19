using System.Collections.Generic;
using System.Linq;
using RevitBallet.Core;
using Xunit;

namespace RevitBallet.Core.Tests
{
    public class FilterQueryParserTests
    {
        [Fact]
        public void PlainTerms_BecomeGeneralFilters_Lowercased()
        {
            var groups = FilterQueryParser.Parse("Door 900");

            var g = Assert.Single(groups);
            Assert.Equal(new[] { "door", "900" }, g.GeneralFilters);
        }

        [Fact]
        public void OrOperator_SplitsIntoMultipleGroups()
        {
            var groups = FilterQueryParser.Parse("door || window");

            Assert.Equal(2, groups.Count);
            Assert.Equal("door", Assert.Single(groups[0].GeneralFilters));
            Assert.Equal("window", Assert.Single(groups[1].GeneralFilters));
        }

        [Fact]
        public void OrOperator_InsideQuotes_IsNotASeparator()
        {
            var groups = FilterQueryParser.Parse("\"a || b\"");

            var g = Assert.Single(groups);
            Assert.Equal("a || b", Assert.Single(g.GeneralFilters));
        }

        [Fact]
        public void Exclusion_PrefixesGeneralFilterWithBang()
        {
            var groups = FilterQueryParser.Parse("!door");

            var g = Assert.Single(groups);
            Assert.Equal("!door", Assert.Single(g.GeneralFilters));
        }

        [Fact]
        public void QuotedTerm_KeepsSpaces()
        {
            var groups = FilterQueryParser.Parse("\"fire door\"");

            var g = Assert.Single(groups);
            Assert.Equal("fire door", Assert.Single(g.GeneralFilters));
        }

        [Fact]
        public void ExactTerm_GoesToGeneralExactFilters()
        {
            var groups = FilterQueryParser.Parse("e\"Door-900\"");

            var g = Assert.Single(groups);
            Assert.Equal("door-900", Assert.Single(g.GeneralExactFilters));
            Assert.Empty(g.GeneralFilters);
        }

        [Fact]
        public void GlobTerm_GoesToGlobPatterns()
        {
            var groups = FilterQueryParser.Parse("door*900");

            var g = Assert.Single(groups);
            Assert.Equal("door*900", Assert.Single(g.GeneralGlobPatterns));
            Assert.Empty(g.GeneralFilters);
        }

        [Fact]
        public void ColumnValueFilter_ParsesColumnAndValue()
        {
            var groups = FilterQueryParser.Parse("$category:doors");

            var g = Assert.Single(groups);
            var f = Assert.Single(g.ColValueFilters);
            Assert.Equal(new[] { "category" }, f.ColumnParts);
            Assert.Equal("doors", f.Value);
            Assert.False(f.IsExclusion);
        }

        [Fact]
        public void ColumnOnly_AddsVisibilityFilterWithoutValueFilter()
        {
            var groups = FilterQueryParser.Parse("$category");

            var g = Assert.Single(groups);
            Assert.Single(g.ColVisibilityFilters);
            Assert.Empty(g.ColValueFilters);
        }

        [Fact]
        public void QuotedColumn_SplitsIntoParts()
        {
            var groups = FilterQueryParser.Parse("$\"type name\":900");

            var g = Assert.Single(groups);
            var f = Assert.Single(g.ColValueFilters);
            Assert.Equal(new[] { "type", "name" }, f.ColumnParts);
            Assert.Equal("900", f.Value);
        }

        [Fact]
        public void DoubleColon_SeparatesValueContainingColon()
        {
            var groups = FilterQueryParser.Parse("$mark::A:1");

            var g = Assert.Single(groups);
            var f = Assert.Single(g.ColValueFilters);
            Assert.Equal("a:1", f.Value);
        }

        [Fact]
        public void ColumnExclusion_SetsIsExclusion()
        {
            var groups = FilterQueryParser.Parse("!$category:doors");

            var g = Assert.Single(groups);
            var f = Assert.Single(g.ColValueFilters);
            Assert.True(f.IsExclusion);
        }

        [Fact]
        public void StandaloneComparison_ParsesOperatorAndValue()
        {
            var groups = FilterQueryParser.Parse(">50");

            var g = Assert.Single(groups);
            var f = Assert.Single(g.ComparisonFilters);
            Assert.Equal(ComparisonOperator.GreaterThan, f.Operator);
            Assert.Equal(50, f.Value);
            Assert.Null(f.ColumnParts);
        }

        [Fact]
        public void ColumnComparison_TargetsColumn()
        {
            var groups = FilterQueryParser.Parse("$area:<100");

            var g = Assert.Single(groups);
            var f = Assert.Single(g.ComparisonFilters);
            Assert.Equal(ComparisonOperator.LessThan, f.Operator);
            Assert.Equal(100, f.Value);
            Assert.Equal(new[] { "area" }, f.ColumnParts);
        }

        [Fact]
        public void SelectionSetFilter_Unquoted()
        {
            var groups = FilterQueryParser.Parse("#temp");

            var g = Assert.Single(groups);
            var f = Assert.Single(g.SelectionSetFilters);
            Assert.Equal("temp", f.SelectionSetName);
            Assert.False(f.IsExclusion);
        }

        [Fact]
        public void SelectionSetFilter_QuotedWithSpacesAndExclusion()
        {
            var groups = FilterQueryParser.Parse("!#\"my set\"");

            var g = Assert.Single(groups);
            var f = Assert.Single(g.SelectionSetFilters);
            Assert.Equal("my set", f.SelectionSetName);
            Assert.True(f.IsExclusion);
        }

        [Fact]
        public void NumericPrefix_CreatesColumnOrdering()
        {
            var groups = FilterQueryParser.Parse("2$category");

            var g = Assert.Single(groups);
            var o = Assert.Single(g.ColumnOrdering);
            Assert.Equal(2, o.Position);
            Assert.Equal(new[] { "category" }, o.ColumnParts);
        }

        [Fact]
        public void ExactColumn_SetsExactMatchFlags()
        {
            var groups = FilterQueryParser.Parse("$e\"category\"::doors");

            var g = Assert.Single(groups);
            var f = Assert.Single(g.ColValueFilters);
            Assert.True(f.IsColumnExactMatch);
            Assert.False(f.IsExactMatch);
        }

        [Fact]
        public void ExactValue_SetsValueExactMatch()
        {
            var groups = FilterQueryParser.Parse("$category:e\"Doors\"");

            var g = Assert.Single(groups);
            var f = Assert.Single(g.ColValueFilters);
            Assert.True(f.IsExactMatch);
            Assert.Equal("doors", f.Value);
        }

        [Fact]
        public void MixedQuery_ParsesAllParts()
        {
            var groups = FilterQueryParser.Parse("door $category:doors !steel >10 #temp");

            var g = Assert.Single(groups);
            Assert.Contains("door", g.GeneralFilters);
            Assert.Contains("!steel", g.GeneralFilters);
            Assert.Single(g.ColValueFilters);
            Assert.Single(g.ComparisonFilters);
            Assert.Single(g.SelectionSetFilters);
        }
    }
}
