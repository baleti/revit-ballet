using System.Collections.Generic;
using System.Linq;
using RevitBallet.Core.Query;
using Xunit;
using static RevitBallet.Core.Tests.QueryTestData;

namespace RevitBallet.Core.Tests
{
    /// <summary>
    /// The old DataGrid language (FilterQueryParser/RowFilter) against its canonical replacement. Every old feature that has a
    /// canonical spelling must keep returning exactly the same rows; the intentional differences are pinned down explicitly below.
    /// </summary>
    public class QueryMigrationTests
    {
        private static string[] OldRows(string oldQuery)
        {
            var entries = Elements().Select(r => new Dictionary<string, object>(r)).ToList();
            var ctx = new RowMatchContext { PropertyNames = ElementSchema().DefaultColumns };
            return RowFilter.Filter(entries, FilterQueryParser.Parse(oldQuery), ctx).Select(r => r["Mark"].ToString()).ToArray();
        }

        private static string[] NewRows(string newQuery)
        {
            return Marks(QueryEngine.Run(Elements(), ElementSchema(), newQuery));
        }

        [Theory]
        // plain terms
        [InlineData("single", "single")]
        [InlineData("single 900", "single 900")]
        [InlineData("SINGLE", "single")]
        [InlineData("zzz", "zzz")]
        // exclusion
        [InlineData("!single", "!single")]
        [InlineData("!door !glazed", "!door !glazed")]
        [InlineData("fire !glass", "fire !glass")]
        // quoted
        [InlineData("\"single flush\"", "\"single flush\"")]
        [InlineData("\"fire rated\"", "\"fire rated\"")]
        // column-scoped value
        [InlineData("$level:2", "//level 2")]
        [InlineData("$level:2", "/fv level:2")]
        [InlineData("$level:2", "/fv/level 2")]
        [InlineData("$mark:d", "//mark d")]
        [InlineData("$\"type name\":flush", "//\"type name\" flush")]
        [InlineData("$comments:fire", "//comments fire")]
        [InlineData("$comments::fire", "//comments fire")]
        [InlineData("!$comments:fire", "!//comments fire")]
        [InlineData("$level:2 single", "//level 2 single")]
        [InlineData("$level:1 !$comments:fire", "//level 1 !//comments fire")]
        // numeric comparison, column-scoped
        [InlineData("$area:>1.8", "//area >1.8")]
        [InlineData("$area:<2", "//area <2")]
        [InlineData("$count:>5", "//count >5")]
        [InlineData("$count:<10", "//count <10")]
        [InlineData("!$count:>5", "!//count >5")]
        public void RowsAreIdenticalForEveryOldFeatureWithACanonicalSpelling(string oldQuery, string newQuery)
        {
            Assert.Equal(OldRows(oldQuery), NewRows(newQuery));
        }

        [Theory]
        [InlineData("$mark", "/ft mark")]
        [InlineData("$mark $level", "/ft mark /at level")]
        [InlineData("$\"type name\"", "/ft \"type name\"")]
        [InlineData("$count $area $comments", "/ft count /at area /at comments")]
        public void ColumnVisibilityIsTheSameSetOldAndNew(string oldQuery, string newQuery)
        {
            var groups = FilterQueryParser.Parse(oldQuery);
            var oldVisible = ElementSchema().DefaultColumns
                .Where(c => groups.SelectMany(g => g.ColVisibilityFilters).Any(parts => parts.All(p => c.ToLowerInvariant().Contains(p))))
                .OrderBy(c => c).ToArray();
            var newVisible = QueryEngine.Run(Elements(), ElementSchema(), newQuery).Columns.OrderBy(c => c).ToArray();
            Assert.Equal(oldVisible, newVisible);
        }

        [Fact]
        public void ScopedValueFilterStillShowsItsColumnLikeTheOldFormDid()
        {
            // old: `$comments:fire` hid every other column; new: the column is auto-shown next to whatever else is shown
            var r = QueryEngine.Run(Elements(), ElementSchema(), "/ft mark //comments fire");
            Assert.Equal(new[] { "Mark", "Comments" }, r.Columns.ToArray());
        }

        // ---- intentional differences -------------------------------------------------------------------------------------

        [Fact]
        public void UnresolvedColumnNowNarrowsToNothingInsteadOfBeingIgnored()
        {
            Assert.Equal(5, OldRows("$zzz:foo").Length);          // old: a column that matches nothing was skipped
            Assert.Empty(NewRows("//zzz foo"));                    // canonical: a complete path that resolves to nothing is an honest empty result
        }

        [Fact]
        public void NegatedQuotedFormsWorkNowTheOldTokenizerSplitThemIntoBrokenHalves()
        {
            // old regex had no alternative for !"..." or !$col:"..." so `!"fire rated"` became two junk words and matched nothing
            Assert.Empty(OldRows("!\"fire rated\""));
            Assert.Equal(new[] { "D10", "D2", "W2" }, NewRows("!\"fire rated\""));
            Assert.Empty(OldRows("!$level:\"level 1\""));
            Assert.Equal(new[] { "D2", "W2" }, NewRows("!//level \"level 1\""));
        }

        [Fact]
        public void OrGroupsAreNotPartOfTheCanonicalGrammar()
        {
            Assert.Equal(new[] { "D1", "D2", "W2" }.OrderBy(x => x).ToArray(), OldRows("single || casement").OrderBy(x => x).ToArray());
            Assert.Empty(NewRows("single || casement"));           // now three ANDed words, one of them the literal "||"
        }

        [Fact]
        public void ExactAndGlobSpellingsAreNotPartOfTheCanonicalGrammar()
        {
            Assert.Equal(new[] { "D1", "D2" }, OldRows("e\"single flush 900\""));
            Assert.Empty(NewRows("e\"single flush 900\""));            // no exact-match form exists: this is just the text esingle flush 900, which no row contains
            Assert.Equal(new[] { "D1", "D2" }, NewRows("\"single flush 900\""));    // the canonical spelling of "contains this exact run"
        }

        [Fact]
        public void BareNumericComparisonOverEveryColumnHasNoCanonicalForm()
        {
            Assert.Equal(new[] { "D10", "W1", "W2" }, OldRows(">5"));
            // new: `>5` is a plain word searched as text; the scoped spelling (`//count >5`) is the replacement
            Assert.Empty(NewRows(">5"));
        }

        [Fact]
        public void ColumnPositionsAreReplacedByTheOrderOfAddType()
        {
            var cols = QueryEngine.Run(Elements(), ElementSchema(), "/ft count /at mark").Columns;
            Assert.Equal(new[] { "Count", "Mark" }, cols.ToArray());      // old `1$count 2$mark`
        }
    }
}
