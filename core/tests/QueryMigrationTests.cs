using System.Linq;
using RevitBallet.Core.Query;
using Xunit;
using static RevitBallet.Core.Tests.QueryTestData;

namespace RevitBallet.Core.Tests
{
    /// <summary>
    /// The old DataGrid search language against its canonical replacement. The rows in the table are what the OLD engine
    /// (FilterQueryParser/RowFilter, removed when the DataGrid moved to the canonical DSL) returned for each old spelling,
    /// recorded before it was deleted; the new spelling must keep returning exactly them. The intentional differences follow below.
    /// </summary>
    public class QueryMigrationTests
    {
        private static string[] NewRows(string query)
        {
            return Marks(QueryEngine.Run(Elements(), ElementSchema(), query));
        }

        [Theory]
        [InlineData("single", "single", "D1,D2")]
        [InlineData("single 900", "single 900", "D1,D2")]
        [InlineData("SINGLE", "single", "D1,D2")]
        [InlineData("zzz", "zzz", "")]
        [InlineData("!single", "!single", "D10,W1,W2")]
        [InlineData("!door !glazed", "!door !glazed", "D1,D2,W1,W2")]
        [InlineData("fire !glass", "fire !glass", "D1")]
        [InlineData("\"single flush\"", "\"single flush\"", "D1,D2")]
        [InlineData("\"fire rated\"", "\"fire rated\"", "D1,W1")]
        [InlineData("$level:2", "//level 2", "D2,W2")]
        [InlineData("$level:2", "/fv level:2", "D2,W2")]
        [InlineData("$level:2", "/fv/level 2", "D2,W2")]
        [InlineData("$mark:d", "//mark d", "D1,D10,D2")]
        [InlineData("$\"type name\":flush", "//\"type name\" flush", "D1,D2")]
        [InlineData("$comments:fire", "//comments fire", "D1,W1")]
        [InlineData("$comments::fire", "//comments fire", "D1,W1")]
        [InlineData("!$comments:fire", "!//comments fire", "D10,D2,W2")]
        [InlineData("$level:2 single", "//level 2 single", "D2")]
        [InlineData("$level:1 !$comments:fire", "//level 1 !//comments fire", "D10")]
        [InlineData("$area:>1.8", "//area >1.8", "D1,D10,D2")]
        [InlineData("$area:<2", "//area <2", "D1,D2,W1,W2")]
        [InlineData("$count:>5", "//count >5", "D10,W1,W2")]
        [InlineData("$count:<10", "//count <10", "D1,D2,W2")]
        [InlineData("!$count:>5", "!//count >5", "D1,D2")]
        public void EveryOldFeatureWithACanonicalSpellingReturnsTheRowsItAlwaysDid(string oldSpelling, string newSpelling, string expectedRows)
        {
            string[] expected = expectedRows.Length == 0 ? new string[0] : expectedRows.Split(',');
            Assert.Equal(expected, NewRows(newSpelling));
        }

        [Theory]
        [InlineData("$mark", "/ft mark", "Mark")]
        [InlineData("$mark $level", "/ft mark /at level", "Level,Mark")]
        [InlineData("$\"type name\"", "/ft \"type name\"", "Type Name")]
        [InlineData("$count $area $comments", "/ft count /at area /at comments", "Area,Comments,Count")]
        public void ColumnVisibilityIsTheSameSetOldAndNew(string oldSpelling, string newSpelling, string visibleColumns)
        {
            var cols = QueryEngine.Run(Elements(), ElementSchema(), newSpelling).Columns.OrderBy(c => c).ToArray();
            Assert.Equal(visibleColumns.Split(','), cols);
        }

        [Fact]
        public void ScopedValueFilterStillShowsItsColumnLikeTheOldFormDid()
        {
            // old: `$comments:fire` hid every other column; new: the column is auto-shown next to whatever else is shown
            var r = QueryEngine.Run(Elements(), ElementSchema(), "/ft mark //comments fire");
            Assert.Equal(new[] { "Mark", "Comments" }, r.Columns.ToArray());
        }

        // ---- intentional differences ----------------------------------------------------------------------------------

        [Fact]
        public void UnresolvedColumnNowNarrowsToNothingInsteadOfBeingIgnored()
        {
            // old: `$zzz:foo` matched all 5 rows (a column that matched nothing was skipped); canonical: an honest empty result
            Assert.Empty(NewRows("//zzz foo"));
        }

        [Fact]
        public void NegatedQuotedFormsWorkNowTheOldTokenizerSplitThemIntoBrokenHalves()
        {
            // old regex had no alternative for !"..." or !$col:"..." so `!"fire rated"` became two junk words and matched nothing at all
            Assert.Equal(new[] { "D10", "D2", "W2" }, NewRows("!\"fire rated\""));
            Assert.Equal(new[] { "D2", "W2" }, NewRows("!//level \"level 1\""));
        }

        [Fact]
        public void OrGroupsAreNotPartOfTheCanonicalGrammar()
        {
            // old: `single || casement` = D1,D2,W2; now three ANDed words, one of them the literal "||"
            Assert.Empty(NewRows("single || casement"));
        }

        [Fact]
        public void ExactAndGlobSpellingsAreNotPartOfTheCanonicalGrammar()
        {
            // old: e"single flush 900" matched D1,D2 by exact cell value; there is no exact-match form, a quoted phrase is "contains this run"
            Assert.Empty(NewRows("e\"single flush 900\""));
            Assert.Equal(new[] { "D1", "D2" }, NewRows("\"single flush 900\""));
        }

        [Fact]
        public void BareNumericComparisonOverEveryColumnHasNoCanonicalForm()
        {
            // old: `>5` = D10,W1,W2 (any numeric column above 5). New: a plain word searched as text; `//count >5` is the scoped replacement
            Assert.Empty(NewRows(">5"));
            Assert.Equal(new[] { "D10", "W1", "W2" }, NewRows("//count >5"));
        }

        [Fact]
        public void ColumnPositionsAreReplacedByTheOrderOfAddType()
        {
            var cols = QueryEngine.Run(Elements(), ElementSchema(), "/ft count /at mark").Columns;
            Assert.Equal(new[] { "Count", "Mark" }, cols.ToArray());      // old `1$count 2$mark`
        }
    }
}
