using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Core.Query;
using Xunit;
using static RevitBallet.Core.Tests.QueryTestData;

namespace RevitBallet.Core.Tests
{
    /// <summary>Row filtering, columns and ordering of the picker DSL over fixtures.</summary>
    public class QueryEngineTests
    {
        private static QueryResult Elements(string q) { return QueryEngine.Run(QueryTestData.Elements(), ElementSchema(), q); }
        private static QueryResult Windows(string q) { return QueryEngine.Run(QueryTestData.Windows(), WindowSchema(), q); }

        // ------------------------------------------------------------------ row filters

        [Fact]
        public void EmptyQueryKeepsEverythingInInputOrderWithDefaultColumns()
        {
            var r = Elements("");
            Assert.Equal(new[] { "D1", "D10", "D2", "W1", "W2" }, Marks(r));
            Assert.Equal(new[] { "Mark", "Type Name", "Level", "Area", "Count", "Comments" }, r.Columns.ToArray());
            Assert.Empty(r.FilterFields);
        }

        [Fact]
        public void BareWordsAreSubstringsOverAllFieldsCaseInsensitive()
        {
            Assert.Equal(new[] { "D1", "D2" }, Marks(Elements("single")));
            Assert.Equal(new[] { "D1" }, Marks(Elements("SINGLE fire")));
            Assert.Equal(new[] { "D1", "D2" }, Marks(Elements("flush")));
        }

        [Fact]
        public void BareWordsAndNotOr()
        {
            Assert.Empty(Marks(Elements("single casement")));            // no row has both
            Assert.Equal(new[] { "D1", "D2" }, Marks(Elements("single 900")));
        }

        [Fact]
        public void QuotedPhraseIsALiteralContiguousSubstring()
        {
            Assert.Equal(new[] { "D1", "D2" }, Marks(Elements("\"single flush\"")));
            Assert.Empty(Marks(Elements("\"single 900\"")));              // words present but not contiguous
            Assert.Equal(new[] { "D1" }, Marks(Elements("\"fire rated\" level")).Take(1).ToArray());
        }

        [Fact]
        public void ScopedValueFiltersOneField()
        {
            Assert.Equal(new[] { "D2", "W2" }, Marks(Elements("/fv level:\"level 2\"")));
            Assert.Equal(new[] { "D1", "D10", "W1" }, Marks(Elements("//level \"level 1\"")));        // substring: Level 1 and Level 10
            Assert.Equal(new[] { "D1", "D2" }, Marks(Elements("//type single")));                     // path is a substring of "Type Name"
        }

        [Fact]
        public void PathSegmentsAreSubstringsUnionedAcrossAmbiguousNames()
        {
            // "e" is inside Type Name, Level, Area and Comments: a row survives if any of them holds the value
            Assert.Equal(new[] { "D1", "W1" }, Marks(Elements("//e fire")));       // both in Comments
            Assert.Equal(new[] { "D10" }, Marks(Elements("//e glaz")));            // "Double Glazed" in Type Name
            Assert.Empty(Marks(Elements("//e 30")));                               // 30 is only in Count, whose name has no "e"
            Assert.Equal(new[] { "W1" }, Marks(Elements("//count 30")));
        }

        [Fact]
        public void UnresolvedCompletePathNarrowsToNothing()
        {
            Assert.Empty(Marks(Elements("//zzz foo")));
            Assert.Empty(Marks(Elements("/fv zzz:foo")));
        }

        [Fact]
        public void StillTypingIsInertNeverFlashesToZero()
        {
            Assert.Equal(5, Elements("//").Rows.Count);
            Assert.Equal(5, Elements("//level").Rows.Count);               // flat field, no value yet: no-op
            Assert.Equal(5, Elements("/fv level:").Rows.Count);
            Assert.Equal(5, Elements("/").Rows.Count);
            Assert.Equal(5, Elements("/f").Rows.Count);
            Assert.Equal(5, Elements("/ft ").Rows.Count);
            Assert.Equal(5, Elements("!").Rows.Count);
            Assert.Equal(5, Elements("//level !").Rows.Count);
        }

        [Fact]
        public void ColonFormAndViaFormAgree()
        {
            Assert.Equal(Marks(Elements("/fv level:2")), Marks(Elements("/fv/level 2")));
            Assert.Equal(Marks(Elements("/fv level:2")), Marks(Elements("//level 2")));
        }

        [Fact]
        public void ComparisonsUseNumbersAndNeverFallBackToLiteralText()
        {
            Assert.Equal(new[] { "D10", "W1", "W2" }.OrderBy(x => x).ToArray(), Marks(Elements("//count >5")).OrderBy(x => x).ToArray());
            Assert.Equal(new[] { "D1", "D2", "W1", "W2" }.OrderBy(x => x).ToArray(), Marks(Elements("//area <=1.89")).OrderBy(x => x).ToArray());
            Assert.Equal(new[] { "D10" }, Marks(Elements("//area >3")));
            Assert.Empty(Marks(Elements("//mark >5")));                    // non-numeric: never matches, is not searched as ">5"
            Assert.Equal(new[] { "D1", "D2", "W2" }.OrderBy(x => x).ToArray(), Marks(Elements("//count <10")).OrderBy(x => x).ToArray());
        }

        [Fact]
        public void ComparisonOperatorsAreExact()
        {
            Assert.Equal(new[] { "D1", "D2", "D10", "W1", "W2" }.OrderBy(x => x).ToArray(), Marks(Elements("//count >=2")).OrderBy(x => x).ToArray());
            Assert.Equal(new[] { "D2" }, Marks(Elements("//count <=2")));
            Assert.Equal(new[] { "D2", "D1" }.OrderBy(x => x).ToArray(), Marks(Elements("//count <5")).OrderBy(x => x).ToArray());
        }

        // ------------------------------------------------------------------ negation

        [Theory]
        [InlineData("single")] [InlineData("\"fire rated\"")] [InlineData("//level 1")] [InlineData("/fv level:2")] [InlineData("//count >5")]
        [InlineData("//zzz foo")] [InlineData("//mark >5")] [InlineData("//comments fire")]
        public void NegationIsTheExactComplement(string positive)
        {
            var all = Elements("").Rows.Count;
            var pos = Marks(Elements(positive));
            string neg = positive.StartsWith("/") ? "!" + positive : "!" + positive;
            var negRows = Marks(Elements(neg));
            Assert.Equal(all, pos.Length + negRows.Length);
            Assert.Empty(pos.Intersect(negRows));
        }

        [Fact]
        public void NegationExamples()
        {
            Assert.Equal(new[] { "D10", "W1", "W2" }, Marks(Elements("!single")));
            Assert.Equal(new[] { "D1", "D10", "D2", "W1", "W2" }, Marks(Elements("!zzz")));
            Assert.Equal(new[] { "D10", "D2", "W2" }, Marks(Elements("!//comments \"fire rated\"")));   // empty and missing cells are kept: the positive form did not match them
            Assert.Equal(new[] { "D1", "D2" }, Marks(Elements("//count !>5")));                          // a comparison inverts like any other term
            Assert.Equal(new[] { "D1", "D2" }, Marks(Elements("!//count >5")));
            Assert.Equal(new[] { "D10", "W1", "W2" }, Marks(Elements("!//count <5")));
        }

        [Fact]
        public void NegatedTermsStillAndWithPositiveOnes()
        {
            Assert.Equal(new[] { "D2" }, Marks(Elements("single !//level \"level 1\"")));
        }

        [Fact]
        public void PurelyNegativeQueryShowsEverythingExceptMatches()
        {
            Assert.Equal(new[] { "W1", "W2" }, Marks(Elements("!door !glazed !single")));
        }

        [Fact]
        public void QuotedBangIsLiteral()
        {
            var rows = new List<IDictionary<string, object>> { Row("t", "important"), Row("t", "!important") };
            var s = QuerySchema.FromColumns(new[] { "t" });
            var r = QueryEngine.Run(rows, s, "\"!important\"");
            Assert.Single(r.Rows); Assert.Equal("!important", r.Rows[0]["t"]);
        }

        // ------------------------------------------------------------------ groups

        [Fact]
        public void GroupExistenceFilter()
        {
            Assert.Equal(new[] { "Alacritty@1", "Alacritty@2" }, Titles(Windows("//claude")));
            Assert.Equal(new[] { "Alacritty@1", "Emacs@10" }, Titles(Windows("//tmux")));
            Assert.Equal(new[] { "Firefox@2" }, Titles(Windows("!//claude !//tmux")));
            Assert.Equal(new[] { "Firefox@2", "Emacs@10" }, Titles(Windows("!//claude")));
        }

        [Fact]
        public void BareGroupScopeSearchesEverySubfield()
        {
            // "ovh" is in claude.title on the first window and only in claude.contents on the last: both survive
            Assert.Equal(new[] { "Alacritty@1", "Alacritty@2" }, Titles(Windows("//claude ovh")));
            Assert.Equal(new[] { "Alacritty@1" }, Titles(Windows("//claude.title ovh")));
            Assert.Equal(new[] { "Alacritty@2" }, Titles(Windows("//claude.contents mentioned")));
            Assert.Equal(new[] { "Alacritty@2" }, Titles(Windows("//cla.cont mentioned")));            // each segment a substring
        }

        [Fact]
        public void NegatedGroupScopeDropsARowIfAnySubfieldMatches()
        {
            Assert.Equal(new[] { "Firefox@2", "Emacs@10" }, Titles(Windows("!//claude ovh")));
        }

        [Fact]
        public void FlatPathOnlyIsFreeTextButGroupPathOnlyIsExistence()
        {
            Assert.Equal(new[] { "Emacs@10" }, Titles(Windows("emacs")));                             // flat/none -> free text
            Assert.Equal(new[] { "Alacritty@1", "Alacritty@2" }, Titles(Windows("claude")));          // a bare word naming a group is an existence filter
        }

        [Fact]
        public void ColonSpellingOfGroupScope()
        {
            Assert.Equal(Titles(Windows("//tmux.session main")), Titles(Windows("/fv tmux.session:main")));
            Assert.Equal(new[] { "Alacritty@1" }, Titles(Windows("/fv tmux:main")));                  // bare group, union of subfields
        }

        // ------------------------------------------------------------------ columns

        [Fact]
        public void FilterTypeIntersectsAddTypeAddsRemoveTypeSubtracts()
        {
            Assert.Equal(new[] { "Mark", "Level" }, Elements("/ft mark /at level").Columns.ToArray());
            Assert.Equal(new[] { "Type Name", "Area" }, Elements("/ft \"type name\" /at area").Columns.ToArray());
            Assert.Equal(new[] { "Mark", "Type Name", "Level", "Area", "Count" }, Elements("/rt comments").Columns.ToArray());
            Assert.Equal(new[] { "Mark", "Level" }, Elements("/filter-type mark /add-type level").Columns.ToArray());
        }

        [Fact]
        public void ColumnVerbsAreOrderSensitive()
        {
            // /ft claude /at workspace keeps workspace; /at workspace /ft claude drops it again
            var w = WindowSchema();
            var a = QueryEngine.Run(QueryTestData.Windows(), w, "/ft claude /at workspace").Columns;
            var b = QueryEngine.Run(QueryTestData.Windows(), w, "/at workspace /ft claude").Columns;
            Assert.Equal(new[] { "workspace" }, a.ToArray());                                          // /ft claude over default columns {title} keeps nothing, then workspace is added
            Assert.Empty(b);
        }

        [Fact]
        public void BareGroupAsColumnPathMeansAllSubfields()
        {
            var cols = QueryEngine.Run(QueryTestData.Windows(), WindowSchema(), "/at claude").Columns;
            Assert.Equal(new[] { "title", "claude.title", "claude.contents" }, cols.ToArray());
            Assert.Equal(cols, QueryEngine.Run(QueryTestData.Windows(), WindowSchema(), "/at claude.*").Columns);
            Assert.Equal(new[] { "title", "claude.title" }, QueryEngine.Run(QueryTestData.Windows(), WindowSchema(), "/at claude.title").Columns.ToArray());
        }

        [Fact]
        public void AddingShownOrRemovingHiddenColumnsIsANoOp()
        {
            Assert.Equal(new[] { "Mark", "Type Name", "Level", "Area", "Count", "Comments" }, Elements("/at mark /rt zzz").Columns.ToArray());
        }

        [Fact]
        public void ColumnVerbsNeverChangeRows()
        {
            Assert.Equal(5, Elements("/ft mark /rt level /at count").Rows.Count);
        }

        [Fact]
        public void FilteredFieldsAreAutoShownEvenPastFtAndRt()
        {
            var r = Elements("/ft mark //comments fire");
            Assert.Equal(new[] { "Mark", "Comments" }, r.Columns.ToArray());
            Assert.Equal(new[] { "Comments" }, r.FilterFields.ToArray());
            var r2 = Elements("/rt comments //comments fire");
            Assert.Contains("Comments", r2.Columns);
        }

        [Fact]
        public void PendingViaOnAnUnambiguousFieldAutoShowsButAFragmentDoesNot()
        {
            Assert.Contains("Comments", Elements("//comments").Columns);
            Assert.Equal(new[] { "Mark", "Type Name", "Level", "Area", "Count", "Comments" }, Elements("//c").Columns.ToArray());      // ambiguous fragment: nothing flickers in
            Assert.Equal(new[] { "Mark", "Type Name", "Level", "Area", "Count", "Comments" }, Elements("//zzz").Columns.ToArray());
            Assert.Equal(new[] { "Mark", "Type Name", "Level", "Area", "Count", "Comments" }, Elements("//").Columns.ToArray());
        }

        [Fact]
        public void NegatedFilterFieldsAreAutoShownToo()
        {
            Assert.Contains("claude.contents", QueryEngine.Run(QueryTestData.Windows(), WindowSchema(), "!//claude").Columns);
        }

        [Fact]
        public void FtViaValuePatternDiscoversTheColumnsHoldingThatValue()
        {
            Assert.Equal(new[] { "Type Name" }, Elements("/ft/1200").Columns.ToArray());          // substring, like everywhere
            Assert.Equal(new[] { "Type Name" }, Elements("/ft/*1200*").Columns.ToArray());
            Assert.Empty(Elements("/ft/1200*").Columns);                                           // a * anchors: nothing STARTS with 1200
            Assert.Equal(new[] { "Comments" }, Elements("/ft/fire*").Columns.ToArray());
            Assert.Equal(new[] { "Level" }, Elements("/ft/level*").Columns.ToArray());
            Assert.Equal(new[] { "Area" }, Elements("/ft/1.89").Columns.ToArray());
            // a discovery reset, not an intersect with what was shown
            Assert.Equal(new[] { "Comments" }, Elements("/ft mark /ft/fire*").Columns.ToArray());
        }

        // ------------------------------------------------------------------ order

        [Fact]
        public void SortNumbersNumericallyNotLexically()
        {
            Assert.Equal(new[] { "D2", "D1", "W2", "D10", "W1" }, Marks(Elements("/s count")));
            Assert.Equal(new[] { "W1", "D10", "W2", "D1", "D2" }, Marks(Elements("/s count desc")));
        }

        [Fact]
        public void SortTextNaturally()
        {
            Assert.Equal(new[] { "D1", "D2", "D10", "W1", "W2" }, Marks(Elements("/s mark")));
            Assert.Equal(new[] { "W2", "W1", "D10", "D2", "D1" }, Marks(Elements("/s mark descending")));
        }

        [Fact]
        public void MultiKeySortBreaksTiesWithTheNextKey()
        {
            // area ties: D1 and D2 are both 1.89 -> broken by mark
            Assert.Equal(new[] { "W2", "W1", "D1", "D2", "D10" }, Marks(Elements("/s/area/mark")));
            Assert.Equal(new[] { "D10", "D2", "D1", "W1", "W2" }, Marks(Elements("/s/area/mark desc")));
        }

        [Fact]
        public void SortIsStableWhenKeysTie()
        {
            Assert.Equal(new[] { "W2", "W1", "D1", "D2", "D10" }, Marks(Elements("/s area")));
        }

        [Fact]
        public void LastSortWins()
        {
            Assert.Equal(Marks(Elements("/s count")), Marks(Elements("/s mark /s count")));
        }

        [Fact]
        public void ReverseFlipsTheDefaultOrderAndASortAndIsIdempotent()
        {
            Assert.Equal(new[] { "W2", "W1", "D2", "D10", "D1" }, Marks(Elements("/rv")));
            Assert.Equal(Marks(Elements("/rv")), Marks(Elements("/rv /rv /reverse")));
            Assert.Equal(Marks(Elements("/s count desc")), Marks(Elements("/s count /rv")));
        }

        [Fact]
        public void SortAndFilterAreIndependent()
        {
            Assert.Equal(new[] { "D2", "D1" }, Marks(Elements("single /s count")));
            Assert.Equal(new[] { "D2", "D1" }, Marks(Elements("/s count single")));
            Assert.Empty(Marks(Elements("/s mark blender")));          // `blender` is not a direction, so it is its own filter and matches nothing
        }

        [Fact]
        public void SortNameBlenderSortsAndFilters()
        {
            var rows = new List<IDictionary<string, object>> { Row("name", "b blender"), Row("name", "a blender"), Row("name", "c") };
            var s = QuerySchema.FromColumns(new[] { "name" });
            var r = QueryEngine.Run(rows, s, "/s nam blen");
            Assert.Equal(new[] { "a blender", "b blender" }, r.Rows.Select(x => x["name"].ToString()).ToArray());
        }

        [Fact]
        public void UnresolvableOrAmbiguousSortKeyMakesTheWholeSortInert()
        {
            Assert.Equal(Marks(Elements("")), Marks(Elements("/s zzz")));
            Assert.Equal(Marks(Elements("")), Marks(Elements("/s e")));                 // ambiguous: Type Name, Level, Comments
            Assert.Equal(Marks(Elements("")), Marks(Elements("/s/count/zzz")));         // one bad key voids the chain
        }

        [Fact]
        public void AgeSortsByBucketSecondsWithTheDirectionTrap()
        {
            // ascending = oldest first (chronological), descending = newest first
            Assert.Equal(new[] { "Alacritty@2", "Emacs@10", "Firefox@2", "Alacritty@1" }, Titles(Windows("/s age")));
            Assert.Equal(new[] { "Alacritty@1", "Firefox@2", "Emacs@10", "Alacritty@2" }, Titles(Windows("/s age desc")));
        }

        [Fact]
        public void AgeComparisonFilterComparesTheRawAge()
        {
            Assert.Equal(new[] { "Emacs@10", "Alacritty@2" }, Titles(Windows("//age >1h")));          // older than 1h
            Assert.Equal(new[] { "Alacritty@1", "Firefox@2" }, Titles(Windows("//age <1h")));
        }

        [Fact]
        public void GroupAsSortKeyMeansItsDefaultSubfieldAndStarIsRejected()
        {
            var s = WindowSchema();
            var viaGroup = QueryEngine.Run(QueryTestData.Windows(), s, "/s claude");        // default sub claude.contents
            Assert.Equal(QueryEngine.Run(QueryTestData.Windows(), s, "/s claude.contents").RowIndices, viaGroup.RowIndices);
            Assert.Equal(new[] { 0, 1, 2, 3 }, QueryEngine.Run(QueryTestData.Windows(), s, "/s claude.*").RowIndices.ToArray());
        }

        [Fact]
        public void RowIndicesPointBackAtTheInputRows()
        {
            var r = Elements("single /s count");
            Assert.Equal(new[] { 2, 0 }, r.RowIndices.ToArray());
            Assert.Equal(new[] { "D2", "D1" }, r.Rows.Select(row => row["Mark"].ToString()).ToArray());
        }
    }
}
