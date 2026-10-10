using System.Linq;
using RevitBallet.Core.Query;
using Xunit;

namespace RevitBallet.Core.Tests
{
    /// <summary>The grammar of the picker DSL (~/.config/docs/query-dsl.md): tokens, verbs, via paths, arity, negation, quoting, inert input.</summary>
    public class QueryParserTests
    {
        [Fact]
        public void EmptyQueryIsAnEmptyPlan()
        {
            var p = QueryParser.Parse("");
            Assert.Empty(p.RowTerms); Assert.Empty(p.ColumnOps); Assert.Null(p.Sort); Assert.False(p.Reverse);
            Assert.Empty(QueryParser.Parse(null).RowTerms);
            Assert.Empty(QueryParser.Parse("   ").RowTerms);
        }

        [Fact]
        public void BareWordsAreFilterValueTermsAndAnd()
        {
            var p = QueryParser.Parse("foo bar");
            Assert.Equal(2, p.RowTerms.Count);
            Assert.All(p.RowTerms, t => Assert.Equal(TermForm.PathOnly, t.Form));
            Assert.Equal(new[] { "foo", "bar" }, p.RowTerms.Select(t => t.Text).ToArray());
        }

        [Fact]
        public void FvAndBareWordAreTheSameThing()
        {
            var a = QueryParser.Parse("/fv foo").RowTerms.Single();
            var b = QueryParser.Parse("foo").RowTerms.Single();
            Assert.Equal(b.Form, a.Form); Assert.Equal(b.Text, a.Text); Assert.Equal(b.Negated, a.Negated);
            var c = QueryParser.Parse("/filter-value foo").RowTerms.Single();
            Assert.Equal(b.Text, c.Text);
        }

        [Theory]
        [InlineData("/ft", Verb.FilterType)] [InlineData("/filter-type", Verb.FilterType)]
        [InlineData("/at", Verb.AddType)] [InlineData("/add-type", Verb.AddType)]
        [InlineData("/rt", Verb.RemoveType)] [InlineData("/remove-type", Verb.RemoveType)]
        public void ColumnVerbsShortAndLong(string verb, Verb expected)
        {
            var p = QueryParser.Parse(verb + " name");
            var op = p.ColumnOps.Single();
            Assert.Equal(expected, op.Verb); Assert.Equal("name", op.Path); Assert.False(op.Via);
        }

        [Theory]
        [InlineData("/FV foo")] [InlineData("/Fv foo")] [InlineData("/fvx foo")] [InlineData("/filter foo")] [InlineData("/f foo")]
        public void VerbNamesAreExactMatchedNoFuzzyNoCase(string q)
        {
            var p = QueryParser.Parse(q);
            // the unknown verb is inert; `foo` falls through as an ordinary term
            Assert.Equal("foo", p.RowTerms.Single().Text);
            Assert.Empty(p.ColumnOps);
        }

        [Theory]
        [InlineData("/")] [InlineData("/f")] [InlineData("/xyz")] [InlineData("/ft")] [InlineData("/s")] [InlineData("!")] [InlineData("!!")]
        public void HalfTypedTokensAreInert(string q)
        {
            var p = QueryParser.Parse(q);
            Assert.Empty(p.RowTerms.Where(t => t.Form == TermForm.Text));
            Assert.Null(p.Sort);
            Assert.Empty(p.ColumnOps);
        }

        [Fact]
        public void ViaFormMeansTheSameAsTheColonForm()
        {
            var via = QueryParser.Parse("/fv/title foo").RowTerms.Single();
            Assert.Equal(TermForm.Via, via.Form); Assert.Equal("title", via.Path); Assert.Equal("foo", via.Value); Assert.True(via.HasValue);
            var colon = QueryParser.Parse("/fv title:foo").RowTerms.Single();
            Assert.Equal(TermForm.Scoped, colon.Form); Assert.Equal("title", colon.Path); Assert.Equal("foo", colon.Value);
        }

        [Fact]
        public void DoubleSlashIsFvVia()
        {
            var t = QueryParser.Parse("//claude.title ovh").RowTerms.Single();
            Assert.Equal(TermForm.Via, t.Form); Assert.Equal("claude.title", t.Path); Assert.Equal("ovh", t.Value);
            var lone = QueryParser.Parse("//").RowTerms.Single();
            Assert.Equal("", lone.Path); Assert.False(lone.HasValue);
        }

        [Fact]
        public void ViaWithNoValueYetIsPending()
        {
            var t = QueryParser.Parse("//tokens").RowTerms.Single();
            Assert.False(t.HasValue);
            var u = QueryParser.Parse("//tokens /s name").RowTerms.Single();     // the next verb is never swallowed as the value
            Assert.False(u.HasValue);
        }

        [Fact]
        public void AViaValueOwnsTheNextTokenEvenIfItLooksLikeAPath()
        {
            var t = QueryParser.Parse("//path /etc").RowTerms.Single();
            Assert.True(t.HasValue); Assert.Equal("/etc", t.Value);
        }

        [Fact]
        public void ArityIsACeilingTokensPastTheVerbBecomeFilters()
        {
            var p = QueryParser.Parse("/ft name blender /fv x");
            Assert.Equal("name", p.ColumnOps.Single().Path);
            Assert.Equal(new[] { "blender", "x" }, p.RowTerms.Select(t => t.Text).ToArray());
        }

        [Fact]
        public void SortTakesAConditionalDirectionAndOtherwiseLeavesTheToken()
        {
            var p = QueryParser.Parse("/sort age descending");
            Assert.Equal(new[] { "age" }, p.Sort.Paths.ToArray()); Assert.True(p.Sort.Descending);

            var q = QueryParser.Parse("/s name blender");
            Assert.Equal(new[] { "name" }, q.Sort.Paths.ToArray()); Assert.False(q.Sort.Descending);
            Assert.Equal("blender", q.RowTerms.Single().Text);          // not swallowed as a rejected direction
        }

        [Theory]
        [InlineData("asc", false)] [InlineData("a", false)] [InlineData("ascending", false)]
        [InlineData("desc", true)] [InlineData("de", true)] [InlineData("d", true)] [InlineData("descending", true)]
        public void DirectionWords(string word, bool descending)
        {
            var p = QueryParser.Parse("/s name " + word);
            Assert.Equal(descending, p.Sort.Descending);
            Assert.Empty(p.RowTerms);
        }

        [Fact]
        public void ViaSortChainsKeys()
        {
            var p = QueryParser.Parse("/sort/tokens/title desc");
            Assert.Equal(new[] { "tokens", "title" }, p.Sort.Paths.ToArray()); Assert.True(p.Sort.Descending);
        }

        [Fact]
        public void LastSortWinsAndReverseIsIdempotent()
        {
            var p = QueryParser.Parse("/s a /s b desc /rv /rv");
            Assert.Equal(new[] { "b" }, p.Sort.Paths.ToArray()); Assert.True(p.Sort.Descending);
            Assert.True(p.Reverse);
        }

        [Fact]
        public void NegationOnRowFiltersAllSpellings()
        {
            Assert.True(QueryParser.Parse("!foo").RowTerms.Single().Negated);
            Assert.True(QueryParser.Parse("!\"foo bar\"").RowTerms.Single().Negated);
            Assert.True(QueryParser.Parse("!//path x").RowTerms.Single().Negated);
            Assert.True(QueryParser.Parse("!/fv/path x").RowTerms.Single().Negated);
            Assert.True(QueryParser.Parse("!/fv path:x").RowTerms.Single().Negated);
            Assert.True(QueryParser.Parse("//path !x").RowTerms.Single().Negated);
            Assert.True(QueryParser.Parse("!//path !x").RowTerms.Single().Negated);     // both is still one negation
            Assert.False(QueryParser.Parse("a!b").RowTerms.Single().Negated);           // ! must lead the token
        }

        [Fact]
        public void NegationOnColumnAndOrderVerbsIsInertButConsumesItsArguments()
        {
            var p = QueryParser.Parse("!/ft name !/s age !/rv");
            Assert.Empty(p.ColumnOps); Assert.Null(p.Sort); Assert.False(p.Reverse); Assert.Empty(p.RowTerms);
        }

        [Fact]
        public void QuotedTokensAreNeverCommands()
        {
            var p = QueryParser.Parse("\"//foo\" \"/fv\"");
            Assert.Equal(new[] { "//foo", "/fv" }, p.RowTerms.Select(t => t.Text).ToArray());
            Assert.All(p.RowTerms, t => Assert.Equal(TermForm.Text, t.Form));
        }

        [Fact]
        public void QuotedValueKeepsWhitespaceAndColonsInsideQuotes()
        {
            var t = QueryParser.Parse("/fv title:\"imperial rome\"").RowTerms.Single();
            Assert.Equal("title", t.Path); Assert.Equal("imperial rome", t.Value);
            var u = QueryParser.Parse("//title \"a b\"").RowTerms.Single();
            Assert.Equal("a b", u.Value);
            var v = QueryParser.Parse("\"10:30\"").RowTerms.Single();
            Assert.Equal(TermForm.Text, v.Form); Assert.Equal("10:30", v.Text);
        }

        [Fact]
        public void UnterminatedQuoteRunsToTheEnd()
        {
            var p = QueryParser.Parse("foo \"bar baz");
            Assert.Equal(2, p.RowTerms.Count);
            Assert.Equal("bar baz", p.RowTerms[1].Text);
        }

        [Fact]
        public void ComparisonValueSpellings()
        {
            var t = QueryParser.Parse("//area >1.5").RowTerms.Single();
            Assert.Equal(">1.5", t.Value);
            var u = QueryParser.Parse("/fv area:<=2").RowTerms.Single();
            Assert.Equal("<=2", u.Value);
        }

        [Fact]
        public void ColumnVerbViaForms()
        {
            var p = QueryParser.Parse("/ft/claude /at/claude.* /rt/tmux.session");
            Assert.Equal(new[] { "claude", "claude.*", "tmux.session" }, p.ColumnOps.Select(o => o.Path).ToArray());
            Assert.All(p.ColumnOps, o => Assert.True(o.Via));
        }

        [Fact]
        public void SplitRespectsQuotesMidToken()
        {
            Assert.Equal(new[] { "a", "title:\"x y\"", "b" }, QueryParser.Split("a title:\"x y\" b").ToArray());
        }

        [Fact]
        public void ReverseWithAViaPathIsNotAVerb()
        {
            var p = QueryParser.Parse("/rv/x");
            Assert.False(p.Reverse);
        }
    }
}
