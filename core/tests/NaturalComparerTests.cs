using System.Collections.Generic;
using System.Linq;
using RevitBallet.Core;
using Xunit;

namespace RevitBallet.Core.Tests
{
    public class NaturalComparerTests
    {
        private readonly NaturalComparer comparer = new NaturalComparer();

        [Theory]
        [InlineData("A2", "A10")]     // numeric segments compare numerically
        [InlineData("2", "10")]       // pure numbers
        [InlineData("Level 1", "Level 2")]
        [InlineData("0.5", "2")]      // decimals
        [InlineData("a1", "B2")]      // case-insensitive text segments
        public void Orders_FirstBeforeSecond(string smaller, string larger)
        {
            Assert.True(comparer.Compare(smaller, larger) < 0);
            Assert.True(comparer.Compare(larger, smaller) > 0);
        }

        [Fact]
        public void Nulls_SortFirst()
        {
            Assert.True(comparer.Compare(null, "a") < 0);
            Assert.Equal(0, comparer.Compare(null, null));
        }

        [Fact]
        public void PlaceholderValues_SortAfterNumbers()
        {
            Assert.True(comparer.Compare("5", "-") < 0);
            Assert.True(comparer.Compare("N/A", "5") > 0);
        }

        [Fact]
        public void SortsSheetNumbersNaturally()
        {
            var input = new List<object> { "A101", "A9", "A10", "A100" };
            input.Sort((a, b) => comparer.Compare(a, b));
            Assert.Equal(new object[] { "A9", "A10", "A100", "A101" }, input.ToArray());
        }
    }

    public class TextMatchTests
    {
        [Theory]
        [InlineData("door-900", "door*", true)]
        [InlineData("door-900", "*900", true)]
        [InlineData("door-900", "d*9*0", true)]
        [InlineData("door-900", "window*", false)]
        [InlineData("Door-900", "door*", true)]    // case-insensitive
        [InlineData("door", "door", true)]         // no wildcard -> contains
        [InlineData("my door", "door", true)]
        public void MatchesGlobPattern(string value, string pattern, bool expected)
        {
            Assert.Equal(expected, TextMatch.MatchesGlobPattern(value, pattern));
        }

        [Theory]
        [InlineData("\"abc\"", "abc")]
        [InlineData("abc", "abc")]
        [InlineData("\"\"", "")]
        [InlineData("\"", "\"")]      // single quote char is not a quoted string
        public void StripQuotes(string input, string expected)
        {
            Assert.Equal(expected, TextMatch.StripQuotes(input));
        }

        [Theory]
        [InlineData("1,234", 1234)]
        [InlineData("$50", 50)]
        [InlineData("50%", 0.5)]
        [InlineData("3.14", 3.14)]
        public void TryParseDouble_HandlesFormats(string input, double expected)
        {
            Assert.True(TextMatch.TryParseDouble(input, out double result));
            Assert.Equal(expected, result, 10);
        }

        [Fact]
        public void TryParseDouble_RejectsText()
        {
            Assert.False(TextMatch.TryParseDouble("abc", out _));
            Assert.False(TextMatch.TryParseDouble("", out _));
        }
    }

    public class NameConventionsTests
    {
        [Theory]
        [InlineData("SwitchView", "switch-view")]
        [InlineData("OpenRvtFilesInNewSessions", "open-rvt-files-in-new-sessions")]
        [InlineData("simple", "simple")]
        [InlineData("", "")]
        public void ConvertToKebabCase(string input, string expected)
        {
            Assert.Equal(expected, NameConventions.ConvertToKebabCase(input));
        }

        [Theory]
        [InlineData("TypeName", "type name")]
        [InlineData("Type_Name", "type  name")] // underscore AND case-break each insert a space
        [InlineData("ElementID", "element id")]
        [InlineData("", "")]
        public void FormatColumnHeader(string input, string expected)
        {
            Assert.Equal(expected, NameConventions.FormatColumnHeader(input));
        }
    }
}
