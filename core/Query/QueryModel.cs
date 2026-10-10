using System.Collections.Generic;

namespace RevitBallet.Core.Query
{
    /// <summary>The six verbs of the picker query DSL (see docs/remote-commands-and-query-dsl.md and ~/.config/docs/query-dsl.md).</summary>
    public enum Verb { FilterValue, FilterType, AddType, RemoveType, Sort, Reverse }

    /// <summary>The shapes a row-filter term can take. Which of them is an existence filter, a free-text word or a no-op
    /// depends on the schema, so the parser only records the spelling and the executor decides.</summary>
    public enum TermForm
    {
        /// <summary>Free text over the haystack: a bare word, a "quoted phrase" or `/fv text` when the text is quoted.</summary>
        Text,
        /// <summary>`/fv path:value` (or a bare `path:value`). An empty value means the value is still being typed.</summary>
        Scoped,
        /// <summary>`/fv path` / a bare word without a colon: an existence filter when the path is a group, free text otherwise.</summary>
        PathOnly,
        /// <summary>`/fv/path [value]` and `//path [value]`. No value yet: a group is an existence filter, a flat field is a no-op.</summary>
        Via
    }

    /// <summary>One row filter. Negation is the exact complement of the positive form.</summary>
    public sealed class RowTerm
    {
        public TermForm Form;
        public bool Negated;
        /// <summary>Text: the text. PathOnly: the path, which is also the free-text fallback.</summary>
        public string Text;
        public string Path;
        public string Value;
        public bool HasValue;
    }

    /// <summary>One column verb (`/ft`, `/at`, `/rt`), applied left to right.</summary>
    public sealed class ColumnOp
    {
        public Verb Verb;
        public string Path;
        /// <summary>Written `/ft/path`: only the via form falls back to a value pattern when the path resolves to no field.</summary>
        public bool Via;
    }

    /// <summary>The effective `/sort`: the last one in the query wins.</summary>
    public sealed class SortSpec
    {
        public List<string> Paths = new List<string>();
        public bool Descending;
    }

    /// <summary>A parsed query: an order-independent AND of row terms, a left-to-right column pipeline, one sort, one reverse flag.</summary>
    public sealed class QueryPlan
    {
        public List<RowTerm> RowTerms = new List<RowTerm>();
        public List<ColumnOp> ColumnOps = new List<ColumnOp>();
        public SortSpec Sort;
        public bool Reverse;
    }

    /// <summary>What running a plan over rows produced.</summary>
    public sealed class QueryResult
    {
        /// <summary>The surviving rows, in display order.</summary>
        public List<IDictionary<string, object>> Rows = new List<IDictionary<string, object>>();
        /// <summary>Index of each surviving row in the input list, same order as Rows.</summary>
        public List<int> RowIndices = new List<int>();
        /// <summary>The columns to show, in order (full field names, `group.sub` for a subfield).</summary>
        public List<string> Columns = new List<string>();
        /// <summary>The fields a filter or existence term refers to; they are already part of Columns (auto-shown filter fields).</summary>
        public List<string> FilterFields = new List<string>();
    }
}
