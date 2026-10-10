using System;
using System.Collections.Generic;
using System.Text;

namespace RevitBallet.Core.Query
{
    /// <summary>
    /// Parses the picker DSL (~/.config/docs/query-dsl.md is the canonical definition) into a <see cref="QueryPlan"/>.
    /// The parser never fails: anything half typed is inert and contributes nothing.
    /// </summary>
    public static class QueryParser
    {
        private static readonly Dictionary<string, Verb> VerbNames = new Dictionary<string, Verb>(StringComparer.Ordinal)
        {
            { "fv", Verb.FilterValue }, { "filter-value", Verb.FilterValue },
            { "ft", Verb.FilterType }, { "filter-type", Verb.FilterType },
            { "at", Verb.AddType }, { "add-type", Verb.AddType },
            { "rt", Verb.RemoveType }, { "remove-type", Verb.RemoveType },
            { "s", Verb.Sort }, { "sort", Verb.Sort },
            { "rv", Verb.Reverse }, { "reverse", Verb.Reverse },
        };

        private sealed class Token
        {
            public bool Neg;
            public string Body;       // without the leading '!'
            public bool Quoted;       // starts with '"': never a command
            public bool IsVerb;       // a recognised /verb, /verb/path or //path
            public Verb Verb;
            public bool HasVia;       // spelled with a second '/'
            public string Via;
        }

        /// <summary>Splits on whitespace; a "..." run stays one token (a quote may start mid-token: title:"a b"). An unterminated quote runs to the end.</summary>
        public static List<string> Split(string text)
        {
            var raw = new List<string>();
            if (string.IsNullOrEmpty(text)) return raw;
            var cur = new StringBuilder();
            bool inQuote = false;
            foreach (char c in text)
            {
                if (c == '"') { inQuote = !inQuote; cur.Append(c); }
                else if (!inQuote && char.IsWhiteSpace(c))
                {
                    if (cur.Length > 0) { raw.Add(cur.ToString()); cur.Clear(); }
                }
                else cur.Append(c);
            }
            if (cur.Length > 0) raw.Add(cur.ToString());
            return raw;
        }

        public static string StripQuotes(string s)
        {
            if (s == null) return null;
            var sb = new StringBuilder();
            foreach (char c in s) if (c != '"') sb.Append(c);
            return sb.ToString();
        }

        private static List<Token> Tokenize(string text)
        {
            var tokens = new List<Token>();
            foreach (string raw in Split(text))
            {
                if (raw == "!") continue;                       // still being typed: inert
                var t = new Token();
                t.Neg = raw.Length > 1 && raw[0] == '!';
                t.Body = t.Neg ? raw.Substring(1) : raw;
                t.Quoted = t.Body.Length > 0 && t.Body[0] == '"';
                if (!t.Quoted && t.Body.StartsWith("//", StringComparison.Ordinal))
                {
                    t.IsVerb = true; t.Verb = Verb.FilterValue; t.HasVia = true; t.Via = t.Body.Substring(2);
                }
                else if (!t.Quoted && t.Body.Length > 0 && t.Body[0] == '/')
                {
                    string rest = t.Body.Substring(1);
                    int slash = rest.IndexOf('/');
                    string name = slash < 0 ? rest : rest.Substring(0, slash);
                    Verb v;
                    if (VerbNames.TryGetValue(name, out v) && !(v == Verb.Reverse && slash >= 0))
                    {
                        t.IsVerb = true; t.Verb = v;
                        if (slash >= 0) { t.HasVia = true; t.Via = rest.Substring(slash + 1); }
                    }
                }
                tokens.Add(t);
            }
            return tokens;
        }

        /// <summary>True when `s` is a non-empty direction word: it must be a substring of "ascending" or "descending".
        /// A prefix wins over a mere substring so "d" reads as descending and "a" as ascending.</summary>
        public static bool TryParseDirection(string s, out bool descending)
        {
            descending = false;
            if (string.IsNullOrEmpty(s)) return false;
            s = s.ToLowerInvariant();
            if ("ascending".StartsWith(s, StringComparison.Ordinal)) return true;
            if ("descending".StartsWith(s, StringComparison.Ordinal)) { descending = true; return true; }
            if ("ascending".Contains(s)) return true;
            if ("descending".Contains(s)) { descending = true; return true; }
            return false;
        }

        public static QueryPlan Parse(string text)
        {
            var plan = new QueryPlan();
            var tokens = Tokenize(text);
            int i = 0;
            while (i < tokens.Count)
            {
                Token t = tokens[i++];
                if (!t.IsVerb)
                {
                    if (t.Quoted || t.Body.Length == 0 || t.Body[0] != '/')
                        AddFilterValue(plan, t, false);
                    // an unrecognised /xyz is inert (still being typed)
                    continue;
                }

                switch (t.Verb)
                {
                    case Verb.FilterValue:
                        if (t.HasVia)
                        {
                            var term = new RowTerm { Form = TermForm.Via, Negated = t.Neg, Path = StripQuotes(t.Via) };
                            if (i < tokens.Count && !tokens[i].IsVerb)
                            {
                                Token v = tokens[i++];
                                term.HasValue = true;
                                term.Value = StripQuotes(v.Body);
                                if (v.Neg) term.Negated = true;
                            }
                            plan.RowTerms.Add(term);
                        }
                        else if (i < tokens.Count && !tokens[i].IsVerb)
                        {
                            AddFilterValue(plan, tokens[i++], t.Neg);
                        }
                        break;

                    case Verb.FilterType:
                    case Verb.AddType:
                    case Verb.RemoveType:
                        {
                            string path = null; bool via = false;
                            if (t.HasVia) { path = StripQuotes(t.Via); via = true; }
                            else if (i < tokens.Count && !tokens[i].IsVerb) path = StripQuotes(tokens[i++].Body);
                            if (path != null && !t.Neg)          // negation is meaningless on column verbs: parsed (args consumed) but inert
                                plan.ColumnOps.Add(new ColumnOp { Verb = t.Verb, Path = path, Via = via });
                            break;
                        }

                    case Verb.Sort:
                        {
                            var spec = new SortSpec();
                            if (t.HasVia)
                            {
                                foreach (string p in StripQuotes(t.Via).Split('/'))
                                    if (p.Length > 0) spec.Paths.Add(p);
                            }
                            else if (i < tokens.Count && !tokens[i].IsVerb)
                            {
                                spec.Paths.Add(StripQuotes(tokens[i++].Body));
                            }
                            if (spec.Paths.Count > 0 && i < tokens.Count && !tokens[i].IsVerb && !tokens[i].Quoted)
                            {
                                bool desc;
                                if (TryParseDirection(tokens[i].Body, out desc)) { spec.Descending = desc; i++; }
                            }
                            if (spec.Paths.Count > 0 && !t.Neg) plan.Sort = spec;     // last wins
                            break;
                        }

                    case Verb.Reverse:
                        if (!t.Neg) plan.Reverse = true;
                        break;
                }
            }
            return plan;
        }

        /// <summary>A bare word, a quoted phrase or the argument of `/fv`.</summary>
        private static void AddFilterValue(QueryPlan plan, Token t, bool extraNeg)
        {
            bool neg = t.Neg || extraNeg;
            string body = t.Body;
            if (t.Quoted)
            {
                plan.RowTerms.Add(new RowTerm { Form = TermForm.Text, Negated = neg, Text = StripQuotes(body) });
                return;
            }
            int colon = body.IndexOf(':');
            if (colon > 0)
            {
                string value = StripQuotes(body.Substring(colon + 1));
                if (value.StartsWith("!", StringComparison.Ordinal) && value.Length > 1) { neg = true; value = value.Substring(1); }
                plan.RowTerms.Add(new RowTerm
                {
                    Form = TermForm.Scoped, Negated = neg,
                    Path = StripQuotes(body.Substring(0, colon)), Value = value, HasValue = value.Length > 0
                });
                return;
            }
            plan.RowTerms.Add(new RowTerm { Form = TermForm.PathOnly, Negated = neg, Path = StripQuotes(body), Text = StripQuotes(body) });
        }
    }
}
