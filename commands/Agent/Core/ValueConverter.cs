using System;
using System.Collections;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace RevitBallet.Commands.Agent
{
    /// <summary>
    /// Script result -> JSON. Must run on the Revit UI thread (it reads a few properties of Revit objects).
    /// Deliberately conservative: primitives, strings, enums, dates, a fixed set of Revit types (compact summaries, never a reflected object graph)
    /// and nested dictionaries/sequences of those become JSON. Anything else is reported through ToString() only.
    /// </summary>
    internal static class ValueConverter
    {
        private const int MaxItems = 200000, MaxDepth = 8;

        /// <summary>Hooks tried before the generic handling: return a compact JSON summary of a type you know, or null (Revit types are registered by the add-in).</summary>
        public static readonly List<Func<object, JToken>> Extensions = new List<Func<object, JToken>>();

        public static void Convert(object value, out string type, out JToken json, out string text)
        {
            bool overflow;
            Convert(value, out type, out json, out text, out overflow);
        }

        /// <param name="overflow">true when the value had more than <see cref="MaxItems"/> items: json is then null (never a silently shortened list).</param>
        public static void Convert(object value, out string type, out JToken json, out string text, out bool overflow)
        {
            overflow = false;
            if (value == null) { type = null; json = null; text = null; return; }
            type = value.GetType().FullName;
            try { text = value.ToString(); } catch (Exception ex) { text = "<ToString failed: " + ex.Message + ">"; }
            int budget = MaxItems;
            try { json = ToToken(value, 0, ref budget); } catch { json = null; }
            if (budget < 0) { json = null; overflow = true; }
        }

        private static JToken ToToken(object v, int depth, ref int budget)
        {
            if (v == null) return JValue.CreateNull();
            if (v is bool b) return new JValue(b);
            if (v is string || v is char) return new JValue(v.ToString());
            if (v is int || v is long || v is short || v is byte || v is sbyte || v is uint || v is ulong || v is ushort || v is float || v is double || v is decimal)
                return new JValue(v);
            if (v is Enum) return new JValue(v.ToString());
            if (v is DateTime dt) return new JValue(dt.ToString("o"));
            if (v is DateTimeOffset dto) return new JValue(dto.ToString("o"));
            if (v is Guid g) return new JValue(g.ToString());

            foreach (var ext in Extensions)
            {
                JToken t = null;
                try { t = ext(v); } catch { }
                if (t != null) return t;
            }
            if (depth >= MaxDepth) return null;

            if (v is IDictionary dict)
            {
                var o = new JObject();
                foreach (DictionaryEntry e in dict)
                {
                    if (--budget < 0) return null;
                    var t = ToToken(e.Value, depth + 1, ref budget);
                    if (t == null) return null;
                    o[System.Convert.ToString(e.Key)] = t;
                }
                return o;
            }
            if (v is IEnumerable seq)
            {
                var a = new JArray();
                foreach (var item in seq)
                {
                    if (--budget < 0) return null;
                    var t = ToToken(item, depth + 1, ref budget);
                    if (t == null) return null;   // not representable -> whole value is text-only
                    a.Add(t);
                }
                return a;
            }
            return null;
        }
    }
}
