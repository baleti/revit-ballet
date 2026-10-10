using System;
using System.Collections;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace RevitBallet.Commands.Agent
{
    /// <summary>
    /// Script result -> JSON. Deliberately conservative: only primitives, strings, enums, dates and (nested) dictionaries/sequences of those become
    /// JSON. Anything else (Revit elements, ids, ...) is reported through ToString() only, because walking arbitrary object graphs
    /// would evaluate every property getter of Revit API objects.
    /// </summary>
    internal static class ValueConverter
    {
        private const int MaxItems = 10000, MaxDepth = 8;

        public static void Convert(object value, out string type, out JToken json, out string text)
        {
            if (value == null) { type = null; json = null; text = null; return; }
            type = value.GetType().FullName;
            try { text = value.ToString(); } catch (Exception ex) { text = "<ToString failed: " + ex.Message + ">"; }
            int budget = MaxItems;
            try { json = ToToken(value, 0, ref budget); } catch { json = null; }
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
            if (depth >= MaxDepth) return null;

            if (v is IDictionary dict)
            {
                var o = new JObject();
                foreach (DictionaryEntry e in dict)
                {
                    if (--budget < 0) break;
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
                    if (--budget < 0) break;
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
