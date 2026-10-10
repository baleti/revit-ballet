using System;
using System.Collections;
using System.Collections.Generic;
using Autodesk.Revit.DB;
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

            var revit = RevitToken(v, depth, ref budget);
            if (revit != null) return revit;
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
            if (v is IEnumerable seq && !(v is Element) && !(v is Document))
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

        /// <summary>Compact summaries of common Revit API types; null when v is not one of them.</summary>
        private static JToken RevitToken(object v, int depth, ref int budget)
        {
            if (v is ElementId eid) return new JObject { ["id"] = IdValue(eid) };
            if (v is Element el)
            {
                var o = new JObject { ["id"] = IdValue(el.Id), ["class"] = el.GetType().Name };
                try { o["uniqueId"] = el.UniqueId; } catch { }
                try { var c = el.Category; if (c != null) o["category"] = c.Name; } catch { }
                try { o["name"] = el.Name; } catch { }
                return o;
            }
            if (v is Category cat) return new JObject { ["id"] = IdValue(cat.Id), ["name"] = cat.Name };
            if (v is XYZ p) return new JObject { ["x"] = p.X, ["y"] = p.Y, ["z"] = p.Z };
            if (v is UV uv) return new JObject { ["u"] = uv.U, ["v"] = uv.V };
            if (v is BoundingBoxXYZ bb)
            {
                return new JObject { ["min"] = bb.Min == null ? null : RevitToken(bb.Min, depth, ref budget), ["max"] = bb.Max == null ? null : RevitToken(bb.Max, depth, ref budget) };
            }
            if (v is Document d)
            {
                var o = new JObject { ["title"] = d.Title };
                try { o["path"] = d.PathName; } catch { }
                return o;
            }
            if (v is Parameter prm)
            {
                var o = new JObject { ["name"] = prm.Definition?.Name, ["storage"] = prm.StorageType.ToString(), ["hasValue"] = prm.HasValue };
                try { o["valueString"] = prm.AsValueString(); } catch { }
                try
                {
                    switch (prm.StorageType)
                    {
                        case StorageType.Double: o["value"] = prm.AsDouble(); break;
                        case StorageType.Integer: o["value"] = prm.AsInteger(); break;
                        case StorageType.String: o["value"] = prm.AsString(); break;
                        case StorageType.ElementId: o["value"] = IdValue(prm.AsElementId()); break;
                    }
                }
                catch { }
                return o;
            }
            return null;
        }

        /// <summary>ElementId.Value (long) exists from Revit 2024, IntegerValue (int) before; read whichever exists.</summary>
        private static long IdValue(ElementId id)
        {
            if (id == null) return -1;
            var t = typeof(ElementId);
            var prop = t.GetProperty("Value") ?? t.GetProperty("IntegerValue");
            return prop == null ? -1 : System.Convert.ToInt64(prop.GetValue(id, null));
        }
    }
}
