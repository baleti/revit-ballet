using System;
using Autodesk.Revit.DB;
using Newtonsoft.Json.Linq;

namespace RevitBallet.Commands.Agent
{
    /// <summary>
    /// Compact JSON summaries of common Revit API types, registered with <see cref="ValueConverter.Extensions"/>. Runs on the Revit UI thread
    /// (it reads a few properties). Never a reflected object graph: walking arbitrary Revit objects would evaluate every property getter.
    /// </summary>
    internal static class RevitValues
    {
        /// <summary>Compact summaries of common Revit API types; null when v is not one of them.</summary>
        public static JToken TryToken(object v)
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
                return new JObject { ["min"] = bb.Min == null ? null : TryToken(bb.Min), ["max"] = bb.Max == null ? null : TryToken(bb.Max) };
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
