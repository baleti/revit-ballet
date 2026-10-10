using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.UI;

namespace RevitBallet.Commands.Agent
{
    /// <summary>
    /// The add-in commands that can be invoked remotely: every concrete IExternalCommand class of this assembly, with the input and output
    /// described by [CommandMeta] / [CommandOutput]. The same discovery InvokeAddinCommand does for its list, minus the UI.
    /// </summary>
    internal static class CommandCatalog
    {
        private static readonly object sync = new object();
        private static List<CommandInfo> cached;
        private static Dictionary<string, Type> types;

        public static IList<CommandInfo> List()
        {
            Ensure();
            return cached;
        }

        public static Type Find(string name)
        {
            Ensure();
            Type t;
            return name != null && types.TryGetValue(name, out t) ? t : null;
        }

        private static void Ensure()
        {
            lock (sync)
            {
                if (cached != null) return;
                var list = new List<CommandInfo>();
                var map = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase);
                Type[] all;
                try { all = typeof(CommandCatalog).Assembly.GetTypes(); }
                catch (ReflectionTypeLoadException ex) { all = ex.Types.Where(t => t != null).ToArray(); }
                foreach (var t in all.Where(t => typeof(IExternalCommand).IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract).OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase))
                {
                    if (map.ContainsKey(t.Name)) continue;                         // names are unique across the add-in; keep the first if not
                    var meta = (CommandMetaAttribute)t.GetCustomAttributes(typeof(CommandMetaAttribute), false).FirstOrDefault();
                    var output = (CommandOutputAttribute)t.GetCustomAttributes(typeof(CommandOutputAttribute), false).FirstOrDefault();
                    map[t.Name] = t;
                    list.Add(new CommandInfo
                    {
                        Name = t.Name, FullName = t.FullName, DisplayName = CommandRegistry.GetDisplayName(t.Name),
                        Input = meta == null ? "" : meta.Input, Output = output == null ? "" : output.Output,
                    });
                }
                types = map; cached = list;
            }
        }
    }
}
