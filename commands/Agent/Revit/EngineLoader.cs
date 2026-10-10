using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using RevitBallet.Agents;

namespace RevitBallet.Commands.Agent
{
    /// <summary>Adapter: engines discovered from engines/ next to the add-in, loaded on first use.</summary>
    internal sealed class EngineLoaderSource : IEngineSource
    {
        private readonly Dictionary<string, EngineEntry> engines = EngineLoader.Discover();
        public IDictionary<string, EngineEntry> Engines { get { return engines; } }
        public IScriptEngine Load(EngineEntry entry) { return EngineLoader.Load(entry); }
        public void Reset(EngineEntry entry) { lock (entry.LoadLock) entry.Instance = null; }
    }

    internal static class EngineLoader
    {
        /// <summary>Set by the runtime: the engines folder next to the installed add-in (this assembly may have been loaded from bytes and then has no location).</summary>
        public static string EnginesRootOverride;
        public static string EnginesRoot =>
            EnginesRootOverride ?? Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "engines");

        private static readonly List<string> engineDirs = new List<string>();
        private static bool resolverInstalled;

        /// <summary>Reads manifests only (no engine code is loaded). Manifests that are invalid or point outside their directory are skipped.</summary>
        public static Dictionary<string, EngineEntry> Discover()
        {
            var result = new Dictionary<string, EngineEntry>(StringComparer.OrdinalIgnoreCase);
            var root = EnginesRoot;
            if (!System.IO.Directory.Exists(root)) return result;
            foreach (var dir in System.IO.Directory.GetDirectories(root))
            {
                var manifestPath = Path.Combine(dir, "engine.json");
                if (!File.Exists(manifestPath)) continue;
                try
                {
                    var m = JsonConvert.DeserializeObject<EngineManifest>(File.ReadAllText(manifestPath));
                    if (m == null || string.IsNullOrWhiteSpace(m.Name) || string.IsNullOrWhiteSpace(m.Assembly) || string.IsNullOrWhiteSpace(m.Type)) continue;
                    if (m.Protocol != 1) { RevitBalletServer.LogToRevit($"[AGENT] engine {m.Name}: unsupported protocol {m.Protocol}"); continue; }
                    var asmPath = Path.GetFullPath(Path.Combine(dir, m.Assembly));
                    if (!asmPath.StartsWith(Path.GetFullPath(dir) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                    { RevitBalletServer.LogToRevit($"[AGENT] engine {m.Name}: assembly path escapes its directory, ignored"); continue; }
                    result[m.Name] = new EngineEntry { Manifest = m, Directory = dir };
                }
                catch (Exception ex) { RevitBalletServer.LogToRevit($"[AGENT] bad engine manifest {manifestPath}: {ex.Message}"); }
            }
            return result;
        }

        /// <summary>Loads the engine assembly and creates the engine. Thread-safe; failures are remembered in <see cref="EngineEntry.LoadError"/>.</summary>
        public static IScriptEngine Load(EngineEntry e)
        {
            lock (e.LoadLock)
            {
                if (e.Instance != null) return e.Instance;
                if (e.LoadError != null) return null;
                try
                {
                    var path = Path.GetFullPath(Path.Combine(e.Directory, e.Manifest.Assembly));
                    RegisterDirectory(e.Directory);
                    Assembly asm;
#if NET8_0_OR_GREATER
                    // Revit 2025+ loads add-ins into their own AssemblyLoadContext; load the engine into the same one so the
                    // shared contract assembly is the same type for host and engine.
                    var alc = System.Runtime.Loader.AssemblyLoadContext.GetLoadContext(typeof(EngineLoader).Assembly);
                    asm = alc.LoadFromAssemblyPath(path);
#else
                    asm = Assembly.LoadFrom(path);
#endif
                    var type = asm.GetType(e.Manifest.Type, throwOnError: true);
                    e.Instance = (IScriptEngine)Activator.CreateInstance(type);
                    RevitBalletServer.LogToRevit($"[AGENT] engine loaded: {e.Manifest.Name} {e.Instance.Version}");
                    return e.Instance;
                }
                catch (Exception ex)
                {
                    var inner = ex is TargetInvocationException tie && tie.InnerException != null ? tie.InnerException : ex;
                    e.LoadError = inner.GetType().Name + ": " + inner.Message;
                    RevitBalletServer.LogToRevit($"[AGENT] engine {e.Manifest.Name} failed to load: {inner}");
                    return null;
                }
            }
        }

        /// <summary>Lets an engine's own dependencies (Roslyn, IronPython, ...) be found next to it.</summary>
        private static void RegisterDirectory(string dir)
        {
            lock (engineDirs)
            {
                if (!engineDirs.Contains(dir, StringComparer.OrdinalIgnoreCase)) engineDirs.Add(dir);
                if (resolverInstalled) return;
                resolverInstalled = true;
#if NET8_0_OR_GREATER
                var alc = System.Runtime.Loader.AssemblyLoadContext.GetLoadContext(typeof(EngineLoader).Assembly);
                alc.Resolving += (ctx, name) => FindIn(name, p => ctx.LoadFromAssemblyPath(p));
#else
                AppDomain.CurrentDomain.AssemblyResolve += (s, args) => FindIn(new AssemblyName(args.Name), Assembly.LoadFrom);
#endif
            }
        }

        private static Assembly FindIn(AssemblyName name, Func<string, Assembly> load)
        {
            string[] dirs; lock (engineDirs) dirs = engineDirs.ToArray();
            foreach (var d in dirs)
            {
                var p = Path.Combine(d, name.Name + ".dll");
                if (File.Exists(p)) { try { return load(p); } catch { } }
            }
            return null;
        }
    }
}
