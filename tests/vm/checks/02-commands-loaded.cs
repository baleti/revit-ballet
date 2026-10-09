var asm = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.GetName().Name == "revit-ballet");
if (asm == null) return "NOT LOADED";
Type[] types;
try { types = asm.GetTypes(); }
catch (System.Reflection.ReflectionTypeLoadException e) { types = e.Types.Where(t => t != null).ToArray(); }
int n = types.Count(t => typeof(IExternalCommand).IsAssignableFrom(t) && !t.IsAbstract);
return n >= 20 ? "OK" : "TOO FEW COMMANDS: " + n;
