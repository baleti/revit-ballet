using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Python.Runtime;
using RevitBallet.Agents;

namespace RevitBallet.Engines.CPython3
{
    /// <summary>Exposed to scripts as <c>cancel</c> (like threading.Event.is_set).</summary>
    public sealed class CancelFlag
    {
        private readonly CancellationToken token;
        public CancelFlag(CancellationToken t) { token = t; }
        public bool is_set() { return token.IsCancellationRequested; }
    }

    /// <summary>Replaces sys.stdout / sys.stderr for the duration of one job.</summary>
    public sealed class PyWriter
    {
        private readonly TextWriter writer;
        public PyWriter(TextWriter w) { writer = w; }
        public int write(string text) { writer.Write(text); return text == null ? 0 : text.Length; }
        public void flush() { }
        public bool isatty() { return false; }
        public void Write(string text) { write(text); }
        public void WriteLine(string text) { writer.WriteLine(text); }
    }

    public sealed class CPython3Engine : IScriptEngine
    {
        public string Name { get { return "cpython3"; } }
        public string Language { get { return "python"; } }
        public string Version { get { return pyVersion ?? "3.12"; } }

        private static readonly object initLock = new object();
        private static bool initialized;
        private static string initError, pyVersion, pyDllPath;
        private ulong runningThread;

        private static string EngineDir { get { return Path.GetDirectoryName(typeof(CPython3Engine).Assembly.Location); } }

        /// <summary>Initialises the embedded interpreter once per process. Returns an error text, or null.</summary>
        private static string EnsureInit()
        {
            lock (initLock)
            {
                if (initialized) return null;
                if (initError != null) return initError;
                try
                {
                    var dir = EngineDir;
                    // not a wildcard: on Windows "python3??.dll" also matches the stub python3.dll, which lacks the symbols pythonnet needs
                    var dll = Directory.GetFiles(dir, "python3*.dll").FirstOrDefault(f => System.Text.RegularExpressions.Regex.IsMatch(Path.GetFileName(f), @"^python3\d+\.dll$"));
                    if (dll == null) return initError = "the embedded Python runtime (python3xx.dll) was not found in " + dir;
                    pyDllPath = dll;
                    Runtime.PythonDLL = dll;
                    PythonEngine.Initialize();
                    pyVersion = null;
                    using (Py.GIL())
                    {
                        dynamic sys = Py.Import("sys");
                        pyVersion = ((string)sys.version.ToString()).Split(' ')[0];
                    }
                    PythonEngine.BeginAllowThreads();   // release the GIL so any thread can take it with Py.GIL()
                    initialized = true;
                    return null;
                }
                catch (Exception ex)
                {
                    var chain = new System.Text.StringBuilder("failed to start CPython (" + (pyDllPath ?? "no dll") + "): ");
                    for (var e = ex; e != null; e = e.InnerException) chain.Append(e.GetType().Name).Append(": ").Append(e.Message).Append(" | ");
                    return initError = chain.ToString();
                }
            }
        }

        public EngineCompileResult Compile(string code, EngineSession session, CancellationToken cancel)
        {
            var result = new EngineCompileResult();
            var err = EnsureInit();
            if (err != null) throw new InvalidOperationException(err);
            using (Py.GIL())
            {
                try
                {
                    // compile() only to surface syntax errors with positions; execution re-parses (cheap, keeps the code object off the worker thread)
                    using (var builtins = Py.Import("builtins"))
                    using (var co = builtins.InvokeMethod("compile", new PyString(code), new PyString("<script>"), new PyString("exec")))
                    { }
                    result.Success = true; result.Compiled = code;
                }
                catch (PythonException ex)
                {
                    result.Success = false;
                    int line = 0, col = 0; string msg = ex.Message;
                    try
                    {
                        using (var v = ex.Value)
                        {
                            if (v.HasAttr("lineno")) { using (var l = v.GetAttr("lineno")) { if (!l.IsNone()) line = l.As<int>(); } }
                            if (v.HasAttr("offset")) { using (var o = v.GetAttr("offset")) { if (!o.IsNone()) col = o.As<int>(); } }
                            if (v.HasAttr("msg")) { using (var m = v.GetAttr("msg")) msg = ex.Type.Name + ": " + m.ToString(); }
                        }
                    }
                    catch { }
                    result.Diagnostics.Add(new EngineDiagnostic { Severity = "error", Id = ex.Type.Name, Line = line, Column = col, Message = msg });
                }
            }
            return result;
        }

        public EngineRunResult Run(object compiled, EngineContext ctx)
        {
            var rr = new EngineRunResult();
            var code = (string)compiled;
            using (Py.GIL())
            {
                PyDict globals = null; bool ownGlobals = false;
                PyObject sys = null, oldOut = null, oldErr = null;
                try
                {
                    if (ctx.Session != null) globals = ctx.Session.State as PyDict ?? (PyDict)(ctx.Session.State = new PyDict());
                    else { globals = new PyDict(); ownGlobals = true; }

                    Set(globals, "uiapp", ctx.UIApp); Set(globals, "UIApp", ctx.UIApp); Set(globals, "__revit__", ctx.UIApp);
                    Set(globals, "uidoc", ctx.UIDoc); Set(globals, "UIDoc", ctx.UIDoc);
                    Set(globals, "doc", ctx.Doc); Set(globals, "Doc", ctx.Doc);
                    Set(globals, "cancel", new CancelFlag(ctx.Cancel));
                    Set(globals, "out", new PyWriter(ctx.Out));
                    if (globals.HasKey("result")) globals.DelItem("result");

                    sys = Py.Import("sys");
                    oldOut = sys.GetAttr("stdout"); oldErr = sys.GetAttr("stderr");
                    var writer = new PyWriter(ctx.Out).ToPython();
                    sys.SetAttr("stdout", writer); sys.SetAttr("stderr", writer);

                    runningThread = PythonEngine.GetPythonThreadID();
                    if (ctx.Cancel.IsCancellationRequested) throw new OperationCanceledException("cancelled before it started");
                    PythonEngine.Exec(code, globals);

                    if (globals.HasKey("result"))
                        using (var r = globals.GetItem("result")) rr.Value = ToManaged(r, 0);
                    rr.Ok = true;
                }
                catch (OperationCanceledException ex) { rr.ErrorKind = "cancelled"; rr.ErrorMessage = ex.Message; }
                catch (PythonException ex)
                {
                    if (ex.Type != null && ex.Type.Name == "KeyboardInterrupt" && ctx.Cancel.IsCancellationRequested)
                    { rr.ErrorKind = "cancelled"; rr.ErrorMessage = "cancelled"; }
                    else
                    {
                        rr.ErrorKind = "runtime";
                        rr.ErrorMessage = (ex.Type != null ? ex.Type.Name + ": " : "") + ex.Message;
                        rr.ErrorStack = ex.StackTrace;
                    }
                }
                catch (Exception ex) { rr.ErrorKind = "runtime"; rr.ErrorMessage = ex.GetType().Name + ": " + ex.Message; rr.ErrorStack = ex.StackTrace; }
                finally
                {
                    runningThread = 0;
                    try { if (sys != null && oldOut != null) sys.SetAttr("stdout", oldOut); if (sys != null && oldErr != null) sys.SetAttr("stderr", oldErr); } catch { }
                    oldOut?.Dispose(); oldErr?.Dispose(); sys?.Dispose();
                    if (ownGlobals && globals != null) globals.Dispose();
                }
            }
            return rr;
        }

        private static void Set(PyDict d, string name, object value)
        {
            using (var py = value == null ? PyObject.None : value.ToPython()) d.SetItem(name, py);
        }

        /// <summary>Plain Python values become plain .NET values (so the host can emit JSON); everything else is reported through str().</summary>
        private static object ToManaged(PyObject o, int depth)
        {
            if (o == null || o.IsNone()) return null;
            var tn = o.GetPythonType().Name;
            if (tn == "bool") return o.As<bool>();
            if (PyInt.IsIntType(o)) { try { return o.As<long>(); } catch { return o.ToString(); } }
            if (PyFloat.IsFloatType(o)) return o.As<double>();
            if (PyString.IsStringType(o)) return o.As<string>();
            if (depth < 8 && (PyList.IsListType(o) || PyTuple.IsTupleType(o)))
            {
                var list = new List<object>();
                using (var seq = new PySequence(o))
                {
                    long n = seq.Length();
                    for (long i = 0; i < n; i++) using (var item = seq[(int)i]) list.Add(ToManaged(item, depth + 1));
                }
                return list;
            }
            if (depth < 8 && PyDict.IsDictType(o))
            {
                var dict = new Dictionary<string, object>();
                using (var d = new PyDict(o))
                    foreach (PyObject k in d.Keys()) using (k) using (var v = d[k]) dict[k.ToString()] = ToManaged(v, depth + 1);
                return dict;
            }
            // a .NET object wrapped by pythonnet (a Revit element, XYZ, ...): hand over the real object so the host can summarise it
            try { var m = o.AsManagedObject(typeof(object)); if (m != null && !(m is PyObject)) return m; } catch { }
            return o.ToString();
        }

        public void ResetSession(EngineSession session)
        {
            if (session == null || !(session.State is PyDict)) { if (session != null) session.State = null; return; }
            var d = (PyDict)session.State; session.State = null;
            using (Py.GIL()) d.Dispose();
        }

        public void Abort(EngineContext context)
        {
            var id = runningThread;
            if (id == 0) return;
            // KeyboardInterrupt in the running thread; needs the GIL, so never block the caller on it
            Task.Run(() => { try { using (Py.GIL()) PythonEngine.Interrupt(id); } catch { } });
        }
    }
}
