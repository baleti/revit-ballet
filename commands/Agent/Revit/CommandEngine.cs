using System;
using System.Diagnostics;
using System.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Agents;

namespace RevitBallet.Commands.Agent
{
    /// <summary>
    /// Built-in job engine "command": the job's code is the class name of an add-in command, which is run as if the user had invoked it
    /// (POST /v1/commands/{name}/invoke makes such a job). Queueing, time limits, dialog reporting and cancellation are the job machinery's.
    /// </summary>
    internal sealed class CommandEngine : IScriptEngine
    {
        public string Name { get { return "command"; } }
        public string Language { get { return "add-in command"; } }
        public string Version { get { return "1.0"; } }

        public EngineCompileResult Compile(string code, EngineSession session, CancellationToken cancel)
        {
            var r = new EngineCompileResult();
            var name = (code ?? "").Trim();
            var type = CommandCatalog.Find(name);
            if (type == null)
            {
                r.Diagnostics.Add(new EngineDiagnostic { Severity = "error", Id = "unknown_command", Message = "no add-in command named '" + name + "' (GET /v1/commands lists them)" });
                return r;
            }
            r.Success = true; r.Compiled = type;
            return r;
        }

        public EngineRunResult Run(object compiled, EngineContext context)
        {
            var result = new EngineRunResult();
            var type = (Type)compiled;
            try
            {
                var command = (IExternalCommand)Activator.CreateInstance(type);
                // the constructor is not public (Revit builds these itself when the user starts a command), the properties are writable
                var data = (ExternalCommandData)Activator.CreateInstance(typeof(ExternalCommandData), true);
                data.Application = (UIApplication)context.UIApp;
                string message = "";
                var elements = new ElementSet();
                var sw = Stopwatch.StartNew();
                Result r = command.Execute(data, ref message, elements);
                CommandUsageTracker.Record(type.Name, r, sw.ElapsedMilliseconds);
                if (!string.IsNullOrEmpty(message)) context.Out.WriteLine(message);
                result.Value = r.ToString();
                if (r == Result.Succeeded) result.Ok = true;
                else if (r == Result.Cancelled) { result.ErrorKind = "cancelled"; result.ErrorMessage = "the command was cancelled"; }
                else { result.ErrorKind = "runtime"; result.ErrorMessage = string.IsNullOrEmpty(message) ? "the command failed" : message; }
            }
            catch (Exception ex)
            {
                var inner = ex is System.Reflection.TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                result.ErrorKind = "runtime"; result.ErrorMessage = inner.GetType().Name + ": " + inner.Message; result.ErrorStack = inner.StackTrace;
            }
            return result;
        }

        public void ResetSession(EngineSession session) { }
        public void Abort(EngineContext context) { }
    }
}
