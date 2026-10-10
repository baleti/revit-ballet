using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using Newtonsoft.Json.Linq;
using RevitBallet.Agents;

namespace RevitBallet.Commands.Agent
{
    internal enum JobState { Queued, Compiling, Running, Done, Failed, Cancelled, Timeout }

    /// <summary>One submitted script. State changes happen under <see cref="Sync"/>.</summary>
    internal sealed class Job
    {
        public string Id;
        public string Engine;
        public string SessionName;
        public string Code;
        public int TimeoutMs = 30000;

        public readonly object Sync = new object();
        public JobState State = JobState.Queued;
        public DateTime Submitted = DateTime.UtcNow;
        public DateTime? Started, Finished;
        public double QueueMs, CompileMs, RunMs;

        public readonly CancellationTokenSource Cts = new CancellationTokenSource();
        public readonly ManualResetEventSlim Done = new ManualResetEventSlim(false);      // job is final
        public readonly ManualResetEventSlim UiStarted = new ManualResetEventSlim(false); // UI thread picked it up
        public readonly ManualResetEventSlim UiDone = new ManualResetEventSlim(false);    // UI thread returned from Run
        public bool Abandoned;        // UI thread must skip it (timed out or cancelled before it started)
        public bool CancelRequested;

        public JobWriter Writer;
        public EngineSession SessionObject;
        public IScriptEngine EngineInstance;
        public EngineContext Context;
        public object Compiled;
        public EngineRunResult RunResult;

        public readonly StringBuilder Output = new StringBuilder();
        public readonly List<EngineDiagnostic> Diagnostics = new List<EngineDiagnostic>();
        public string ErrorKind, ErrorMessage, ErrorStack;
        public JToken ValueJson; public string ValueType, ValueText;

        // server-sent events: each entry is "event: name\ndata: json\n\n"
        public readonly List<string> Events = new List<string>();

        public bool IsFinal => State == JobState.Done || State == JobState.Failed || State == JobState.Cancelled || State == JobState.Timeout;

        public void AddEvent(string name, JToken data)
        {
            lock (Events)
            {
                Events.Add("event: " + name + "\ndata: " + data.ToString(Newtonsoft.Json.Formatting.None) + "\n\n");
                Monitor.PulseAll(Events);
            }
        }

        public void SetState(JobState s)
        {
            lock (Sync) { State = s; }
            AddEvent("state", new JObject { ["state"] = StateName(s) });
        }

        public static string StateName(JobState s) => s.ToString().ToLowerInvariant();

        public JObject ToJson()
        {
            lock (Sync)
            {
                string output; lock (Output) output = Output.ToString();
                var j = new JObject
                {
                    ["id"] = Id,
                    ["engine"] = Engine,
                    ["session"] = SessionName,
                    ["state"] = StateName(State),
                    ["ok"] = State == JobState.Done,
                    ["submitted"] = Submitted.ToString("o"),
                    ["started"] = Started?.ToString("o"),
                    ["finished"] = Finished?.ToString("o"),
                    ["output"] = output,
                    ["value"] = ValueType == null ? null : new JObject { ["type"] = ValueType, ["json"] = ValueJson, ["text"] = ValueText },
                    ["error"] = ErrorKind == null ? null : new JObject { ["kind"] = ErrorKind, ["message"] = ErrorMessage, ["stack"] = ErrorStack },
                    ["timings"] = new JObject { ["queueMs"] = Math.Round(QueueMs), ["compileMs"] = Math.Round(CompileMs), ["runMs"] = Math.Round(RunMs) },
                };
                var diags = new JArray();
                foreach (var d in Diagnostics)
                    diags.Add(new JObject { ["severity"] = d.Severity, ["id"] = d.Id, ["line"] = d.Line, ["column"] = d.Column, ["message"] = d.Message });
                j["diagnostics"] = diags;
                return j;
            }
        }
    }

    /// <summary>Per-job output: kept in the job and streamed as "output" events.</summary>
    internal sealed class JobWriter : TextWriter
    {
        private readonly Job job;
        public JobWriter(Job job) { this.job = job; }
        public override Encoding Encoding => Encoding.UTF8;

        public override void Write(char value) => Emit(value.ToString());
        public override void Write(string value) { if (!string.IsNullOrEmpty(value)) Emit(value); }
        public override void Write(char[] buffer, int index, int count) => Emit(new string(buffer, index, count));

        private void Emit(string text)
        {
            lock (job.Output) job.Output.Append(text);
            job.AddEvent("output", new JObject { ["text"] = text });
        }
    }

    /// <summary>
    /// Console.Out stays global in the process; this router forwards to the current job's writer only on the thread that is running
    /// that job, so output of other code never lands in a job's result.
    /// </summary>
    internal sealed class ConsoleRouter : TextWriter
    {
        [ThreadStatic] private static TextWriter current;
        private readonly TextWriter original;
        private ConsoleRouter(TextWriter original) { this.original = original; }
        public override Encoding Encoding => original.Encoding;

        private static ConsoleRouter installed;
        public static void Install()
        {
            if (installed != null) return;
            installed = new ConsoleRouter(Console.Out);
            Console.SetOut(installed);
        }
        public static IDisposable Scope(TextWriter w)
        {
            var prev = current; current = w;
            return new Restore(prev);
        }
        private sealed class Restore : IDisposable
        {
            private readonly TextWriter prev; public Restore(TextWriter p) { prev = p; }
            public void Dispose() { current = prev; }
        }
        private TextWriter Target => current ?? original;
        public override void Write(char value) => Target.Write(value);
        public override void Write(string value) => Target.Write(value);
        public override void Write(char[] buffer, int index, int count) => Target.Write(buffer, index, count);
        public override void Flush() => Target.Flush();
    }
}
