using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using Newtonsoft.Json.Linq;
using RevitBallet.Agents;

namespace RevitBallet.Commands.Agent
{
    internal enum DialogMode { Report, Fail, Respond }

    /// <summary>What a job does when a modal dialog blocks Revit while it waits or runs.</summary>
    internal sealed class DialogPolicy
    {
        public DialogMode Mode = DialogMode.Report;
        public string Button;   // for Respond
        public static readonly DialogPolicy Default = new DialogPolicy();
    }

    internal enum JobState { Queued, Compiling, Running, Done, Failed, Cancelled, Timeout }

    /// <summary>One submitted script. State changes happen under <see cref="Sync"/>.</summary>
    internal sealed class Job
    {
        public string Id;
        public string Engine;
        public string SessionName;
        public string Code;
        public string Client;       // X-Agent-Name header, for the audit log
        public string RequestId;    // optional idempotency key from the client
        public int TimeoutMs = 30000;
        public DialogPolicy DialogPolicy = DialogPolicy.Default;
        public DateTime LastDialogPoll;
        public DateTime? DialogSeenSince;
        public DialogInfo ErrorDialog;                                 // set when the job ended because of a dialog
        public DialogInfo BlockedBy;                                   // while a dialog has been holding this job up longer than the grace period
        public readonly List<JObject> Dialogs = new List<JObject>();   // every distinct dialog seen while the job waited or ran, with what was done about it
        public readonly HashSet<string> DialogKeys = new HashSet<string>();

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
        public bool ValueTruncated, OutputDropped;
        public int EventChars; public bool EventsCapped;

        public const int MaxInline = 256 * 1024;        // output / value text returned inside the job object
        public const int MaxStored = 4 * 1024 * 1024;   // output kept in memory (served by GET /v1/jobs/{id}/output)
        public const int MaxEventChars = 1024 * 1024;   // streamed over SSE

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

        /// <summary>The text of the output between offset and offset+limit (for GET /v1/jobs/{id}/output).</summary>
        public string OutputSlice(int offset, int limit, out int total)
        {
            lock (Output)
            {
                total = Output.Length;
                if (offset < 0) offset = 0;
                if (offset >= total) return "";
                return Output.ToString(offset, Math.Min(Math.Max(limit, 0), total - offset));
            }
        }

        /// <param name="full">true for the audit log: whole stored output and the code.</param>
        public JObject ToJson(bool full = false)
        {
            lock (Sync)
            {
                string output; int outputLength;
                lock (Output) { outputLength = Output.Length; output = full || outputLength <= MaxInline ? Output.ToString() : Output.ToString(0, MaxInline); }
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
                    ["client"] = Client,
                    ["output"] = output,
                    ["outputTruncated"] = outputLength > output.Length || OutputDropped,
                    ["outputLength"] = outputLength,
                    ["value"] = ValueType == null ? null : new JObject { ["type"] = ValueType, ["json"] = ValueJson, ["text"] = ValueText, ["truncated"] = ValueTruncated },
                    ["error"] = ErrorKind == null ? null : new JObject { ["kind"] = ErrorKind, ["message"] = ErrorMessage, ["stack"] = ErrorStack, ["dialog"] = ErrorDialog == null ? null : ErrorDialog.ToJson() },
                    ["blockedBy"] = BlockedBy == null ? null : BlockedBy.ToJson(DialogSeenSince.HasValue ? (DateTime.UtcNow - DialogSeenSince.Value).TotalMilliseconds : (double?)null),
                    ["dialogs"] = new JArray(Dialogs.ToArray()),
                    ["timings"] = new JObject { ["queueMs"] = Math.Round(QueueMs), ["compileMs"] = Math.Round(CompileMs), ["runMs"] = Math.Round(RunMs) },
                };
                var diags = new JArray();
                foreach (var d in Diagnostics)
                    diags.Add(new JObject { ["severity"] = d.Severity, ["id"] = d.Id, ["line"] = d.Line, ["column"] = d.Column, ["message"] = d.Message });
                j["diagnostics"] = diags;
                if (full) { j["code"] = Code; j["requestId"] = RequestId; }
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
            lock (job.Output)
            {
                int room = Job.MaxStored - job.Output.Length;
                if (room <= 0) { job.OutputDropped = true; return; }
                if (text.Length > room) { text = text.Substring(0, room); job.OutputDropped = true; }
                job.Output.Append(text);
            }
            // SSE: stream only the first MaxEventChars; a client that wants more reads /output
            if (job.EventChars < Job.MaxEventChars)
            {
                job.EventChars += text.Length;
                job.AddEvent("output", new JObject { ["text"] = text });
            }
            else if (!job.EventsCapped)
            {
                job.EventsCapped = true;
                job.AddEvent("output", new JObject { ["text"] = "\n[output stream truncated here; the rest is available at GET /v1/jobs/" + job.Id + "/output]\n" });
            }
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
