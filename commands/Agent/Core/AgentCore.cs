using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Newtonsoft.Json.Linq;
using RevitBallet.Agents;

namespace RevitBallet.Commands.Agent
{
    internal sealed class AgentException : Exception
    {
        public readonly int Status; public readonly string Kind;
        public AgentException(int status, string kind, string message) : base(message) { Status = status; Kind = kind; }
    }

    /// <summary>
    /// The agent core: job queue, worker thread, engines and sessions. Protocol: docs/agent-protocol.md. No Revit types in here: the UI thread, the
    /// engines and facts about the process come in through <see cref="IUiBridge"/>, <see cref="IEngineSource"/> and <see cref="IAgentEnvironment"/>,
    /// so the whole queue/timeout/cancel logic can be tested without Revit (tests/host).
    /// Pipeline per job: worker thread compiles, the UI thread (Revit external event) runs, the worker watches timeouts. One job at a time,
    /// strictly in submission order, so named sessions see their own state in order.
    /// </summary>
    internal sealed class AgentCore
    {
        public static readonly string[] KnownEngines = { "roslyn", "ironpython2", "cpython3" };
        private const int MaxJobsKept = 100, MaxTimeoutMs = 600000, MaxCodeBytes = 4 * 1024 * 1024;

        private readonly IUiBridge ui;
        private readonly IAgentEnvironment env;
        private readonly IEngineSource source;
        private IDictionary<string, EngineEntry> engines { get { return source.Engines; } }
        private readonly object stateLock = new object();
        private readonly Dictionary<string, Job> jobs = new Dictionary<string, Job>();
        private readonly List<string> order = new List<string>();
        private readonly Dictionary<string, EngineSession> sessions = new Dictionary<string, EngineSession>(StringComparer.OrdinalIgnoreCase);
        private readonly BlockingCollection<Job> queue = new BlockingCollection<Job>();
        private readonly ConcurrentQueue<Job> ready = new ConcurrentQueue<Job>();
        private int counter;
        // written by the worker thread, read by HTTP threads (Info): volatile
        private volatile Job current, stuck;
        private Thread worker;
        private volatile bool initialized;

        public AgentCore(IUiBridge ui, IAgentEnvironment env, IEngineSource source)
        {
            this.ui = ui; this.env = env; this.source = source;
        }

        /// <summary>Stops accepting jobs and lets the worker thread end (tests; the add-in never stops it).</summary>
        public void Stop() { queue.CompleteAdding(); }

        /// <summary>True when a compiled job is waiting for the UI thread.</summary>
        public bool HasReadyJob { get { return !ready.IsEmpty; } }

        /// <summary>Starts the worker thread. Call once.</summary>
        public void Start()
        {
            if (initialized) return;
            initialized = true;
            ConsoleRouter.Install();
            AuditLog.Prune(env.AuditDirectory);
            worker = new Thread(WorkerLoop) { IsBackground = true, Name = "revit-ballet-agent" };
            worker.Start();
            AgentLog.Write($"[AGENT] ready, engines: {(engines.Count == 0 ? "none installed" : string.Join(", ", engines.Keys))}");
        }

        // ---------------------------------------------------------------- submit / query

        private readonly Dictionary<string, string> requestIds = new Dictionary<string, string>();   // requestId -> job id

        public Job Submit(string engine, string code, int timeoutMs, string session, string client = null, string requestId = null)
        {
            if (!string.IsNullOrEmpty(requestId))
                lock (stateLock)
                {
                    string existing; Job j;
                    if (requestIds.TryGetValue(requestId, out existing) && jobs.TryGetValue(existing, out j)) return j;   // client retry: same job, not run twice
                }
            if (string.IsNullOrWhiteSpace(engine)) throw new AgentException(400, "bad_request", "\"engine\" is required");
            if (code == null) throw new AgentException(400, "bad_request", "\"code\" is required");
            if (System.Text.Encoding.UTF8.GetByteCount(code) > MaxCodeBytes) throw new AgentException(413, "too_large", "code exceeds 4 MiB");
            if (timeoutMs <= 0) timeoutMs = 30000;
            if (timeoutMs > MaxTimeoutMs) throw new AgentException(400, "bad_request", "timeoutMs exceeds " + MaxTimeoutMs);
            EngineEntry e;
            if (!engines.TryGetValue(engine, out e))
                throw new AgentException(KnownEngines.Contains(engine, StringComparer.OrdinalIgnoreCase) ? 409 : 404, "engine",
                    KnownEngines.Contains(engine, StringComparer.OrdinalIgnoreCase) ? $"engine '{engine}' is not installed" : $"unknown engine '{engine}'");
            if (e.LoadError != null) throw new AgentException(409, "engine", $"engine '{engine}' failed to load: {e.LoadError}");
            if (!initialized) throw new AgentException(503, "internal", "agent host not initialised");

            var job = new Job
            {
                Id = "j-" + Interlocked.Increment(ref counter).ToString("D4"),
                Engine = e.Manifest.Name, SessionName = string.IsNullOrWhiteSpace(session) ? null : session,
                Code = code, TimeoutMs = timeoutMs, Client = client, RequestId = string.IsNullOrEmpty(requestId) ? null : requestId,
            };
            job.Writer = new JobWriter(job);
            lock (stateLock)
            {
                jobs[job.Id] = job; order.Add(job.Id);
                if (job.RequestId != null) requestIds[job.RequestId] = job.Id;
                while (order.Count > MaxJobsKept)
                {   // drop the oldest finished job
                    var old = order.FirstOrDefault(id => jobs[id].IsFinal);
                    if (old == null) break;
                    var oldReq = jobs[old].RequestId; if (oldReq != null) requestIds.Remove(oldReq);
                    order.Remove(old); jobs.Remove(old);
                }
            }
            queue.Add(job);
            return job;
        }

        public Job Get(string id) { lock (stateLock) { Job j; return jobs.TryGetValue(id, out j) ? j : null; } }
        public List<Job> List() { lock (stateLock) return order.Select(id => jobs[id]).ToList(); }

        public bool WaitFinal(Job job, int waitMs)
        {
            return waitMs <= 0 ? job.Done.IsSet : job.Done.Wait(Math.Min(waitMs, MaxTimeoutMs));
        }

        public void Cancel(Job job)
        {
            lock (job.Sync)
            {
                if (job.IsFinal) return;
                job.CancelRequested = true;
                if (!job.UiStarted.IsSet)
                {   // not running yet: remove it
                    job.Abandoned = true;
                    job.State = JobState.Cancelled; job.ErrorKind = "cancelled"; job.ErrorMessage = "cancelled before it started";
                    job.Finished = DateTime.UtcNow; job.Done.Set();
                }
            }
            job.Cts.Cancel();
            if (job.Done.IsSet)
            {
                job.AddEvent("state", new JObject { ["state"] = "cancelled" });
                job.AddEvent("done", job.ToJson());
                AuditLog.Write(job, env.AuditDirectory);
            }
            else if (job.EngineInstance != null && job.Context != null)
            {
                try { job.EngineInstance.Abort(job.Context); } catch (Exception ex) { AgentLog.Write("[AGENT] abort failed: " + ex.Message); }
            }
        }

        // ---------------------------------------------------------------- sessions

        public List<EngineSession> ListSessions() { lock (stateLock) return sessions.Values.ToList(); }

        private static string SessionKey(string engine, string name) => engine + "/" + name;

        private EngineSession GetSession(Job job)
        {
            if (job.SessionName == null) return new EngineSession();
            lock (stateLock)
            {
                var key = SessionKey(job.Engine, job.SessionName);
                EngineSession s;
                if (!sessions.TryGetValue(key, out s)) sessions[key] = s = new EngineSession { Name = job.SessionName };
                return s;
            }
        }

        public bool DropSession(string engine, string name)
        {
            EngineSession s;
            lock (stateLock)
            {
                var key = SessionKey(engine, name);
                if (!sessions.TryGetValue(key, out s)) return false;
                sessions.Remove(key);
            }
            EngineEntry e;
            if (engines.TryGetValue(engine, out e) && e.Instance != null) { try { e.Instance.ResetSession(s); } catch { } }
            return true;
        }

        // ---------------------------------------------------------------- info

        public JObject Info(string sessionId, int port)
        {
            var modal = env.ModalDialogTitle;
            var responsive = env.UiResponsive;
            Job running; int queued; double oldest = 0;
            lock (stateLock)
            {
                running = jobs.Values.FirstOrDefault(j => j.State == JobState.Running)
                          ?? (stuck != null && !stuck.UiDone.IsSet ? stuck : null);
                var waiting = jobs.Values.Where(j => j.State == JobState.Queued || j.State == JobState.Compiling).ToList();
                queued = waiting.Count;
                if (waiting.Count > 0) oldest = (DateTime.UtcNow - waiting.Min(j => j.Submitted)).TotalMilliseconds;
            }
            var engArr = new JArray();
            foreach (var name in KnownEngines.Concat(engines.Keys).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                EngineEntry e;
                if (engines.TryGetValue(name, out e))
                {
                    var o = new JObject
                    {
                        ["name"] = e.Manifest.Name, ["language"] = e.Manifest.Language,
                        ["version"] = e.Instance != null ? e.Instance.Version : e.Manifest.Version,
                        ["available"] = e.LoadError == null, ["sessions"] = e.Manifest.Sessions,
                    };
                    if (e.LoadError != null) o["reason"] = e.LoadError;
                    engArr.Add(o);
                }
                else engArr.Add(new JObject { ["name"] = name, ["available"] = false, ["reason"] = "engine not installed" });
            }
            return new JObject
            {
                ["protocol"] = 1,
                ["revit"] = new JObject { ["year"] = env.RevitYear, ["version"] = env.RevitVersion },
                ["session"] = new JObject { ["id"] = sessionId, ["pid"] = int.TryParse(sessionId, out var pid) ? pid : 0, ["port"] = port },
                // idle: nothing running; busy: a job (or other work) is running; modal: a dialog is open; unresponsive: message pump hung
                ["ui"] = new JObject { ["responsive"] = responsive, ["modalDialog"] = modal, ["state"] = !responsive ? "unresponsive" : modal != null ? "modal" : running != null ? "busy" : "idle" },
                ["queue"] = new JObject { ["queued"] = queued, ["running"] = running?.Id, ["oldestQueuedMs"] = Math.Round(oldest) },
                ["engines"] = engArr,
            };
        }

        // ---------------------------------------------------------------- worker (compile, dispatch, watch)

        private void WorkerLoop()
        {
            foreach (var job in queue.GetConsumingEnumerable())
            {
                try { Process(job); }
                catch (Exception ex)
                {
                    AgentLog.Write("[AGENT] internal error: " + ex);
                    Finish(job, JobState.Failed, "internal", ex.Message, ex.StackTrace);
                }
            }
        }

        private void Process(Job job)
        {
            lock (job.Sync) { if (job.IsFinal) return; }
            // a previous job that timed out may still occupy the UI thread: keep this one queued until it is free
            while (stuck != null && !stuck.UiDone.IsSet)
            {
                if (job.IsFinal) return;
                Thread.Sleep(100);
            }
            stuck = null;
            lock (job.Sync) { if (job.IsFinal) return; }

            var entry = engines[job.Engine];
            var engine = source.Load(entry);
            if (engine == null) { Finish(job, JobState.Failed, "engine", $"engine '{job.Engine}' failed to load: {entry.LoadError}", null); return; }
            job.EngineInstance = engine;
            current = job;
            var session = GetSession(job);

            job.QueueMs = (DateTime.UtcNow - job.Submitted).TotalMilliseconds;
            job.SetState(JobState.Compiling);
            var t0 = DateTime.UtcNow;
            EngineCompileResult cr;
            try { cr = engine.Compile(job.Code, session, job.Cts.Token); }
            catch (Exception ex) { Finish(job, JobState.Failed, "engine", ex.GetType().Name + ": " + ex.Message, ex.StackTrace); return; }
            job.CompileMs = (DateTime.UtcNow - t0).TotalMilliseconds;
            lock (job.Sync) job.Diagnostics.AddRange(cr.Diagnostics);
            if (job.IsFinal) return;   // cancelled while compiling
            if (!cr.Success)
            {
                var first = cr.Diagnostics.FirstOrDefault(d => d.Severity == "error");
                Finish(job, JobState.Failed, "compile", first != null ? first.Message : "compilation failed", null);
                return;
            }

            job.Compiled = cr.Compiled;
            job.Context = new EngineContext { Out = job.Writer, Cancel = job.Cts.Token, Session = job.SessionName == null ? null : session };
            job.SessionObject = session;
            job.SetState(JobState.Queued);   // compiled, waiting for the UI thread
            ready.Enqueue(job);
            ui.RaiseDrain();

            // wait for the UI thread to pick it up
            var dispatched = DateTime.UtcNow; var lastRaise = dispatched;
            while (!job.UiStarted.Wait(50))
            {
                if (job.IsFinal) return;   // cancelled
                // Revit services external events only when it is idle; keep asking (the Idling handler is the second drain path)
                if ((DateTime.UtcNow - lastRaise).TotalMilliseconds > 1000) { lastRaise = DateTime.UtcNow; try { ui.RaiseDrain(); } catch { } }
                if ((DateTime.UtcNow - dispatched).TotalMilliseconds > job.TimeoutMs)
                {
                    lock (job.Sync) { if (job.UiStarted.IsSet) break; job.Abandoned = true; }
                    var modal = env.ModalDialogTitle;
                    if (modal != null) Finish(job, JobState.Timeout, "blocked_by_dialog", $"Revit is blocked by the open dialog '{modal}'; the script never started", null);
                    else Finish(job, JobState.Timeout, "timeout", $"the UI thread did not pick up the script within {job.TimeoutMs} ms", null);
                    return;
                }
            }

            // wait for it to finish running
            var started = DateTime.UtcNow;
            while (!job.UiDone.Wait(50))
            {
                if ((DateTime.UtcNow - started).TotalMilliseconds > job.TimeoutMs)
                {
                    job.Cts.Cancel();
                    try { engine.Abort(job.Context); } catch { }
                    stuck = job;
                    var modal = env.ModalDialogTitle;
                    if (modal != null) Finish(job, JobState.Timeout, "blocked_by_dialog", $"the script is waiting on the open dialog '{modal}'", null);
                    else Finish(job, JobState.Timeout, "timeout", $"the script did not finish within {job.TimeoutMs} ms (cancellation requested)", null);
                    return;
                }
            }
            Complete(job);
        }

        /// <summary>The UI thread has returned from Run: turn its result into the final job state.</summary>
        private void Complete(Job job)
        {
            var rr = job.RunResult;
            if (rr == null) { Finish(job, JobState.Failed, "internal", "no result from the engine", null); return; }
            lock (job.Sync) job.Diagnostics.AddRange(rr.Diagnostics);
            if (rr.Ok) Finish(job, JobState.Done, null, null, null);
            else if (job.CancelRequested || rr.ErrorKind == "cancelled")
                Finish(job, JobState.Cancelled, "cancelled", rr.ErrorMessage ?? "cancelled", rr.ErrorStack);
            else
                Finish(job, JobState.Failed, rr.ErrorKind ?? "runtime", rr.ErrorMessage, rr.ErrorStack);
        }

        private void Finish(Job job, JobState state, string kind, string message, string stack)
        {
            lock (job.Sync)
            {
                if (job.IsFinal) return;
                job.State = state; job.ErrorKind = kind; job.ErrorMessage = message; job.ErrorStack = stack;
                job.Finished = DateTime.UtcNow; job.Done.Set();
            }
            // A finished job is only history: do not keep the compiled script alive. Its compilation (symbols, metadata) is huge, and
            // the last 100 jobs are retained for GET /v1/jobs.
            job.Compiled = null; job.Context = null; job.RunResult = null; job.EngineInstance = null;
            if (job.SessionObject != null) job.SessionObject.Jobs++;
            job.AddEvent("state", new JObject { ["state"] = Job.StateName(state) });
            job.AddEvent("done", job.ToJson());
            AuditLog.Write(job, env.AuditDirectory);
            AgentLog.Write($"[AGENT] {job.Id} {job.Engine} -> {Job.StateName(state)}{(kind != null ? " (" + kind + ")" : "")}");
        }

        // ---------------------------------------------------------------- UI thread

        /// <summary>Convert the result on the UI thread (Revit objects), then cap it so a huge value cannot bloat memory or the response.</summary>
        private void ApplyValue(Job job, EngineRunResult rr)
        {
            string type, text; JToken json; bool truncated;
            try { ValueConverter.Convert(rr.Value, out type, out json, out text, out truncated); }
            catch (Exception ex) { type = rr.Value?.GetType().FullName; json = null; text = "<value conversion failed: " + ex.Message + ">"; truncated = false; }
            if (json != null && json.ToString(Newtonsoft.Json.Formatting.None).Length > Job.MaxInline) { json = null; truncated = true; }
            if (text != null && text.Length > Job.MaxInline) { text = text.Substring(0, Job.MaxInline); truncated = true; }
            lock (job.Sync) { job.ValueType = type; job.ValueJson = json; job.ValueText = text; job.ValueTruncated = truncated; }
        }

        /// <summary>Runs on the UI thread (in Revit: from the external event or the Idling handler). One job per call.</summary>
        public void RunReadyJob(object uiApp, object uiDoc, object doc)
        {
            Job job;
            while (ready.TryDequeue(out job))
            {
                lock (job.Sync)
                {
                    if (job.Abandoned || job.IsFinal) continue;
                    job.State = JobState.Running; job.Started = DateTime.UtcNow;
                }
                job.AddEvent("state", new JObject { ["state"] = "running" });
                job.UiStarted.Set();

                var ctx = job.Context;
                ctx.UIApp = uiApp; ctx.UIDoc = uiDoc; ctx.Doc = doc;
                var t0 = DateTime.UtcNow;
                EngineRunResult rr;
                using (ConsoleRouter.Scope(job.Writer))
                {
                    try { rr = job.EngineInstance.Run(job.Compiled, ctx); }
                    catch (Exception ex)
                    {
                        rr = new EngineRunResult { Ok = false, ErrorKind = "engine", ErrorMessage = ex.GetType().Name + ": " + ex.Message, ErrorStack = ex.StackTrace };
                    }
                }
                job.RunMs = (DateTime.UtcNow - t0).TotalMilliseconds;
                if (rr.Ok) ApplyValue(job, rr);
                job.RunResult = rr;
                job.UiDone.Set();
                break;   // one job per tick
            }
            if (!ready.IsEmpty) ui.RaiseDrain();
        }
    }
}
