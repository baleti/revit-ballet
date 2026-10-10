using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Autodesk.Revit.UI;
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
    /// The agent host: job queue, worker thread, engines and sessions. Protocol: docs/agent-protocol.md.
    /// Pipeline per job: worker thread compiles, the UI thread (Revit external event) runs, the worker watches timeouts. One job at a time,
    /// strictly in submission order, so named sessions see their own state in order.
    /// </summary>
    internal static class AgentHost
    {
        public static readonly string[] KnownEngines = { "roslyn", "ironpython2", "cpython3" };
        private const int MaxJobsKept = 100, MaxTimeoutMs = 600000, MaxCodeBytes = 4 * 1024 * 1024;

        private static Dictionary<string, EngineEntry> engines = new Dictionary<string, EngineEntry>();
        private static readonly object stateLock = new object();
        private static readonly Dictionary<string, Job> jobs = new Dictionary<string, Job>();
        private static readonly List<string> order = new List<string>();
        private static readonly Dictionary<string, EngineSession> sessions = new Dictionary<string, EngineSession>(StringComparer.OrdinalIgnoreCase);
        private static readonly BlockingCollection<Job> queue = new BlockingCollection<Job>();
        private static readonly ConcurrentQueue<Job> ready = new ConcurrentQueue<Job>();
        private static ExternalEvent uiEvent;
        private static int counter;
        private static Job current, stuck;
        private static Thread worker;
        private static bool initialized;

        /// <summary>Must be called from a Revit API context (ExternalEvent.Create requirement).</summary>
        public static void Initialize()
        {
            if (initialized) return;
            initialized = true;
            ConsoleRouter.Install();
            engines = EngineLoader.Discover();
            uiEvent = ExternalEvent.Create(new AgentExecHandler());
            worker = new Thread(WorkerLoop) { IsBackground = true, Name = "revit-ballet-agent" };
            worker.Start();
            RevitBalletServer.LogToRevit($"[AGENT] ready, engines: {(engines.Count == 0 ? "none installed" : string.Join(", ", engines.Keys))}");
        }

        // ---------------------------------------------------------------- submit / query

        public static Job Submit(string engine, string code, int timeoutMs, string session)
        {
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
                Code = code, TimeoutMs = timeoutMs,
            };
            job.Writer = new JobWriter(job);
            lock (stateLock)
            {
                jobs[job.Id] = job; order.Add(job.Id);
                while (order.Count > MaxJobsKept)
                {   // drop the oldest finished job
                    var old = order.FirstOrDefault(id => jobs[id].IsFinal);
                    if (old == null) break;
                    order.Remove(old); jobs.Remove(old);
                }
            }
            queue.Add(job);
            return job;
        }

        public static Job Get(string id) { lock (stateLock) { Job j; return jobs.TryGetValue(id, out j) ? j : null; } }
        public static List<Job> List() { lock (stateLock) return order.Select(id => jobs[id]).ToList(); }

        public static bool WaitFinal(Job job, int waitMs)
        {
            return waitMs <= 0 ? job.Done.IsSet : job.Done.Wait(Math.Min(waitMs, MaxTimeoutMs));
        }

        public static void Cancel(Job job)
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
            }
            else if (job.EngineInstance != null && job.Context != null)
            {
                try { job.EngineInstance.Abort(job.Context); } catch (Exception ex) { RevitBalletServer.LogToRevit("[AGENT] abort failed: " + ex.Message); }
            }
        }

        // ---------------------------------------------------------------- sessions

        public static List<EngineSession> ListSessions() { lock (stateLock) return sessions.Values.ToList(); }

        private static string SessionKey(string engine, string name) => engine + "/" + name;

        private static EngineSession GetSession(Job job)
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

        public static bool DropSession(string engine, string name)
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

        public static JObject Info(string sessionId, int port)
        {
            // From the Revit executable itself: independent of when the UIApplication becomes available. 24.x.y.z = Revit 2024.
            var fvi = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileVersionInfo;
            var year = 2000 + fvi.FileMajorPart;
            var modal = UiProbe.ModalDialogTitle();
            Job running; int queued;
            lock (stateLock)
            {
                running = jobs.Values.FirstOrDefault(j => j.State == JobState.Running)
                          ?? (stuck != null && !stuck.UiDone.IsSet ? stuck : null);
                queued = jobs.Values.Count(j => j.State == JobState.Queued || j.State == JobState.Compiling);
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
                ["revit"] = new JObject { ["year"] = year, ["version"] = fvi.FileVersion },
                ["session"] = new JObject { ["id"] = sessionId, ["pid"] = int.Parse(sessionId), ["port"] = port },
                ["ui"] = new JObject { ["responsive"] = UiProbe.IsResponsive(), ["modalDialog"] = modal },
                ["queue"] = new JObject { ["queued"] = queued, ["running"] = running?.Id },
                ["engines"] = engArr,
            };
        }

        // ---------------------------------------------------------------- worker (compile, dispatch, watch)

        private static void WorkerLoop()
        {
            foreach (var job in queue.GetConsumingEnumerable())
            {
                try { Process(job); }
                catch (Exception ex)
                {
                    RevitBalletServer.LogToRevit("[AGENT] internal error: " + ex);
                    Finish(job, JobState.Failed, "internal", ex.Message, ex.StackTrace);
                }
            }
        }

        private static void Process(Job job)
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
            var engine = EngineLoader.Load(entry);
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
            uiEvent.Raise();

            // wait for the UI thread to pick it up
            var dispatched = DateTime.UtcNow;
            while (!job.UiStarted.Wait(50))
            {
                if (job.IsFinal) return;   // cancelled
                if ((DateTime.UtcNow - dispatched).TotalMilliseconds > job.TimeoutMs)
                {
                    lock (job.Sync) { if (job.UiStarted.IsSet) break; job.Abandoned = true; }
                    var modal = UiProbe.ModalDialogTitle();
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
                    var modal = UiProbe.ModalDialogTitle();
                    if (modal != null) Finish(job, JobState.Timeout, "blocked_by_dialog", $"the script is waiting on the open dialog '{modal}'", null);
                    else Finish(job, JobState.Timeout, "timeout", $"the script did not finish within {job.TimeoutMs} ms (cancellation requested)", null);
                    return;
                }
            }
            Complete(job);
        }

        /// <summary>The UI thread has returned from Run: turn its result into the final job state.</summary>
        private static void Complete(Job job)
        {
            var rr = job.RunResult;
            if (rr == null) { Finish(job, JobState.Failed, "internal", "no result from the engine", null); return; }
            lock (job.Sync) job.Diagnostics.AddRange(rr.Diagnostics);
            if (rr.Ok)
            {
                string type, text; JToken json;
                ValueConverter.Convert(rr.Value, out type, out json, out text);
                lock (job.Sync) { job.ValueType = type; job.ValueJson = json; job.ValueText = text; }
                Finish(job, JobState.Done, null, null, null);
            }
            else if (job.CancelRequested || rr.ErrorKind == "cancelled")
                Finish(job, JobState.Cancelled, "cancelled", rr.ErrorMessage ?? "cancelled", rr.ErrorStack);
            else
                Finish(job, JobState.Failed, rr.ErrorKind ?? "runtime", rr.ErrorMessage, rr.ErrorStack);
        }

        private static void Finish(Job job, JobState state, string kind, string message, string stack)
        {
            lock (job.Sync)
            {
                if (job.IsFinal) return;
                job.State = state; job.ErrorKind = kind; job.ErrorMessage = message; job.ErrorStack = stack;
                job.Finished = DateTime.UtcNow; job.Done.Set();
            }
            if (job.SessionObject != null) job.SessionObject.Jobs++;
            job.AddEvent("state", new JObject { ["state"] = Job.StateName(state) });
            job.AddEvent("done", job.ToJson());
            RevitBalletServer.LogToRevit($"[AGENT] {job.Id} {job.Engine} -> {Job.StateName(state)}{(kind != null ? " (" + kind + ")" : "")}");
        }

        // ---------------------------------------------------------------- UI thread

        /// <summary>Runs on the Revit UI thread (external event).</summary>
        internal static void RunReadyJob(UIApplication app)
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

                var uidoc = app.ActiveUIDocument;
                var ctx = job.Context;
                ctx.UIApp = app; ctx.UIDoc = uidoc; ctx.Doc = uidoc?.Document;
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
                job.RunResult = rr;
                job.UiDone.Set();
                break;   // one job per tick
            }
            if (!ready.IsEmpty) uiEvent.Raise();
        }
    }

    internal sealed class AgentExecHandler : IExternalEventHandler
    {
        public void Execute(UIApplication app) { AgentHost.RunReadyJob(app); }
        public string GetName() => "RevitBalletAgentExecHandler";
    }
}
