# Agent protocol (v1)

How external tools (scripts, CI, LLM agents) talk to a running Revit session through Revit Ballet.
Protocol version: **1**. Status: **implemented** (host, `/v1` API and the three engines; covered by the VM CI on Revit 2020-2026). This document is the contract; `commands/Agent/` implements the host,
`agents/` the script engines.

## Goals and non-goals

- One transport, many languages: C# (Roslyn), IronPython 2.7 and CPython 3 run behind the same API.
- Scripts run on Revit's UI thread (the Revit API requires it), so the protocol must make queueing,
  cancellation, timeouts and "Revit is blocked by a dialog" visible instead of silent.
- The host (the `revit-ballet` add-in) contains **no scripting engine**. Engines are separate, optionally
  installed components (own DLLs, own signing, loaded only when used). Rationale: antivirus products
  flag embedded compilers and interpreters; the less the main add-in contains, the less it is flagged.
- Non-goals: remote (non-loopback) access, sandboxing scripts (a script has the full power of the Revit
  process and of the user; the token is the boundary), multi-user auth.

## Transport and auth (unchanged)

- HTTPS on `127.0.0.1` (self-signed certificate), port picked from 23717-23817, one port per Revit process.
- Token in `%APPDATA%\revit-ballet\runtime\network\token`, sent as `X-Auth-Token: <token>` or
  `Authorization: Bearer <token>`. Missing or wrong token: `401`.
- Live sessions are listed in `%APPDATA%\revit-ballet\runtime\documents` (unchanged).
- All bodies are UTF-8 JSON unless stated. Errors use the shape `{"error":{"kind":"...","message":"..."}}`
  with an HTTP status: 400 bad request, 401 auth, 404 unknown job/session/engine, 409 engine unavailable or
  session busy, 413 body too large (default limit 4 MiB).

## Endpoints

### `GET /v1/info`

```json
{
  "protocol": 1,
  "revit": {"year": 2024, "version": "24.3.0.108", "build": "20240425_1500"},
  "session": {"id": "12345", "pid": 12345, "port": 23717},
  "ui": {"responsive": true, "modalDialog": null, "state": "idle"},
  "queue": {"queued": 0, "running": null, "oldestQueuedMs": 0},
  "engines": [
    {"name": "roslyn", "language": "csharp", "version": "3.11.0", "available": true, "sessions": true},
    {"name": "ironpython2", "language": "python", "version": "2.7.12", "available": true, "sessions": true},
    {"name": "cpython3", "language": "python", "version": "3.12.4", "available": false,
     "reason": "engine not installed"}
  ]
}
```

`ui.responsive` is `false` when the UI thread has not picked up a probe within 2 seconds.
`ui.modalDialog` is `null` or the open modal dialog: `{"title","kind","text","buttons":["OK","Cancel"],"dialogId","sinceMs"}`
(see "Modal dialogs"). `ui.state` is `idle`, `busy`, `modal` or `unresponsive`.

### `POST /v1/jobs`

Request:

```json
{"engine": "roslyn", "code": "Console.WriteLine(Doc.Title);", "timeoutMs": 30000, "session": null, "wait": 0}
```

| field | meaning |
|---|---|
| `engine` | required: `roslyn`, `ironpython2`, `cpython3` (see `/v1/info`) |
| `code` | required: source text |
| `timeoutMs` | run time limit (not counting queue time), default 30000, max 600000 |
| `session` | optional name of a persistent engine session (state kept between jobs); `null` = isolated |
| `wait` | milliseconds to hold the HTTP response open until the job finishes (0 to 600000). Default 0 |
| `dialogPolicy` | what to do when a modal dialog holds Revit up while the job waits or runs: `"report"` (default), `"fail"`, or `{"respond": "<button>"}` |
| `requestId` | optional idempotency key: a retry with the same id returns the existing job instead of running it twice |

Response: `202` with the job (state `queued`) when it did not finish within `wait`, else `200` with the finished job.

### `GET /v1/jobs/{id}?wait=ms`

Returns the job; `wait` long-polls until it is finished (or the wait elapses). `GET /v1/jobs` lists the last 100.

### `GET /v1/jobs/{id}/events`

Server-sent events: `state` (`{"state":"running"}`), `output` (`{"text":"..."}` for each write to the script's
output), `done` (the final job object). The stream ends after `done`.

### `DELETE /v1/jobs/{id}`

Cancel. A queued job is removed (state `cancelled`). A running job gets its cancellation token set and the
engine's abort hook called; see "Cancellation" for what that can and cannot do. Response: the job.

### `GET /v1/dialog`, `POST /v1/dialog/respond`

`GET /v1/dialog` returns `{"dialog": null}` or `{"dialog": {...}}`. `POST /v1/dialog/respond {"button": "OK"}` presses that button of the open
dialog (matched by caption, case-insensitive): `200 {"responded": true, ...}`, `404` when no dialog is open, `409` with the list of buttons when there is no
such button or pressing it did not close the dialog. Both work **while the job queue is stuck behind the dialog**: they need the dialog, not the UI thread.

### Sessions

`GET /v1/sessions` lists `{"engine","name","jobs","created"}`; `DELETE /v1/sessions/{engine}/{name}` drops the
state. Jobs that use the same session run in submission order and see each other's variables. A session name
is scoped to its engine.

## The job object

```json
{
  "id": "j-0007",
  "engine": "roslyn",
  "session": null,
  "state": "done",
  "ok": true,
  "submitted": "2026-10-10T14:00:00.000Z",
  "started":   "2026-10-10T14:00:00.020Z",
  "finished":  "2026-10-10T14:00:00.310Z",
  "output": "Model.rvt\n",
  "value": {"type": "System.Int32", "json": 42, "text": "42"},
  "error": null,
  "diagnostics": [{"severity": "warning", "id": "CS0219", "line": 3, "column": 5, "message": "..."}],
  "timings": {"queueMs": 20, "compileMs": 180, "runMs": 90}
}
```

- `state`: `queued`, `compiling`, `running`, `done`, `failed`, `cancelled`, `timeout`.
  `ok` is true only for `done`.
- `output`: everything the script wrote through its `Out` writer / `print` / `Console.WriteLine`
  (per job; it never contains output of other code in the process).
- `value`: the result. Roslyn: the script's return value; Python: the variable `result` if the script sets it.
  `json` is best-effort JSON (null when not serialisable), `text` is `ToString()`/`str()`.
- `error` (when not `ok`): `{"kind","message","stack"}`. Kinds: `compile` (see `diagnostics`), `runtime` (exception in the
  script), `timeout`, `blocked_by_dialog` (timeout while a modal dialog was open; `message` carries its title),
  `cancelled`, `engine` (engine failed to load or crashed), `engine_hung` (see Watchdog), `ui_busy` (the UI thread is still
  running an earlier job that timed out and could not be stopped), `internal`.
- `diagnostics`: compile errors and warnings with 1-based `line`/`column` where the engine can give them.
- `blockedBy`: `null`, or the dialog that has been holding this job up for more than 1.5 s (the job is still `queued` or `running`).
  `dialogs`: every distinct dialog seen while the job waited or ran, each with `firstSeen` and the `action` taken (`report`, `fail`, `responded:<button>`).
  When a job ends because of a dialog, `error.kind` is `blocked_by_dialog` and `error.dialog` holds it.

## Execution model

- One FIFO queue per Revit process. A Revit external event drains it on the UI thread, one job per tick.
  Jobs never overwrite each other (the old single pending slot is gone).
- Compilation happens on a worker thread (Roslyn) and does not occupy the UI thread. Execution is on the UI thread.
- Globals every engine provides (language-appropriate spelling):

| meaning | C# | Python |
|---|---|---|
| UIApplication | `UIApp` | `uiapp` (alias `UIApp`) |
| active UIDocument or null | `UIDoc` | `uidoc` (alias `UIDoc`) |
| active Document or null | `Doc` | `doc` (alias `Doc`) |
| output writer | `Out` (`Console.WriteLine` also works) | `out` (`print` also works) |
| cancellation | `Cancel` (`CancellationToken`) | `cancel` (object with `is_set()`) |

- No transaction is opened; scripts open their own, as before.

## Modal dialogs

A modal dialog stops the UI thread, so no script can run and an external event is not serviced. The agent therefore watches for one while a job waits or runs:

- **Detection**: a visible window owned by Revit's main window while the main window is disabled. Text and button captions come from UI Automation; for
  Revit's own TaskDialogs the `DialogBoxShowing` event adds the dialog id and message. `kind` is `task` (a Revit TaskDialog), `win32` (any other
  modal window) or `unknown` (Revit is in a modal state but no dialog window was found).
- **Reporting**: a dialog open for more than 1.5 s (so one that flashes by is ignored) appears in `ui.modalDialog`, in `GET /v1/dialog`, in the
  `blockedBy` and `dialogs` fields of every job it affects, as a `dialog` event on that job's SSE stream, in the audit log and in `server.log`. This
  happens while the job is still waiting or running, not after its timeout.
- **Policy** (`dialogPolicy` on the job): `report` keeps waiting until the job's own timeout (a person or `POST /v1/dialog/respond` may answer);
  `fail` ends the job at once as `failed`/`blocked_by_dialog` (the UI thread may stay blocked, so later jobs wait until the dialog is gone);
  `{"respond": "OK"}` presses that button after the grace period and carries on, and fails the job if the dialog has no such button.
- **Answering**: UI Automation Invoke, then a Win32 click, then a real mouse click; each is checked by seeing the dialog close. The mouse fallback needs an
  interactive desktop (a disconnected RDP session has none).
- **Not covered**: dialogs that are not modal (docked panes, balloon notifications) do not block anything and are not reported.

## Timeouts, cancellation, blocked UI

- The run time limit starts when the job starts running. On expiry the job is marked `timeout` (or
  `blocked_by_dialog`), its cancellation token is set and the engine abort hook runs.
- What can be stopped depends on the engine. `ironpython2` checks the cancellation token on every traced line, so even `while True: pass` stops.
  `cpython3` injects a `KeyboardInterrupt` into the running thread, which also stops pure-Python loops. `roslyn` can only stop cooperatively
  (a script that never checks `Cancel` runs to completion). A script blocked inside a Revit API call cannot be interrupted by any engine.
- A job that cannot be stopped keeps the UI thread busy and later jobs stay `queued`. This is visible: `/v1/info` reports `ui.responsive:false`
  (when the message pump hangs) and the stuck job in `queue.running`.
- If a modal dialog is open, jobs cannot start at all (the external event does not run). The job reports the dialog (see "Modal dialogs"); if
  nothing answers it, it times out after `timeoutMs` counted from submission, with kind `blocked_by_dialog`.

## Engines

An engine is a directory `engines/<name>/` next to `revit-ballet.dll` containing a manifest and its assemblies:

```json
{"name": "roslyn", "language": "csharp", "version": "3.11.0",
 "assembly": "revit-ballet-engine-roslyn.dll", "type": "RevitBallet.Engines.Roslyn.RoslynEngine",
 "protocol": 1, "sessions": true}
```

The host reads manifests at start-up (no engine code is loaded), and loads an engine's assembly the first
time a job asks for it. Engines implement the contract in `revit-ballet.agents.abstractions.dll`:

```csharp
public interface IScriptEngine {
    string Name { get; }
    EngineCompileResult Compile(string code, EngineSession session);            // worker thread
    EngineRunResult Run(object compiled, EngineContext ctx, EngineSession session); // UI thread
    void ResetSession(EngineSession session);
}
```

`EngineContext` carries `UIApp`, `UIDoc`, `Doc` (as `object`; engines cast to the Revit API types of the
running year), the per-job output writer and the cancellation token.

### The engines

| engine | runtime | notes |
|---|---|---|
| `roslyn` | Roslyn scripting 3.11 (Revit 2020-2024) / 4.8 (2025-2026) | C# script, same globals and imports as the old `/roslyn`. Sessions use script continuation (`ContinueWith`). `Return` value becomes `value`. |
| `ironpython2` | IronPython 2.7.12 | Python 2 syntax. Standard library from a bundled `Lib.zip`. `from Autodesk.Revit.DB import *` works without `clr.AddReference`. |
| `cpython3` | CPython 3.12 (official embeddable distribution, bundled) through pythonnet 3.0 | Real Python 3 (f-strings, type hints, modern stdlib). `import clr` then `clr.AddReference("RevitAPI")` as in pythonnet. No system Python needed. One interpreter per Revit process, shared by all jobs. |

Python specifics: the script's value is the variable `result` (plain Python values become JSON; anything else is reported through `str()`).
`out.write(...)` and `print(...)` both reach the job output. A Python traceback is returned in `error.stack`, syntax errors as `compile` diagnostics with line and column.

Install size: the Python 3 runtime adds about 22 MB, IronPython about 6 MB. The installer accepts `--engines all|none|name[,name...]`
(default `all`); engines that were installed earlier and are not in an explicit list are removed. All engines are separate files, so a user who does not want an
embedded interpreter on the machine can leave it out.

Known limitations: one job runs at a time per Revit process (the UI thread is the bottleneck by design); the `cpython3` interpreter lives for the lifetime of
the Revit process and cannot be re-initialised; other add-ins that embed IronPython or Python (pyRevit, RevitPythonShell) load their own copies, so versions are not shared.

## Compatibility

Kept as thin shims on the new core, same request and response shapes as before:

- `POST /roslyn` (raw C# body) -> a `roslyn` job with `wait = 30000`; response is the old `ScriptResponse`.
- `POST /query/...`, `GET /screenshot` unchanged.

New clients should use `/v1`. The shims are removed no earlier than protocol version 2.

## Security notes

- Same trust model as before: loopback only, bearer token, scripts are unsandboxed. Do not expose the port.
- Engines are loaded only from `engines/` under the add-in folder, which is writable only by the installing user;
  the host refuses a manifest whose assembly path escapes its directory.
- Release builds sign the host and each engine assembly separately.


## Watchdog

Compiling runs on its own thread, so a hung compiler never blocks the queue. A compile is declared hung when its thread has used
no CPU for 20 s (deadlock, waiting forever) or when it is still running after max(120 s, 2 x `timeoutMs`) even though it uses CPU
(endless loop). A long compile that keeps working is left alone. The job ends `failed` with `error.kind: engine_hung`; the message
says which rule tripped, the thread state and CPU time. The engine instance is dropped (the next job loads a fresh one), the
stuck thread is abandoned, `GET /v1/info` counts it in `hungCompiles`, and the details go to `runtime\server.log`. A queued job
that waits longer than its `timeoutMs` for a UI thread still held by an earlier timed-out job ends `ui_busy` (or `blocked_by_dialog`
when a dialog is the reason). `Abort` calls and audit writes cannot stall the worker either (2 s limit; audit is asynchronous).
