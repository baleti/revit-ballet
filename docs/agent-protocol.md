# Agent protocol (v1)

How external tools (scripts, CI, LLM agents) talk to a running Revit session through Revit Ballet.
Protocol version: **1**. Status: **draft, implementation in progress** (the contract assembly and job model exist; the /v1 endpoints and engines are not wired in yet). This document is the contract; `commands/Agent/` implements the host,
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
  "ui": {"responsive": true, "modalDialog": null},
  "queue": {"queued": 0, "running": null},
  "engines": [
    {"name": "roslyn", "language": "csharp", "version": "3.11.0", "available": true, "sessions": true},
    {"name": "ironpython2", "language": "python", "version": "2.7.12", "available": true, "sessions": true},
    {"name": "cpython3", "language": "python", "version": "3.12.4", "available": false,
     "reason": "engine not installed"}
  ]
}
```

`ui.responsive` is `false` when the UI thread has not picked up a probe within 2 seconds.
`ui.modalDialog` is the title of an open modal dialog of the Revit process, or `null`.

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

Response: `202` with the job (state `queued`) when it did not finish within `wait`, else `200` with the finished job.

### `GET /v1/jobs/{id}?wait=ms`

Returns the job; `wait` long-polls until it is finished (or the wait elapses). `GET /v1/jobs` lists the last 100.

### `GET /v1/jobs/{id}/events`

Server-sent events: `state` (`{"state":"running"}`), `output` (`{"text":"..."}` for each write to the script's
output), `done` (the final job object). The stream ends after `done`.

### `DELETE /v1/jobs/{id}`

Cancel. A queued job is removed (state `cancelled`). A running job gets its cancellation token set and the
engine's abort hook called; see "Cancellation" for what that can and cannot do. Response: the job.

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
  `cancelled`, `engine` (engine failed to load or crashed), `internal`.
- `diagnostics`: compile errors and warnings with 1-based `line`/`column` where the engine can give them.

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

## Timeouts, cancellation, blocked UI

- The run time limit starts when the job starts running. On expiry the job is marked `timeout` (or
  `blocked_by_dialog`), its cancellation token is set and the engine abort hook runs.
- A script that never checks `Cancel` and never returns cannot be stopped from outside: the UI thread stays busy and
  later jobs stay `queued`. This is visible: `/v1/info` reports `ui.responsive:false` and the stuck job in
  `queue.running`. Python engines also check `cancel` between statements (trace hook), so pure-Python loops are stoppable.
- If a modal dialog is open, jobs cannot start at all (the external event does not run). They time out after
  `timeoutMs` counted from submission in that case, with kind `blocked_by_dialog`.

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

Engines: `roslyn` (C# scripting, Roslyn), `ironpython2` (IronPython 2.7), `cpython3` (a real CPython 3
embedded through pythonnet; ships its own embeddable Python runtime, so it needs no system Python).
They can be installed independently; an uninstalled engine shows `available:false` in `/v1/info`.

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
