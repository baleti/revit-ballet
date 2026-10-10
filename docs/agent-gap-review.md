# Agent protocol: gap review against existing Revit agent tooling

Written 2026-10-10 against commit 6e0dcdb plus the uncommitted working tree (`commands/Agent/*`, `agents/*`).
Input: a sweep of what already exists for letting LLM agents drive Revit. Purpose: list what `docs/agent-protocol.md`
and the host do not cover yet, so you can close the gaps. Nothing here requires MCP, and we are not adopting any of these
projects. They are used only as a checklist of what people found they needed.

## What was surveyed

- Autodesk Revit Public MCP Server (Revit 2027, technical preview): read-only tools, stdio. Autodesk says write tools will come as a separate server.
- pyRevit 7 agent runtime (experimental, `pyrevit mcp` + in-Revit AgentHost over a named pipe, ExternalEvent, IronPython 2.7/3.4 + CPython 3.12).
  This is the closest design to ours and the best source of ideas: https://docs.pyrevitlabs.io/agent-runtime/
- Community: mcp-servers-for-revit/revit-mcp (archived, `send_code_to_revit` = C# snippet in an `Execute(Document)` template),
  LuDattilo/RevitCortex (173 tools, C# exec with a namespace block-list), Demolinator/revit-mcp-server, oakplank/RevitMCP,
  mk92at/pyrevit-mcp (all via pyRevit Routes, unauthenticated HTTP), schauh11, KenLP, PiggyAndrew.

## What we already do better or equal (do not redo)

- Persistent named sessions per engine (none of the surveyed projects have this).
- Three engines behind one API, engines as separate signed assemblies (AV-friendly).
- Real queue with job ids, long-poll, SSE output, cancel, timeouts, `blocked_by_dialog` and `ui.responsive` visibility.
- Per-job output capture (ConsoleRouter) and structured diagnostics with line/column.
- Multi-session discovery via the `runtime\documents` registry.

## Gaps, in priority order

### 1. Event stream to the client (nobody else has this; it is a stated goal)

The spec has job-scoped SSE only. There is no way for an agent to learn that something happened in Revit.
pyRevit arms events internally during a run but exposes nothing to the agent, so this is open ground.

Suggested shape:

- `GET /v1/events?types=documentChanged,viewActivated,...&since=<seq>` SSE, global (not per job). Every event has a monotonically
  increasing `seq`; keep a ring buffer (e.g. 5000) so a client can resume with `Last-Event-ID`/`since`. Report a `gap` event if `since` fell out of the buffer.
- Event types worth having: `documentOpened`, `documentClosing`, `documentClosed`, `documentSaving`, `documentSaved`, `documentSynchronized`,
  `documentChanged` (added/modified/deleted ElementIds + transaction names + originating job id), `viewActivated`, `selectionChanged`
  (native event on newer years; on older years detect via the Idling handler comparing selection ids), `dialogShowing`, `failuresProcessing`
  (warnings/errors text), `idling` (opt-in only, coalesced and rate limited), `applicationClosing`, `jobStateChanged`, `engineLoaded`.
- Payloads must be small: ids and counts, never element dumps. The agent pulls details with a job.
- Tag changes caused by our own jobs with the job id (set a "current job" marker while `Run` executes, read it in the DocumentChanged handler).
  Without this an agent reacting to events will loop on its own edits.
- Revit constraint to document: inside `DocumentChanged` and most event handlers the document is read-only. A reaction that must modify
  the model has to go through the normal job queue (ExternalEvent), never run inline in the handler.
- Subscriptions are per client connection, not global, so a dropped client does not leave handlers armed. Register Revit handlers lazily
  (first subscriber) and unregister at zero subscribers to keep the idle cost nil.
- Stretch: `POST /v1/triggers` registers a named script (engine + code or stored script name) that runs as a job when a matching event fires, with
  a debounce and a max-rate. This is the "respond to events" feature without a client having to stay connected. Triggers must show in `/v1/info`
  and be removable, and should be off by default after Revit restart unless marked persistent.

### 2. Change tracking and a transaction safety net per job (pyRevit's best idea)

Our jobs return output/value only. The agent has no machine-readable answer to "what did my script just change".
Add an optional job field `mode`:

- `"raw"` (current behaviour, default for compatibility): no wrapping.
- `"query"`: the engine runs the script inside a `TransactionGroup` that is always rolled back (guards a read-only intent; not a sandbox, say so).
- `"modify"`: runs in a `TransactionGroup`, assimilated into one undo entry named `Agent: <label>` on success; rolled back on error or cancel.
- `"dryRun"`: like modify but always rolled back; still returns the change report so the agent can preview.

And always (even in `raw`) return `changes: {added:[], modified:[], deleted:[], truncated:false}` collected from `DocumentChanged` while the job ran,
plus `failures: [{severity, message, elementIds}]` from `FailuresProcessing` (let the job option `failurePolicy: "report"|"delete-warnings"` decide whether warnings are swallowed).
Add `label` (short human title, used for the undo entry and in the job list).
Note this interacts with scripts that open their own transactions: a TransactionGroup around them is fine, but a script that starts its own TransactionGroup
and leaves it open must be detected and rolled back as an error.

Optional approval gate (human in the loop), as pyRevit does: a per-session/server policy `ask` that, after a modify job, shows the changed elements
(select + temp isolate) and asks Keep/Discard in Revit. Only worth building if you want agents running unattended on real project data; keep it out of v1 otherwise.

### 3. Modal dialogs: report content and allow a response

The host detects a modal dialog and reports its title, but cannot do anything about it, and the queue is dead while one is open
(the ExternalEvent does not fire). pyRevit captures dialogs and dismisses them during runs.

- `/v1/info.ui.modalDialog` should be an object: `{title, kind: "task"|"win32", text, buttons:[...], dialogId}`. For Revit TaskDialogs get this from the
  `DialogBoxShowing` event (`TaskDialogShowingEventArgs.Message/DialogId`); for plain Win32 dialogs use the HWND already found by `UiProbe`
  (read child static text and button captions).
- `POST /v1/dialog/respond {"button": "..."}`: for TaskDialogs use `OverrideResult` in the event handler (arm it only when a respond request is pending);
  for Win32 dialogs send a BM_CLICK to the matching button. This is the only way an agent can recover from `blocked_by_dialog` without a human.
- Per-job option `dialogPolicy`: `"fail"` (default, job fails with kind `blocked_by_dialog` and the dialog text), `"dismiss"` (auto-answer a configured default), or a map `{dialogId: result}`.
  Always log every auto-answered dialog in the job object (`dialogs: [...]`).
- Make this fully independent from the job queue, because it must work exactly when the queue is stuck.

### 4. ExternalEvent liveness when Revit is idle, minimised or mid-command

A known Revit behaviour: `ExternalEvent.Raise()` is only serviced when Revit is idle, and in practice it can sit unserviced while Revit is
in the background or the user is in the middle of an interactive command (placing an element, sketch mode, in-place editing), sometimes until the user moves the mouse.
The current design relies purely on `uiEvent.Raise()`.

- Add an `Idling` handler as a second drain path (and as a keep-alive that re-raises the event while `ready` is non-empty). Keep it subscribed only while the queue is non-empty.
- Expose in `/v1/info`: `queue.oldestQueuedMs` (age of the oldest waiting job) and `ui.state: "idle"|"busy"|"modal"|"unresponsive"`, plus a best-effort `ui.inCommand` if you can detect an active interactive command, so the agent can tell "waiting for the user" from "broken".
- Decide and document what happens to a job submitted while the user is in a sketch/edit mode (Revit refuses many API calls there). Return a clear error kind (`revit_in_edit_mode`) rather than a generic runtime exception.

### 5. Self-description for the agent (API discovery and context), without extra tools

pyRevit ships `get_context`, `lookup_revit_api` (reflection over the running Revit's assemblies) and "skills" (markdown task guides). We have none of that, so each agent
burns tokens guessing. Because we are not using MCP these should be plain endpoints and a helper prelude:

- `GET /v1/context`: active document (title, path, central/local, workshared, units, project info), open documents in this session, links,
  active view (id, name, type, level, scale, phase), selection (ids + categories, capped), levels, worksets in use, current user, Revit year/build, add-ins loaded. One call, no script needed.
- `GET /v1/api/type?name=Autodesk.Revit.DB.Wall` and `/v1/api/search?q=...`: reflection over the loaded `RevitAPI.dll`/`RevitAPIUI.dll` (members, signatures, obsolete flags, and the XML doc comments from `RevitAPI.xml` if present beside the DLL).
  Doing it against the running assembly gives year-correct answers, which matters across 2020-2026 (for example `ElementId` is `long` from 2024, `BuiltInParameterGroup` vs `GroupTypeId`, `ForgeTypeId` units).
- Roslyn engine: expose completion and quick-info (`POST /v1/roslyn/complete {code, position}`), cheap to add and very useful for agents writing C#.
- Helper prelude in every engine, short and documented: `Tx(name, () => ...)` (transaction with auto-rollback on throw), `Q(category)` collector shortcut,
  `Inspect(id)` returning id/category/type/level/parameters as a dictionary, `Params(el)` with display and internal values and units, `Mm(x)`/`ToMm()` unit helpers,
  `Select(ids)` / `Show(ids)` routed through `SelectionModeManager` (our own rule in CLAUDE.md), `Link(name)` for linked documents. Keeps agent code short and consistent.
- `GET /v1/skills` and `GET /v1/skills/{name}`: serve markdown task guides from `docs/agent-skills/` (unit conversion, transactions, worksharing, linked models, ElementId type changes by year, selection manager rules, which operations are unsafe on the UI thread). CLAUDE.md already holds some of this; make it addressable so a fresh agent can pull exactly what it needs.
- Saved scripts ("automations"): `PUT /v1/scripts/{name} {engine, code, params-schema}` and `POST /v1/scripts/{name}/run {args}`. Cheaper than resending big scripts, and lets us ship vetted helpers. Store under `%APPDATA%\revit-ballet\scripts\`.

### 6. Visual feedback that does not depend on the window

`/screenshot` captures the Revit window, so it is wrong when Revit is minimised, covered, on another desktop, or showing a dialog.
pyRevit's `capture_view` renders a chosen view. Add `GET /v1/view/{id}/image?w=1600&format=png` using `Document.ExportImage` with `ImageExportOptions`
(`ExportRange.SetOfViews`) into the runtime dir, returned as the body. Must go through the UI thread queue like any job. Also `?elementIds=` to temporarily isolate or colour specific elements for "show me what you mean" checks (use a temporary view mode, never modify the view permanently).

### 7. Output and result handling limits

- Cap `output` and `value.json` per job (for example 256 KiB each) and set `outputTruncated` / `valueTruncated`; serve the full text at `GET /v1/jobs/{id}/output?offset=&limit=`. A runaway `foreach` printing every element otherwise bloats memory (100 jobs kept) and the response.
- `ValueConverter`: make sure it handles cycles, `ElementId`, `Element` (return `{id, category, name}` not a reflected graph), `XYZ`/`BoundingBoxXYZ`, `Parameter`, `IEnumerable<T>` with a length cap, and records the cap. Revit API objects reflected naively can be enormous or throw on property access.
- Keep a job on disk, not only in memory: write `{job}.json` (code, output, value, error, changes, timings, client name) to `%APPDATA%\revit-ballet\runtime\agent\jobs\` with a retention (days/MB). This is the audit trail pyRevit has (14 days) and what you need to answer "what did the agent do to my model last night". Redact the token, never log headers.
- Idempotency/correlation: optional `clientId` and `requestId` on `POST /v1/jobs`; a repeated `requestId` returns the existing job instead of running twice (agents retry on timeouts).

### 8. Concurrency and lifecycle details worth tightening

- `AgentHost.current` and `stuck` are written on the worker thread and read from HTTP threads (`Info`) without a lock or `volatile`. Make them volatile or read under `stateLock`.
- A timed-out script keeps running on the UI thread (`stuck`); the next job waits in a 100 ms poll. Surface this explicitly: `/v1/info.queue.stuckJob` and a way to ask for escalation (`POST /v1/jobs/{id}/abort` that uses `Thread.Abort`-equivalents where the engine supports it, or at least flags the session as tainted).
- Session state hazards: a persistent session can hold `Element`/`ElementId`/`Document` objects that become invalid after a document closes or Revit regenerates. Document it, drop or flag session variables on `documentClosing`, and mention `session.reset` on `/v1/sessions`.
- Multiple documents: globals `Doc`/`UIDoc` are the *active* document at run time. Add an explicit per-job `document` selector (title or path) that sets `Doc` without switching the active document (we verified the API supports reading and writing non-active documents; note selection must still use `SelectionStorage`).
- Per-job `priority` (`interactive` jumps ahead of `bulk`) is optional but cheap once there is one queue, and stops a long batch from starving a quick query.
- Dispose/Reset engine state on add-in unload and hot-reload, otherwise hot-reloaded DLLs will leave old engines holding stale assemblies (relevant to our `--hot-reload` workflow).

### 9. Security hardening (cheap)

- Reject requests that carry a browser `Origin` header or whose `Host` is not `127.0.0.1:<port>`/`localhost:<port>` (blocks DNS-rebinding/CSRF from a web page). The token is in a header so risk is low, but the check costs a few lines.
- Accept the token only in headers, never in the query string (it ends up in logs).
- Record a client name (`X-Agent-Name`) in each job for the audit log.
- The protocol doc says the token is the boundary. Say explicitly that the `query` mode in section 2 is a mistake guard, not a permission system, so nobody mistakes it for one.
- Optional: a second, read-only-intended token that forces `mode: "query"` server-side. Same caveat, but useful to hand to less trusted agents.

### 10. Client side and test coverage

- A small first-party CLI/library for the host operator and agents (`rb jobs run --engine roslyn --file x.cs --session s1 --wait`, `rb events --follow`, `rb context`, `rb dialog respond`). Agents can then use plain Bash instead of hand-written curl (CLAUDE.md already warns about non-ASCII bodies corrupting `curl --data-binary`; a CLI removes that class of bug).
- Extend `tests/vm/guest/v1-tests.ps1` to cover: queue ordering with concurrent submits, cancel while compiling/queued/running, timeout then next job, modal dialog block and recovery via `/v1/dialog/respond`, session persistence across jobs, event stream resume with `since`, change report accuracy, output truncation, and the Revit 2025/2026 (.NET 8) path for each engine, in particular CPython/pythonnet and IronPython 2.7 loading under .NET 8.

## Suggested order of work

1. Wire up what exists (the protocol doc still says "not wired in"; the working tree has AgentHost/AgentHttp now - update the status line when it is true).
2. Section 4 (Idling drain + liveness fields) and section 7 (limits, persistence): they affect correctness of everything else.
3. Section 2 (change report + modes) and section 5 (`/v1/context`, API lookup, prelude): biggest gain in agent reliability per line of code.
4. Section 1 (event stream), then triggers.
5. Section 3 (dialog response), section 6 (view image), section 9, section 10.
