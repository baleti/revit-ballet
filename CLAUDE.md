# Revit Ballet - Code Conventions & AI Integration

> **Documentation Size Limit**: This file must stay under 20KB. Keep changes concise.

This file provides guidance when working with code in this repository and querying Revit sessions.

## Project Overview

Revit Ballet is a collection of custom commands for Revit. Two main components:

1. **Commands** (`/commands/`) - C# plugin with custom Revit commands
2. **Installer** (`/installer/`) - Windows Forms installer application

## Development Environment

Runtime data lives on the production Windows machine at `%APPDATA%\revit-ballet\`. Key paths:
- `runtime\network\token` — shared auth token
- `runtime\documents` — live session registry (CSV)
- `runtime\server.log` — server log
- `screenshots\` — captured screenshots
- `diagnostics\` — diagnostic files

**ALL runtime data MUST be stored in `%APPDATA%\revit-ballet\`** — logs, network data, screenshots, diagnostics.

## ⚠️ Git Operations - Protect Uncommitted Work

**NEVER**:
- Replace/overwrite the `.git` directory (even from a clone or backup) — destroys stash refs permanently
- Run destructive operations (`git checkout --`, `reset --hard`, etc.) without first running `git status` and `git stash list`
- Stash/discard uncommitted work without explicit user permission

**Before any destructive git operation:** check `git status` and `git stash list`. If there's uncommitted work, ASK first — offer to commit, stash (risky), or cancel.

**If rebase fails:** `git rebase --abort` and ask the user how to proceed. Do NOT try to "fix" it by manipulating `.git` directly.

### Diagnostic Logging

For non-trivial operations (unit conversion, coordinate transforms, API quirks, perf), write diagnostics to `%appdata%\revit-ballet\diagnostics\` using `PathHelper.RuntimeDirectory`. Name files `{Operation}_{yyyyMMdd_HHmmss}.txt`. Log timestamps, element IDs, internal units (feet) AND display units, intermediate steps. **NEVER** write diagnostics to temp or document folders.

## Architecture

### Multi-Version Revit Support
Supports Revit 2020-2026 via conditional compilation on `RevitYear`:
- 2020: .NET Framework 4.7.2
- 2021-2024: .NET Framework 4.8
- 2025-2026: .NET 8.0 Windows

Revit.NET package versions are selected automatically based on target year.

### Agent Server (scripts and AI agents run inside Revit)

Each Revit instance hosts a local HTTPS server so scripts and AI agents can run code in the live session. The host (`commands/Agent/`) has no scripting engine; engines are separate components under `agents/` (installed to `engines/<name>/` next to `revit-ballet.dll`, loaded on first use): `roslyn` (C#), `ironpython2` (IronPython 2.7), `cpython3` (CPython 3.12 via pythonnet, bundled runtime). Protocol: `docs/agent-protocol.md`.

**Peer-to-Peer Network:**
- Port range 23717-23817 (first available); shared 256-bit token across all sessions (`runtime\network\token`, header `X-Auth-Token`)
- Discovery: `runtime\documents` CSV (`DocumentTitle,DocumentPath,SessionId,Port,Hostname,ProcessId,RegisteredAt,LastHeartbeat,LastSync`)

**Endpoints:** `GET /v1/info` (Revit year, engines, `ui.responsive`, `ui.modalDialog`), `POST /v1/jobs` (`{engine, code, timeoutMs, session, wait}`), `GET /v1/jobs/{id}?wait=ms`, `GET /v1/jobs/{id}/events` (SSE), `DELETE /v1/jobs/{id}` (cancel), `GET /v1/sessions`, `DELETE /v1/sessions/{engine}/{name}`; legacy `POST /roslyn` (raw C# body, old `{Success,Output,Error,Diagnostics,ProcessingLog}` response), `/screenshot`, `/query/*`.

**Implementation:** jobs queue FIFO; compile on a worker thread, run on the UI thread via `ExternalEvent` (`commands/Agent/AgentHost.cs`). Engine contract: `agents/abstractions`. Heartbeat 30s, 2min dead-session timeout.

## Querying Revit Sessions

The server runs on Windows at `127.0.0.1:PORT`. Access from a remote dev machine requires port forwarding to reach it.

**Bodies are UTF-8 JSON; the legacy `/roslyn` body is raw C#.** (The old "ASCII only" limit is gone: Content-Length is now counted in bytes.)

### Available Context

Globals: C# `UIApp`, `UIDoc`, `Doc`, `Out`, `Cancel`; Python `uiapp`, `uidoc`, `doc`, `out`, `cancel` (aliases `UIApp`, `UIDoc`, `Doc`). Python scripts return a value by setting `result`. C# pre-imports `System`, `System.IO`, `System.Linq`, `System.Collections.Generic`, `Autodesk.Revit.DB`, `Autodesk.Revit.UI`, `RevitBallet.Commands`. Python can `from Autodesk.Revit.DB import *` directly.

**Job result:** `{state, ok, output, value:{type,json,text}, error:{kind,message,stack}, diagnostics[{severity,id,line,column,message}], timings}`. Error kinds: `compile`, `runtime`, `timeout`, `blocked_by_dialog`, `cancelled`, `engine`, `internal`.

**Timeouts:** default 30 s of run time (`timeoutMs`, max 600000). A script stuck in a modal dialog reports `blocked_by_dialog` with the dialog title. Scripts can be stopped cooperatively through `Cancel` / `cancel.is_set()`; Python engines also stop hot loops.

## Error Handling

Compile errors give `state: failed`, `error.kind: compile` and positioned `diagnostics`. Runtime errors give `error.kind: runtime` with the message, a stack/traceback and any output written before the throw. The server keeps listening after errors. The legacy `/roslyn` shim reports the same in the old shape.

## Revit .NET API Best Practices

### Transactions
Always wrap document modifications:
```csharp
using (var trans = new Transaction(Doc, "Description"))
{
    trans.Start();
    // modifications
    trans.Commit();
}
```
Required for creating/modifying/deleting elements, changing parameters, adding to symbol tables.

### ExternalEvent for Background Threads
Operations initiated off the UI thread (like the server) must marshal via `ExternalEvent`:
```csharp
private ExternalEvent myEvent;
private MyHandler myHandler;

public void Initialize()
{
    myHandler = new MyHandler();
    myEvent = ExternalEvent.Create(myHandler);
}
myEvent.Raise();  // from background
```

### FilteredElementCollector
```csharp
// By category
new FilteredElementCollector(Doc).OfCategory(BuiltInCategory.OST_Walls)
    .WhereElementIsNotElementType().ToElements();

// By class
new FilteredElementCollector(Doc).OfClass(typeof(Level)).Cast<Level>();
```

### Output and Nulls
- Use `Console.WriteLine` for all output; format consistently (`ELEMENT|{id}|{name}|{category}`) for parsing
- Scripts may `return` a value — it's appended to output
- Always null-check (`element.Category?.Name`, `Doc == null` guard at top)

### Unit Conversion
**Revit stores everything internally in imperial (feet, sq ft, cu ft)** regardless of project settings.

- Parameters: `SetValueString()` auto-converts; `Set()` assumes feet
- Coordinates: use `UnitUtils.ConvertFromInternalUnits` / `ConvertToInternalUnits` with `Doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()`

## Security Considerations

- HTTPS with TLS 1.2/1.3, self-signed certs
- Localhost-only binding (127.0.0.1) — not network-accessible
- 256-bit shared token at `network\token` (plain text)
- Scripts have **full Revit API access** — can read, modify, delete model data
- Intended for local dev/automation only. Do NOT expose to untrusted networks, share the token, or run Revit elevated while the server is up.

## Naming Conventions

Revit Ballet uses **PascalCase** throughout (standard C#). This codebase was partially ported from AutoCAD Ballet (kebab-case) and fully migrated.

- Classes/Methods/Properties/Files/Directories: PascalCase
- Variables: camelCase
- The product name `revit-ballet` stays kebab-case (brand, not code)

Example migrations: `searchbox-queries` → `SearchboxQueries`, `server-cert.pfx` → `ServerCert.pfx`.

## Command Scopes (View, Document, Session, Network)

Commands use suffix naming to indicate scope:

1. **`InViews`** — current active view (e.g. `FilterSelectedInViews`)
2. **`InDocument`** — current active document (e.g. `FilterSelectedInDocument`). Previously `InProject`, renamed for consistency with AutoCAD Ballet.
3. **`InSession`** — all open documents in this Revit process (e.g. `SelectByCategoriesInSession`). Uses SelectionStorage for cross-doc selections.
4. **`InNetwork`** — all Revit sessions in the peer-to-peer network (to be implemented).

### Multi-Document (Session Scope) Operations

**Verified**: the Revit API fully supports reading AND writing across multiple open documents without switching the active document. `UIApp.Application.Documents` gives access to all; `FilteredElementCollector(doc)` and `Transaction(doc, "...")` work on any. Skip `doc.IsLinked == true` for modifications.

### Cross-Document Selection Storage

**Critical limitation**: `UIDocument.Selection.SetElementIds()` silently fails for element IDs from non-active documents — the call succeeds but selection stays empty. Revit's selection system is tied to the active UIDocument only.

**Solution**: SQLite-based `SelectionStorage` at `%appdata%\revit-ballet\selection.sqlite`. Stores `UniqueId` (stable across sessions), `DocumentTitle`, `DocumentPath`, and optional `SessionId` (for Network scope).

**API** (`RevitBallet.Commands.SelectionStorage`):
- `SaveSelection(items)` — replace
- `AddToSelection(items)` — append
- `LoadSelection()` / `LoadSelectionForOpenDocuments(app)`
- `ClearSelection()`

Retrieve elements via `doc.GetElement(item.UniqueId)` after grouping by `DocumentTitle`. See `commands/SelectionStorage.cs` and `commands/SelectByCategoriesInSession.cs` for the reference implementation.

**Why not SelectionModeManager?** That's for single-document workflows with view-switching and linked references. SelectionStorage is for multi-document (UniqueId-based, survives close/reopen).

## Selection Mode Manager

**CRITICAL**: All commands working with element selection MUST use `SelectionModeManager` extension methods, not direct Revit UI selection APIs.

Two modes: `RevitUI` (standard blue highlight) and `SelectionSet` (stored in a persistent `SelectionFilterElement` named "temp", enabling persistent cross-view selection and linked-model references).

**Always use:**
```csharp
uidoc.GetSelectionIds();
uidoc.SetSelectionIds(elementIds);
uidoc.AddToSelection(elementIds);
uidoc.GetReferences();           // including linked elements
uidoc.SetReferences(references);
```

**Never use** `uidoc.Selection.GetElementIds()` / `SetElementIds()` directly — bypasses SelectionModeManager.

Implementation: `commands/SelectionModeManager.cs`. Mode persisted at `%appdata%\revit-ballet\SelectionMode`; linked refs at `%appdata%\revit-ballet\SelectionSet-LinkedModelReferences`.

## DataGrid Automatic Editable Columns

DataGrid columns auto-enable editing based on column **name** via a handler registry. Commands just create dictionaries — edit detection, validation, and apply happen automatically.

**Usage:**
```csharp
CustomGUIs.SetCurrentUIDocument(uidoc);   // required once
var selected = CustomGUIs.DataGrid(gridData, propertyNames, false);
```

**Pre-registered editable columns** (case-insensitive): `Family`, `Type Name`, `Comments`, `Mark`.

**Requirements**: column name matches a registered handler; `SetCurrentUIDocument` called; row has `ElementIdObject` or `Id` (auto-injected if missing).

**Registering custom handlers** (in `RevitBallet.cs` OnStartup): build a `CustomGUIs.ColumnHandler` with `ColumnName`, `Getter`, `Setter`, optional `Validator`, and call `ColumnHandlerRegistry.Register(...)`.

Full details: `commands/DataGrid/IMPLEMENTATION_SUMMARY.md`. Code across `DataGrid.ColumnHandlers.cs`, `.Validation.cs`, `.EditMode.cs`, `.EditApply.cs`, `.Main.cs`.

## Building and Deploying

**Always use `Build.sh`** — it anchors paths to its own location so it works correctly from git worktrees as well as the main tree. Never `cd ~/revit-ballet/commands` directly; that always targets the main tree and will silently build stale code when working in a worktree.

```bash
# Build all Revit versions + installer (run from repo root or any worktree)
bash Build.sh

# Deploy to Windows with hot-reload (commands available immediately via InvokeAddinCommand)
scp -i ~/.ssh/hwo-18 installer/bin/Release/installer.exe daniel@192.168.231.1:installer.exe
ssh daniel@192.168.231.1 -i ~/.ssh/hwo-18 'powershell -NoProfile -Command "& .\installer.exe --quiet --hot-reload"'
```

### Hot-reload vs keyboard shortcuts — critical distinction

`--hot-reload` copies the DLL to `%APPDATA%\revit-ballet\hot-reload\{year}\`. **Only `InvokeAddinCommand` checks that folder** (it sets up an `AssemblyResolve` hook scoped to the invocation). Keyboard shortcuts bypass hot-reload entirely — Revit invokes them using the DLL it loaded at startup.

**When Revit is running and the DLL is locked**, the installer falls back to an update-folder strategy: it copies the new DLL to `Addins\revit-ballet.update\` and rewrites the `.addin` manifest to point there. The next Revit startup picks this up automatically — **you do not need to close Revit before running the installer**.

**What you DO need:** a full Revit restart (close the application, reopen it) after any deploy that renames or adds command classes. Until then, the running session still has the old DLL in memory and keyboard shortcuts will show "Wrong Full Class Name". `InvokeAddinCommand` works immediately via hot-reload without restart.

Summary:
- `InvokeAddinCommand` / `Q\`` — works immediately after deploy (hot-reload)
- Keyboard shortcuts — require Revit restart after deploy

## VM Testing / CI

End-to-end tests run against an offline Windows+Revit VM: `tests/vm/ci.sh` (overlay of a read-only baseline image, installs the built installer, runs engine and `/v1` checks per Revit year). Supported Revit: 2020-2026. Full description: `docs/vm-testing.md`.

## Key Files

- `commands/commands.csproj` — main plugin project
- `commands/RevitBallet.cs` — extension application entry point (initializes server)
- `commands/Server.cs` — auto-started HTTPS server (routing, auth); `commands/Agent/` — job queue and `/v1` API; `agents/` — script engines
- `commands/PathHelper.cs` — runtime path management

## Command Registration Policy

When creating new commands, **do not** automatically register them. Never modify without explicit request:
- `installer/revit-ballet.addin` (command manifest)
- `KeyboardShortcuts-custom.xml` (shortcut assignments)

Create the `.cs` in `commands/` and tell the user it's ready but unregistered.

## Keyboard Shortcuts and Addin Manifest — Keeping Them in Sync

**Three files must stay consistent:** `revit-ballet.addin`, `KeyboardShortcuts-custom.xml`, and the C# class name. When any one changes, update all three.

### Addin manifest (`revit-ballet.addin`)
Each command needs an entry with a stable `AddInId` (GUID) and a `FullClassName` that exactly matches the C# class name. If you rename a class, update `FullClassName` immediately — Revit will fail to initialize the add-in on startup if the class cannot be found, and the shortcut will silently not work.

Remove manifest entries for classes that no longer exist.

### Keyboard shortcuts (`KeyboardShortcuts-custom.xml`)

Two kinds of shortcut entries:

**Addin commands** (registered via `.addin` file — most revit-ballet commands, e.g. DD, DS):
```xml
<ShortcutItem CommandId="2a7e4f9b-3d6c-4e8a-7f2e-9c4d1a8b3e6f" Shortcuts="DD"/>
```
`CommandId` = the `AddInId` from the manifest (lowercase). No `CommandName`/`Paths` needed.

**External-tool commands** (registered via Revit ribbon, e.g. SMV/SMD/SMS):
```xml
<ShortcutItem CommandName="External &#xA;Tools:SelectByMaterialInView" CommandId="<guid>" Shortcuts="SMV" Paths="Add-Ins&gt;External"/>
```
Both `CommandName` (`"External &#xA;Tools:<ClassName>"`) and `Paths` (`"Add-Ins&gt;External"`) are required — Revit silently ignores bare entries for this type.

`CommandId` GUIDs must be **lowercase**. Each `CommandId` must appear **once** — duplicates cause a Revit import error.

Never edit Revit's `KeyboardShortcuts.xml` directly — Revit rewrites it on startup. Always update `KeyboardShortcuts-custom.xml` and run the installer.

## Command Metadata

Every `IExternalCommand` class should carry a `[CommandMeta("...")]` attribute (defined in `CommandMetaAttribute.cs`, global namespace — no `using` needed). The `Input` string describes what the user must have ready before invoking the command and is shown as the **Input** column in `InvokeAddinCommand`'s DataGrid.

Common values:
- `""` — no prior selection needed; command opens its own picker/dialog
- `"Any"` — one or more elements must be selected (any type)
- `"Door"`, `"Wall"`, `"View"`, `"Grid"`, etc. — specific element type required
- `"View, Link"` — context-aware command: list the optional selection types comma-separated when the command behaves differently based on what's selected (e.g. no selection → global, views selected → per-view, links selected → on linked model)

New commands must include the attribute. When the required input is not yet known, use `""`.

## Command Design Principles

**Silent Completion.** Commands finish without success dialogs. No "Operation Complete" / summary popups. Dialogs are only for errors (failure the user needs to know about), required input (decision needed), or warnings (important conditions).

## File Naming Conventions

- Command files: PascalCase matching class name (`DataGrid2EditMode.cs`)
- One command per file, complete implementation
- Internal/utility files: PascalCase (`Server.cs`, `PathHelper.cs`)
