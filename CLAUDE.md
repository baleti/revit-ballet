# Revit Ballet - AI Agent Integration & Code Conventions

> **Documentation Size Limit**: This file must stay under 20KB. Keep changes concise.

This file provides guidance to AI agents (like Claude Code) when working with code in this repository and querying Revit sessions.

## Project Overview

Revit Ballet is a collection of custom commands for Revit. Two main components:

1. **Commands** (`/commands/`) - C# plugin with custom Revit commands
2. **Installer** (`/installer/`) - Windows Forms installer application

## Development Environment

Runtime data lives on the Windows production machine at `%APPDATA%\revit-ballet\` (not in this repo). Access it via SSH — credentials are stored in agent private memory (not committed here). Key paths on Windows:
- `network\token` — shared auth token
- `documents` — live session registry (CSV)
- `server.log` — server log
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

**Incident 2026-01-24**: agent replaced `.git` from a temp clone during a rebase workaround, destroying stash refs and losing a day of the user's work. Recovery was impossible. User's uncommitted work is more valuable than clean history.

### Case-Insensitive Filesystems (CIFS/drvfs)

Repo is accessed via case-insensitive Windows filesystems but git is case-sensitive. Key implications:

- Use consistent casing in filenames (PascalCase for C# files)
- Before rebasing, check for case-variant duplicates: `git ls-files | sort -f | uniq -di`
- Prefer WSL drvfs over Linux CIFS for rebases
- If `git reset --hard` / `checkout --` fail despite file existing (CIFS cache corruption), try `git update-index --skip-worktree <file>` temporarily
- Git stash may show error messages but still succeed — always verify with `git stash list`

### Diagnostic Logging

For non-trivial operations (unit conversion, coordinate transforms, API quirks, perf), write diagnostics to `%appdata%\revit-ballet\diagnostics\` using `PathHelper.RuntimeDirectory`. Name files `{Operation}_{yyyyMMdd_HHmmss}.txt`. Log timestamps, element IDs, internal units (feet) AND display units, intermediate steps. **NEVER** write diagnostics to temp or document folders.

## Version Control Practices

**One Agent Session = One Commit.** Accumulate all changes, commit once at the end with a descriptive message. Don't commit unrelated changes alongside your primary work — stage only files relevant to the session's goal. Verify with `git status` and `git diff --cached` before committing.

## Architecture

### Multi-Version Revit Support
Supports Revit 2017-2026 via conditional compilation on `RevitYear`:
- 2017-2018: .NET Framework 4.6
- 2019-2020: .NET Framework 4.7
- 2021-2024: .NET Framework 4.8
- 2025-2026: .NET 8.0 Windows

Revit.NET package versions are selected automatically based on target year.

### Roslyn Server for AI Agents

Revit Ballet runs a Roslyn compiler-as-a-service so AI agents can execute C# in Revit session context.

**Peer-to-Peer Network:**
- Each Revit instance runs its own HTTPS server, auto-started when Revit loads
- Port range 23717-23817 (first available)
- Shared 256-bit token across all sessions — first session generates it, others reuse
- File-based discovery via `%appdata%\revit-ballet\`:
  - `documents` — CSV: `DocumentTitle,DocumentPath,SessionId,Port,Hostname,ProcessId,RegisteredAt,LastHeartbeat,LastSync`
  - `network\token` — shared auth token

**Endpoints:**
- `/roslyn` — execute C# scripts in Revit context
- `/screenshot` — capture Revit window to `%appdata%\revit-ballet\screenshots\`

**Implementation:** `commands/RevitBallet.cs` initializes server; background thread marshals to UI via `ExternalEvent`; heartbeat every 30s, 2min dead-session timeout. For server-level hangs, check `server.log` on Windows via SSH.

## Querying and Testing Revit Sessions

The Roslyn server runs on Windows at `127.0.0.1:PORT`. Access from Linux requires SSH port forwarding (credentials in private agent memory).

**Canonical pattern — SSH tunnel + file-based query:**

```bash
# Get token and port from Windows via SSH
TOKEN=$(ssh daniel@192.168.231.1 -i ~/.ssh/hwo-18 \
  'powershell -NoProfile -Command "Get-Content $env:APPDATA\revit-ballet\network\token"')
PORT=$(ssh daniel@192.168.231.1 -i ~/.ssh/hwo-18 \
  'powershell -NoProfile -Command "(Get-Content $env:APPDATA\revit-ballet\documents | Where-Object { $_ -notmatch \"^#\" } | Select-Object -First 1).Split(\",\")[3]"')

# Tunnel the port
ssh -f -N -L ${PORT}:127.0.0.1:${PORT} daniel@192.168.231.1 -i ~/.ssh/hwo-18

# Write C# query to a file and run it
cat > /tmp/my-query.cs << 'EOF'
var levels = new FilteredElementCollector(Doc).OfClass(typeof(Level)).Cast<Level>();
Console.WriteLine("Total Levels: " + levels.Count());
EOF
curl -k -s -X POST https://127.0.0.1:${PORT}/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  --data-binary @/tmp/my-query.cs | jq
```

**Why files?** Avoids bash escaping issues with quotes/operators and allows multi-line C#. Use `-k` for self-signed TLS (localhost only).

**ASCII only in query files.** Multi-byte UTF-8 (✓, ❌, ⚠) can corrupt `curl --data-binary` transmission (Content-Length mismatch → hung request). Use `OK`, `FAIL`, `WARNING` instead.

**Bash tool quirk:** Claude Code's Bash tool can fail on multiple `$(...)` in one call. Workaround: put the script in a file and `bash` it, or use only one substitution per call.

**Tail server log:**
```bash
ssh daniel@192.168.231.1 -i ~/.ssh/hwo-18 \
  'powershell -NoProfile -Command "Get-Content $env:APPDATA\revit-ballet\server.log -Tail 50"'
```

## Available Context

Scripts have globals `UIApp`, `UIDoc`, `Doc`. Pre-imported: `System`, `System.Linq`, `System.Collections.Generic`, `Autodesk.Revit.DB`, `Autodesk.Revit.UI`.

**Response format:** `{ Success, Output, Error, Diagnostics, ProcessingLog }`. `ProcessingLog` is a timestamped trace of compile/execute stages — check it for post-mortem when a query fails or hangs. For server-level hangs, check `server.log` via SSH.

**Script timeout:** 30 seconds. Timeouts return `Success: false` with an explanatory `Error` and the timeout marker in `ProcessingLog`.

### Screenshot Capture

```bash
# Tunnel port first, then:
curl -k -X POST https://127.0.0.1:${PORT}/screenshot \
  -H "X-Auth-Token: $TOKEN" | jq
```

Captures the full Revit window to `%appdata%\revit-ballet\runtime\screenshots\{timestamp}-{sessionId}.png`. Auto-cleanup keeps last 20. Returns absolute path in `Output`.

## Error Handling

Compile errors return `Success: false` with `Diagnostics` (e.g. `(1,23): error CS1002: ; expected`). Runtime errors return `Success: false` with the exception string in `Error` plus any output written before the throw. Server continues listening after errors.

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
- 256-bit shared token at `runtime/network/token` (plain text)
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

## Testing Across Revit Versions

```bash
cd ~/revit-ballet/commands && for i in {2026..2017}; do dotnet build -c Release -p:RevitYear=$i; done
```

Recommended when modifying core functionality, Revit API interactions, or doing significant refactors.

## Key Files

- `commands/commands.csproj` — main plugin project
- `commands/RevitBallet.cs` — extension application entry point (initializes server)
- `commands/Server.cs` — auto-started Roslyn server
- `commands/PathHelper.cs` — runtime path management

## Command Registration Policy

When creating new commands, **do not** automatically register them. Never modify without explicit request:
- `installer/revit-ballet.addin` (command manifest)
- `KeyboardShortcuts-custom.xml` (shortcut assignments)

Create the `.cs` in `commands/` and tell the user it's ready but unregistered.

## Command Design Principles

**Silent Completion.** Commands finish without success dialogs. No "Operation Complete" / summary popups. Dialogs are only for errors (failure the user needs to know about), required input (decision needed), or warnings (important conditions).

## Worktree Branch Naming (Agent Sessions)

When a session starts in a worktree with an auto-generated branch name matching `claude\d?-\d+` (e.g. `claude3-152818212`), **immediately rename** to something descriptive:
```bash
git branch -m claude3-152818212 fix-installer-partial-build
```
Do this first so the PR and merge commit have meaningful names.

## Worktree Cleanup (Agent Sessions)

**Problem**: agents run with Bash CWD set to the worktree. Once it's deleted, all subsequent Bash commands fail with "Path does not exist" — even with `cd /other/path &&` prefixes.

**Solution**: merge, remove worktree, delete branch in a **single Bash call**:
```bash
cd /home/user/revit-ballet && \
  git merge BRANCH_NAME --no-ff -m "Merge branch 'BRANCH_NAME': description" && \
  git worktree remove .worktrees/BRANCH_NAME && \
  git branch -d BRANCH_NAME
```
After this completes the agent's bash tool will error, but the work is done.

**If `ExitWorktree` is available** (only when `EnterWorktree` was used in the same session): `ExitWorktree(action: "remove")` cleans up before the directory is touched.

**If bash breaks after cleanup**: do NOT assume it failed. The merge/remove/delete very likely all succeeded. Tell the user to check `git log --oneline --decorate` and `git branch -v` before asking for manual intervention.

## File Naming Conventions

- Command files: PascalCase matching class name (`DataGrid2EditMode.cs`)
- One command per file, complete implementation
- Internal/utility files: PascalCase (`Server.cs`, `PathHelper.cs`)
