# Remote command invocation and the canonical query DSL

Status: implemented 2026-10-10 (everything below exists and is tested; "Open" at the end lists what is not done). The text keeps its design form.

Two connected changes:

1. **One query language.** The DataGrid search box in Revit and the pickers on the user's Linux desktop (`mod+`, `mod+v`,
   `mod+ctrl+n`, the launcher) must speak the same DSL. The canonical definition is `~/.config/docs/query-dsl.md` (dotfiles repo).
   revit-ballet adapts to it, not the other way round.
2. **Remote command invocation.** Today `Q` then `` ` `` opens a list of `IExternalCommand`s, but only inside one Revit session. The same list
   must be reachable from outside Revit, and must be able to run a command in *any* Revit instance: on the same Windows machine, in a VM,
   or on another computer. Shortcut: `mod+`` on Linux (Hyprland) and the nearest free equivalent on Windows.

## Part A - the DSL

### What exists

`core/Filtering/FilterQueryParser.cs` + `RowFilter.cs` (~800 lines, 65 xunit tests in `core/tests`), one consumer: `DataGrid.Filtering.cs`
(`ApplyFilters`, column visibility/ordering). Old syntax (documented in `readme.org`):

| old | meaning |
|---|---|
| `a b` | AND of substrings over all columns |
| `a \|\| b` | OR of groups |
| `!a` | exclusion |
| `"a b"` | quoted term |
| `e"x"` | exact cell match |
| `a*b` | glob |
| `>50` `<50` | numeric compare on any column |
| `$col:val` `$col::val` `$"col name":val` `$e"col":val` | column-scoped value filter (also shows that column) |
| `$col` | show only that column |
| `2$col` | show column at position 2 |
| `#set` `#"set name"` | element in a Revit selection set |

### Target

The canonical grammar, exactly (`query-dsl.md`): verbs `/fv /ft /at /rt /s /rv` and long forms, exact-matched; via paths `/fv/path value`,
`//path value`; type paths with `.` subsets and `*`; substring matching, case-insensitive, quoted values literal; `!` negation on row filters;
`/fv path:value` with `> < >= <=` comparison; `/sort/a/b/c [asc|desc]`, last `/sort` wins, `/rv` idempotent; sort comparator sniffs integers and
age buckets; half-typed input is inert, never an error.

### Mapping of old features

| old | canonical | notes |
|---|---|---|
| `a b` | `a b` | unchanged (bare words are `/fv` terms, ANDed) |
| `!a` | `!a` | unchanged |
| `"a b"` | `"a b"` | now a literal contiguous substring (was: same, via StripQuotes) |
| `$col:val` | `//col val` (also `/fv col:val`) | the old form's side effect "also show that column" becomes the *auto-shown filter field* rule of the DSL |
| `!$col:val` | `!//col val` | |
| `$col` | `/ft col` | |
| `2$col` | order of `/at` verbs | the DSL has no numeric positions; left to right is the order |
| `$col:>50` | `//col >50` | same comparator as `/sort` |
| `#set` | `//set name` | selection sets become a multi-valued field `set` (names of every set the element is in) |
| click a column header | `/s col` | one comparator for both (`compare_field_values`) |

### Features the canonical DSL does not have

These existed and have no spelling in `query-dsl.md`. Each needs a decision (recorded here once taken):

1. `a || b` (OR groups).
2. `e"exact"` exact cell match, and `$e"col"` exact column match.
3. `*` globs in values.
4. Bare `>50` over *all* columns.

Recommendation: keep the canonical grammar closed (it is shared by about ten pickers in three languages) and treat the old spellings as removed,
because each has a canonical neighbour (quoted substring, `//col >50`, `!`). If any of them is worth keeping, add it to `query-dsl.md` first
and implement it everywhere, not as a revit-ballet-only extension. **Open: the user decides.** Until then the engine implements the canonical
grammar only and the migration notes list what changed.

### Implementation

New namespace `RevitBallet.Core.Query` (no Revit types, runs and is tested on Linux):

- `QueryTokenizer` / `QueryParser` -> `QueryPlan { RowTerms, ColumnOps, Sort, Reverse }`.
- `QuerySchema`: flat fields, groups with a default subfield, a free-text haystack function. The host supplies one per picker.
- `PathResolver`: per-segment substring union, `*`, group/flat resolution.
- `QueryExecutor.Apply(rows, schema, plan)` -> `{ rows, columns }` (filter, then columns, then order).
- `FieldCompare`: integers, age buckets, lexicographic, direction trap.
- Completion (Part C) is a separate layer on the same tokenizer.

The old parser/matcher is removed in the same change that migrates `DataGrid`, so there is never a second engine.

## Part B - remote command invocation

### Pieces

1. **Command catalogue in the agent.** `GET /v1/commands` lists every `IExternalCommand` with `name`, `displayName`, `input`, `output`
   (the `CommandMeta`/`CommandOutput` attributes), `years`; `POST /v1/commands/{name}/invoke` runs it on the UI thread through the same
   ExternalEvent path as jobs (queued FIFO, timeouts, modal-dialog reporting, `/v1/jobs/{id}` for the result). Auth token as for the rest of `/v1`.
2. **Target discovery** (`revit-ballet-remote`, a small tool, one implementation per OS family, same behaviour):
   - local: Revit sessions are the listeners on 23717-23817 with `%APPDATA%\revit-ballet\runtime\network\token`;
   - remote: a targets file lists transports (ssh, ssh over vsock `ProxyCommand`, direct https). Each target yields its sessions the same way.
   Result: a flat list of *instances* `{target, year, pid, port, document, modal dialog, ui responsive}`.
3. **The picker.** One row per (instance x command)? No: instances are a column, commands are rows, and the instance is chosen by the DSL
   (`//instance vm2`, `//year 2026`) or, when only one instance matches, implicitly. Fields: `command`, `input`, `output`, `instance`, `target`,
   `year`, `document`, `state`. Enter invokes. Same query box, same verbs, same history, same auto-shown filter fields as the other pickers.
4. **Shortcuts.** Linux: `mod+`` in `hypr/keybinds.lua` -> `qs ipc call revitRemote toggle` (free, no existing binding). Windows: a hotkey
   helper (today `InvokeAddinCommandInNetwork` registers `Alt+Q`, `` ` ``); the installer offers it as an opt-in. Candidates in order: `Win+`` (to be
   tested in the VM with `RegisterHotKey`), `Alt+`` , the existing two-step sequence.
5. **Installer.** New optional component "Remote command picker" (default off): installs the helper, registers autostart, writes the hotkey
   choice. Dialog text states what it does and that it listens on loopback only.

### Security

The agent already requires the shared token and binds loopback. Remote reach goes through the transports above (ssh / vsock), never by opening
the port to the network. The picker only ever holds a token for a target the user listed.

### Testing

- Unit (Linux, xunit): the whole DSL, schema, comparator, plan execution, an exhaustive old-to-new migration table, property tests for
  "T and !T partition the rows".
- Agent: `/v1/commands` and invoke in `tests/host` against the Revit-free core.
- End to end in the offline VMs (`tests/vm`): list commands from outside Revit, invoke a harmless one, check the document changed,
  invoke while a modal dialog is up, two instances in two VMs addressed by the DSL; Windows hotkey registers and fires (`SendInput`).
- Linux picker: a QML test harness fed fake instances, plus one real run against the VMs.

## Part C - the Linux picker

`~/.config/quickshell/revitremote/` (new consumer): reuses `launcher/QueryDsl.qml` for parsing and completion, a backend script
(`revit-remote list|invoke`) for discovery and invocation, an IPC handler `revitRemote toggle`, a row in the consumer table of
`query-dsl.md`, a description on the Hyprland bind so it shows up in the `mod+?` cheat sheet.

## What exists (2026-10-10)

| Piece | Where | Tests |
|---|---|---|
| DSL engine (parser, schema, executor, comparator) | `core/Query/` | `core/tests` (Query*Tests, 167 tests incl. the old-syntax regression table) |
| DataGrid search box on the engine | `commands/DataGrid/DataGrid.Filtering.cs` | VM: `datagrid-dsl-in-the-real-dialog` drives the real dialog (15 queries) |
| `GET /v1/commands`, `POST /v1/commands/{name}/invoke`, built-in `command` engine | `commands/Agent/` (protocol: `docs/agent-protocol.md`) | `tests/host/CommandTests.cs` (12), VM `commands-*` checks |
| Linux backend `revit-remote` | `tools/remote/` | `tests/remote` (34, fake transport) + real runs against the VMs |
| Linux picker (`mod+``) | dotfiles `~/.config/quickshell/revitremote/`, bind in `hypr/keybinds.lua` | `qmltestrunner` model tests (20), headless load of the real window |
| Windows picker `RevitBalletRemote.exe` (tray, `Win+``, picker) | `remote/RevitBalletRemote/`, logic in `core/Remote/` | `core/tests/RemoteTests.cs`, `tests/vm/remote-picker.sh` (real hotkey and typing) |
| Installer opt-in (`--remote-picker`) | `installer/RemotePicker.cs` | `tests/vm/guest/remote-installer.ps1` via `remote-picker.sh` |

### Decisions taken

- **The canonical grammar stays closed.** `a || b`, `e"exact"`, `*glob*`, a bare `>50`, `N$col` and `#set` have no canonical spelling and were removed
  (`readme.org` lists the replacements; `QueryMigrationTests` pins what each old form returned). Selection sets became the computed column
  `Selection Set` (`//selection "my set"`). If any removed form should come back it has to be added to `query-dsl.md` and every picker, not to revit-ballet alone.
- **Deviations from the spec, on purpose:** comparisons (`>`, `<`, `>=`, `<=`) work in every consumer here (the spec lists them for two pickers) and
  accept decimals, because Revit areas and lengths are decimals; text sorts naturally (`A2` before `A10`) through the existing `NaturalComparer`;
  `QuerySchema.IgnoreSpacesInNames` lets `type name` reach the key `TypeName` in the DataGrid.
- **`/ft` intersects with the shown columns**, as both QML reference implementations do (the spec's `/ft claude /at workspace` example is looser than its rule).
- **Commands run through the job queue**, not a side channel: an invocation queues FIFO behind scripts, is cancellable, reports modal dialogs and has the
  maximum time limit by default because commands wait for people.
- **Nothing is installed on a target**: a PowerShell probe is piped over ssh (`tools/remote/probe.ps1`, `invoke.ps1`, embedded in `RevitBallet.Core` for the Windows program).
- **Windows hotkey `Win+``**: free on Windows 11 (`RegisterHotKey` succeeds; `Win+Q`, `Win+Z` do not), resolved with `VkKeyScan` so a UK layout
  (backtick is `VK_OEM_8`) works, re-resolved when the layout changes. Configurable (`"hotkey"` in `remote.json`). The opt-in is off by default.

### Open

- The Windows picker has no Tab completion or inline command colouring yet (the engine supports the queries; the Linux picker has the shared completion).
- There is no `document` column (the open documents are in `documents.sqlite`, not exposed by the agent; `/v1/info` would need a UI-thread cache).
- `.NET Framework` Revit years (2020-2024): the DataGrid check is skipped in the quick tier (it needs the installed build); the full CI covers it.
- Older agents (an installed build without `/v1/commands`) are listed as instances with no commands until they are updated and Revit is restarted.
- Remote reach beyond ssh (a direct https target) is not implemented.
