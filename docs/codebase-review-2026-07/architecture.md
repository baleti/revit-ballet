# Architecture Findings

## What is already good

Worth stating explicitly, because these should be preserved through any refactor:

- **Single parameterized csproj** (`commands/commands.csproj`) mapping `RevitYear` to
  TFM, API package version, and `REVIT{year}` define. This is the modern recommended
  pattern for multi-version Revit addins; many commercial addins do this worse.
- **`Build.sh` anchored to its own location** — worktree-safe, with per-year
  pass/fail summary.
- **Threading model of the server is correct**: background `TcpListener` loop, all Revit
  API access marshalled through `ExternalEvent` handlers, 30s script timeout,
  heartbeat/dead-session cleanup. This is the hard part of a Revit-embedded server and
  it was done right.
- **Git hygiene**: no `bin`/`obj` tracked, 4.7 MB pack, meaningful commit messages.
- **CLAUDE.md / AGENTS.md** (symlinked, not duplicated) is a genuinely good conventions
  document — scope-suffix naming, silent-completion policy, CommandMeta attribute,
  runtime-path discipline.
- **DataGrid already virtual-mode** with a prebuilt search index and change-detecting
  column-visibility updates — someone did real optimization work here.

## Finding 1: Namespace chaos (25 namespaces, many from tutorials)

Distribution across `commands/`:

```
16 RevitCommands            5 YourNamespace              1 YourCompany.YourAddin
16 RevitBallet.Commands     5 FilterDoorsWithWallOffsets 1 YourCompany.YourAddIn
12 RevitAddin               4 MyCompany.RevitCommands    1 TransactionMonitor... etc.
 8 MyRevitCommands          2 YourAddinNamespace         (+ ~180 files in the GLOBAL namespace)
```

`YourNamespace   // ← adjust` appears verbatim — the adjust comment included. These are
fossilized copy-paste origins (tutorials, forum posts, earlier AI models). Consequences:

- **Scripting-server fragility**: Roslyn scripts and any reflection-based dispatch
  (`InvokeAddinCommand`, the stack-walking helpers) must guess where a type lives.
  `SelectionModeManager.GetCallingCommandName` scans *entire assemblies* partly because
  there is no namespace convention to filter on.
- **Global-namespace pollution**: `CustomGUIs`, `ElementDataHelper`, `SelectionModeManager`,
  and most command classes sit in the global namespace. Revit loads all addins into one
  process (one AppDomain on .NET Framework); type identity is assembly-qualified so this
  won't hard-collide, but it makes IntelliSense, scripting, and grep noisier than needed.
- `commands.csproj` sets `<RootNamespace>scripts</RootNamespace>` — a third convention
  nobody follows.

**Recommendation**: one mechanical pass to `RevitBallet.Commands` (commands),
`RevitBallet.UI` (DataGrid/dialogs), `RevitBallet.Infrastructure` (server, storage,
paths). This is scriptable (each file has exactly one namespace declaration or none) and
should be a single commit with no logic changes. Update `RootNamespace` to `RevitBallet`.

## Finding 2: `CustomGUIs` — a static god object

The DataGrid UI is `public partial class CustomGUIs` spread across 10+ files
(~8,800 lines in `commands/DataGrid/` plus `commands/DataGrid.cs`). All state is
`static`:

```
DataGrid.EditMode.cs:13       private static bool _isEditMode
DataGrid.Filtering.cs:14-27   _cachedOriginalData, _cachedFilteredData, _currentGrid,
                              _searchIndexByColumn, _searchIndexAllColumns, ...
DataGrid.Main.cs:15-29        _currentFontSize, _currentScreenState, _initialSizingDone, ...
DataGrid.ColumnHandlers.cs    _currentUIDoc, _hasRevitApiAccess, _revitWindowHandle
```

Problems:

- **Not reentrant.** Two grids can never be open at once (nested pickers, or a future
  modeless grid) — the second would corrupt the first's cached/filtered data.
- **State leaks between invocations.** `_initialSizingDone`, `_pendingCellEdits`,
  `_selectionAnchor` etc. rely on every code path resetting them; a missed reset in one
  of the 282 callers shows up as a bug in an unrelated command.
- **Untestable.** Filtering/sorting logic is welded to `DataGridView` and static fields.
- `CustomGUIs.SetCurrentUIDocument(uidoc)` is required *ambient setup* — the classic
  temporal-coupling smell; forgetting it silently disables editing.

**Recommendation**: introduce an instance class (e.g. `DataGridSession`) holding all
per-invocation state; keep the existing static `CustomGUIs.DataGrid(...)` signature as a
thin facade constructing one session per call, so no caller changes. This is also the
prerequisite for both unit testing and any UI-framework experiment (see datagrid-ui.md).

## Finding 3: Call-stack reflection to discover "who called me"

Two independent implementations of the same fragile idea:

- `DataGrid/DataGrid.Main.cs:41` `InferCommandNameFromCallStack()` — walks
  `StackTrace`, pattern-matches class names, converts to kebab-case (for search history
  keying).
- `SelectionModeManager.cs:33` `GetCallingCommandName()` — walks the stack, then scans
  **all types in all assemblies on the stack** looking for concrete `IExternalCommand`
  implementations, with inheritance-depth heuristics.

These are slow (assembly-wide `GetTypes()` on every selection call), break under
inlining/Release optimization, and duplicate information the codebase already models:
every command carries `[CommandMeta]` and a class name. **Recommendation**: pass the
command name (or `Type`) explicitly as a parameter / ambient context object set at
command entry. One `ExecutionContext.CurrentCommand` set in a shared command base class
would replace both.

## Finding 4: File organization

- `commands/` is a flat directory of 282 files. The scope-suffix convention makes names
  self-describing, but discovery is hard. Subfolders by verb domain (Selection/, Views/,
  Sheets/, Revisions/, Export/, Geometry/, Network/, Infrastructure/) would cost nothing
  at build time (SDK-style csproj globs everything).
- **Two multi-file conventions coexist**: underscore-suffix file families
  (`CopyElementAlongContainingGroupByRoom_*.cs` — 8 files,
  `FilterDoorsWithWallOffsets_*.cs` — 5 files) versus a real folder (`DataGrid/`).
  Pick folders; the underscore families are folders in denial.
- **Name collision**: `commands/DataGrid.cs` (contains `ElementDataHelper`,
  `ListElementsBase`, and a command class named `DataGrid`) vs. `commands/DataGrid/`
  (the `CustomGUIs` grid UI) vs. the method `CustomGUIs.DataGrid(...)`. Three different
  things named DataGrid. Rename the command class (it is the "list elements" command)
  and move `ElementDataHelper` to its own file.
- **Dead/odd files**: `DataGrid/ColumnHandlers.Framework.cs.txt` is a design proposal
  stored as `.txt` inside the source tree — move to `docs/`;
  `installer/_installer-*.py` are Python helpers with a `_` prefix convention used
  nowhere else.

## Finding 5: Always-true conditional compilation

`RevitBallet.cs`, `Server.cs` (+2 more) open with:

```csharp
#if REVIT2011 || REVIT2012 || ... || REVIT2025 || REVIT2026
```

Every build defines exactly one `REVIT{year}`, so this is always true — it's a no-op
wrapper that must be hand-extended every year (note it already lists 2011–2016, years the
project doesn't target). Delete these guards; keep `#if` only where API surfaces actually
differ (the pattern used correctly elsewhere, e.g. SQLite/Roslyn package splits).

## Finding 6: No CI

There is no `.github/` directory. Every per-year compile break is discovered manually via
`Build.sh`. Because the Revit API is consumed from NuGet
(`Revit_All_Main_Versions_API_x64`) and `EnableWindowsTargeting` is already set, **a
plain `ubuntu-latest` GitHub Actions job can compile all ten Revit years** — no Windows
runner, no Revit license, no VM:

```yaml
strategy:
  matrix:
    year: [2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026]
steps:
  - uses: actions/checkout@v4
  - uses: actions/setup-dotnet@v4
  - run: dotnet build commands -c Release -p:RevitYear=${{ matrix.year }}
```

This is the single highest value-per-hour improvement available. Once
`RevitBallet.Core` exists (testing.md), the same workflow runs the unit tests.

## Finding 7: The stalled standardization plan

`docs/command-standardization-plan.org` (April 2026, still DRAFT) already prescribes the
right things: no "Selected" in names, singular command forms, shared `InputResolver`.
`InputResolver.cs` was built, but only **8 of 282 files** reference it, while the
selection-or-fallback dance is hand-rolled in dozens of others (e.g.
`SelectByCategoriesInView.cs:18-57` re-implements `ResolveViews` inline). The recent
commits ("universal selection-first + picker-fallback", "extend to 11 more commands")
show migration is underway — the plan is good; it just needs to be finished and then
*enforced* (a CI grep that fails when a new command hand-rolls view resolution is crude
but effective).
