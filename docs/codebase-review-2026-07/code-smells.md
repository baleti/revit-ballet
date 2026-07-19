# Code Smells Catalog

Concrete antipatterns found, with representative locations. Counts are from
`commands/` excluding `bin`/`obj`.

## 1. Silent exception swallowing (168 bare `catch {` + 26 `catch (Exception)`)

The single most consequential smell. Examples:

- `RevitBallet.cs` `OnStartup` — five consecutive try/catch blocks, four of which are
  `catch { /* Silently fail - don't interrupt Revit startup */ }`. Not interrupting
  startup is correct; discarding the failure is not. If the server fails to start or the
  column-handler registry fails to register, the user discovers it minutes later as a
  mysteriously broken feature with no trail.
- `Server.cs:44` — `catch { // Silently fail - server may already be running }` conflates
  "already running" with *every possible failure* (port exhaustion, cert generation
  failure, filesystem permissions).
- `DataGrid.cs:41` (ElementDataHelper) — `catch { /* Skip problematic links */ }` around
  linked-document scope-box collection: reasonable intent, but a corrupt link degrades
  silently on every invocation.

**Recommendation**: one tiny static helper writing to the diagnostics directory the
project already standardizes on:

```csharp
public static class Log
{
    public static void Warn(string context, Exception ex) =>
        File.AppendAllText(PathHelper.GetRuntimeFilePath("server.log"), // or addin.log
            $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{context}] {ex}\n");
}
```

Then a mechanical pass: every `catch { }` becomes `catch (Exception ex) { Log.Warn("...", ex); }`
unless the swallow is provably intentional (add a comment saying why). The startup blocks
in `RevitBallet.cs` should additionally set a status flag a diagnostics command can read.

## 2. Certificate validation disabled, copy-pasted 10×

```
CopyTypeParametersInNetwork.cs:1059,1087,1229,1249
SelectByFamilyTypesInNetwork.cs:483,711
SelectByCategoriesInNetwork.cs:337,547
SelectByWorksetsInNetwork.cs:288,510
```

All are `ServerCertificateCustomValidationCallback = (...) => true`. Justified for
self-signed localhost certs, but it should exist in exactly one place
(a shared `NetworkClient` — see duplication.md) where the justification is documented
and where certificate pinning against the known self-signed cert could later be added
in one edit. There are also 18 `new HttpClient(...)` sites — per-call client creation
inside `using` blocks. On .NET Framework this risks socket exhaustion under repeated
use (TIME_WAIT accumulation); a single static client per process is the standard fix
and again falls out of a shared `NetworkClient`.

## 3. Giant methods and giant files

- `CustomGUIs.DataGrid(...)` in `DataGrid/DataGrid.Main.cs:155` is a ~750-line method
  containing form construction, ~20 nested lambdas/closures for event handling, sizing
  logic, and key handling. Closures over ~15 mutable locals make this effectively
  untestable and risky to modify (any edit can capture the wrong variable lifetime).
- `DataGrid/DataGrid.EditApply.cs` — 2,223 lines; `Server.cs` — 2,042 lines (server +
  three ExternalEvent handler classes + response types in one file);
  `DataGrid/DataGrid.ColumnHandlers.cs` — 1,756 lines.
- `DataGrid.cs` mixes three responsibilities in 1,090 lines: element data extraction
  (`ElementDataHelper`), an abstract command base (`ListElementsBase`), and a concrete
  command (`DataGrid`).

**Recommendation**: don't refactor these for aesthetics alone — do it when extracting
`RevitBallet.Core` (testing.md), which forces the natural seams: pure logic out of the
event handlers, one class per file.

## 4. Success dialogs vs. the Silent Completion policy

507 `TaskDialog.Show` calls across 183 files. CLAUDE.md mandates dialogs only for
errors, required input, or warnings. Spot-checking suggests many are error dialogs
(fine), but the volume implies completion/summary popups survive in older commands.
Worth a scripted audit: grep for `TaskDialog.Show` whose message contains
`"Success"`, `"Complete"`, `"Copied"`, counts of processed elements, etc., and convert
those to silent completion (or a status-bar/log write).

## 5. Stringly-typed cross-file contracts

- Column behavior keyed by display-string column names (`"Type Name"`, `"Comments"`),
  case-insensitively, across DataGrid, handlers, and commands. A renamed header silently
  detaches its edit handler. Mitigation: `public static class ColumnNames { public const string TypeName = "Type Name"; ... }`.
- The `documents` session registry is hand-rolled CSV (`DocumentTitle,DocumentPath,...`)
  parsed independently in each InNetwork command. Document titles containing commas will
  corrupt parsing. Either escape properly in one shared parser or switch the registry to
  JSON lines (the codebase already ships Newtonsoft.Json).
- Magic strings for the selection-set name (`"temp"`), file names, and parameter names
  are declared per-file rather than centrally.

## 6. Commented-out and vestigial code

- `DataGrid/DataGrid.EditMode.cs:31` — commented-out `_globalValidationDecision` field.
- `DataGrid/ColumnHandlers.Framework.cs.txt` — a whole proposed framework as `.txt`.
- `#if REVIT2011 || REVIT2012 || ...` guards including six years the project has never
  targeted (see architecture.md finding 5).
- 7 TODO/FIXME/HACK markers — actually a *low* count; fine.

## 7. `Nullable` disabled

`<Nullable>disable</Nullable>` with defensive `?.`/null-checks applied inconsistently by
hand. Flipping the whole project would produce thousands of warnings; instead enable it
file-by-file (`#nullable enable`) in new files and in anything touched during the Core
extraction. The Revit API itself is un-annotated, so expect `!` at API boundaries —
still worth it for internal code.

## 8. UI-thread data extraction

`ElementDataHelper.GetElementData` (DataGrid.cs:15) runs on the UI thread inside command
execution: for every element it reads parameters and intersects bounding boxes against
**every scope box in the document** (O(elements × scope boxes)), including linked-model
scope boxes when link mode is on. There is a cancellable progress dialog, which helps,
but on large selections this is the dominant cost of opening any picker — see
datagrid-ui.md, because this matters more than the choice of UI framework. Cheap wins
visible without profiling:

- Pre-transform scope-box bounding boxes once (currently plausible per-element work).
- Skip scope-box/centroid computation entirely unless those columns are requested.
- Build the search index lazily (first keystroke) instead of before first paint —
  `BuildSearchIndex` (DataGrid.Helpers.cs:318) walks every cell of every row up front.

## 9. Minor items

- `Helpers.GetMainWindowHandle(UIApplication uiApp)` ignores its parameter — remove it
  or use the API-appropriate handle per version.
- Natural-sort comparer, kebab-case converter, and header formatter live inside
  `CustomGUIs` but are general-purpose — Core candidates.
- `SelectionModeManager` persists mode in a bare file; fine, but reads it on every
  selection call — cache with a `FileSystemWatcher` or timestamp check if profiling ever
  shows it (low priority).
