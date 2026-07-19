# DataGrid UI: Performance Analysis & Framework Alternatives

Question asked: *"I want to move from showing DataGrid with WinForms to perhaps
something faster and more flexible — is that feasible, i.e. is there any framework that
could outperform the current implementation?"*

## TL;DR

- **Faster raw grid rendering: no.** A virtual-mode `DataGridView` painting GDI text
  is already at or near the ceiling for on-screen tabular throughput on Windows. No
  mainstream framework renders a filtered list of rows meaningfully faster than what
  you have.
- **Faster *perceived* open-to-first-keystroke: yes — but through architecture, not a
  framework.** The latency you feel is Revit data extraction + upfront search-index
  build + first-open form creation/JIT, all of which survive any framework swap
  unchanged.
- **More flexible: yes.** WebView2 + a JS grid is the strongest flexibility play and
  can *also* feel faster than today if kept pre-warmed. Avalonia's TreeDataGrid is the
  best native-code alternative but carries real deployment risk inside Revit's shared
  process.
- **Either way, the first step is identical**: extract the filter/sort/column engine
  from `CustomGUIs` into a UI-agnostic core so backends are swappable and the question
  becomes a cheap experiment instead of a rewrite bet.

## Where the current implementation's time actually goes

From reading the code (profile to confirm — see "Measure first" below):

1. **Data extraction** (`ElementDataHelper.GetElementData`, DataGrid.cs): per element —
  parameter reads through the Revit API (each one a native interop call), scope-box
  bounding-box intersection against *all* scope boxes (plus linked-model scope boxes in
  link mode), centroid computation. On a picker over thousands of elements this
  dominates everything else, and it runs on the UI thread before the form appears.
2. **Search-index construction** (`BuildSearchIndex`, DataGrid.Helpers.cs:318): called
  in `DataGrid()` at line 201 *before first paint* — stringifies every cell of every row
  into per-column and all-column indexes. For a 50k-row × 20-column grid that's a
  million string operations before the user sees anything, to accelerate a search they
  haven't typed yet.
3. **Form construction + first-use JIT**: a fresh `Form`, `DataGridView`, ~20 event
  hookups, `AutoResizeColumns` per invocation. First invocation in a session also pays
  JIT for the whole DataGrid partial class.
4. **Grid painting**: virtual mode already means only visible cells materialize. This is
  the part people assume is slow and it isn't.

Consequence: **a framework swap replaces only item 4 — the part that is already fast.**

## Cheap wins inside the current implementation

Ordered; all are independent of any framework decision:

1. **Lazy search index** — build on first keystroke (or in the background after the form
   is shown). Removes the entire index cost from time-to-first-paint.
2. **Show the form immediately, stream rows in** — virtual mode makes this natural:
   show with `RowCount = 0`, extract data on a background thread in batches, bump
   `RowCount` per batch (marshalled via `BeginInvoke`). The user can start typing while
   extraction continues. This is the change that makes big-model pickers feel instant.
3. **Column-lazy extraction** — skip scope-box/centroid work unless those columns are
   requested by the calling command (many pickers never show them).
4. **Warm form cache** — keep one hidden `Form`+`DataGridView` instance alive per Revit
   session and re-bind data per invocation. Eliminates per-open construction and layout;
   also makes font-size/screen-state persistence natural instead of static-field-based.
5. **`Ctrl`-level micro-items**: `DoubleBuffered` on the grid via reflection if not
   already, `BeginUpdate`-style suspend during rebinds, avoid `grid.RowCount = 0; grid.RowCount = n`
   double-reset (forces two layout passes) in `UpdateFilteredGrid` (DataGrid.Main.cs).

If, after items 1–4, opening a picker still feels slow, the bottleneck is the Revit API
itself and *no* UI change will help.

## Framework alternatives, honestly assessed

Constraints that matter: runs inside Revit's process (shared with other addins),
must target net46→net8 across Revit 2017–2026, keyboard-first modal picker workflow,
row counts occasionally in the tens of thousands.

### WPF (Revit's own native UI stack)

- Already loaded in the Revit process (`UseWPF` is even enabled in the csproj) — zero
  deployment risk, no new dependencies.
- Stock WPF `DataGrid` is *slower* than your virtual WinForms grid; acceptable only
  with row+column virtualization, recycling mode, and deferred scrolling carefully
  enabled. Third-party (DevExpress/Syncfusion) WPF grids are fast but commercial.
- Gains: styling, DPI scaling, data templates (inline swatches, icons), smoother
  composition. Performance: lateral at best, easy to regress.
- **Verdict**: the safe modernization path if you want nicer visuals with zero
  deployment risk, but it does not "outperform" — and you'd rewrite ~9k lines of
  interaction code for a sideways move.

### Avalonia + TreeDataGrid

- `TreeDataGrid` is arguably the fastest open-source .NET grid (built for
  million-row virtualization); Avalonia 11 supports .NET Framework 4.6.2+ and .NET 8,
  matching your matrix (2017–2018 on net46 would be the sticking point — 4.6 vs 4.6.2).
- Risks inside Revit: Avalonia + SkiaSharp native binaries loaded into a process where
  *other addins* may load different SkiaSharp versions (assembly/native-DLL version
  conflicts are a classic Revit-addin failure mode); needs its own dispatcher pumped in
  the host window. Embedding Avalonia in a Win32 host app works, but you'd be the
  integration test.
- **Verdict**: highest raw-grid ceiling of the native options, but the performance
  headroom is above a bottleneck you don't have, and the deployment risk is real.
  Not recommended unless the WinForms grid itself (item 4 above) is ever measured as
  the constraint — which is unlikely.

### WebView2 + JS grid (AG Grid Community / Tabulator / custom virtual list)

- **Flexibility winner by far**: rich styling, fuzzy-search UX, column pinning/grouping,
  inline editors, themes — all cheap in HTML/JS; iteration speed is much higher than
  WinForms owner-draw work.
- Performance truth: AG Grid comfortably virtualizes 100k+ rows; *warm* WebView2 render
  of a filtered list is competitive with native. The costs are (a) cold-start of the
  WebView2 environment (~hundreds of ms first time — mitigate by creating one hidden
  WebView at addin startup and reusing it: the same warm-form-cache trick as above),
  (b) marshalling row data across the boundary — for large sets, serve it from the
  **Roslyn server you already run** (the grid page fetches
  `https://127.0.0.1:{port}/griddata` in chunks) instead of `PostWebMessage`-ing a giant
  JSON blob.
- Deployment: WebView2 Evergreen runtime is preinstalled on current Win10/11 and on any
  machine with Office/Teams; near-zero risk for your environment. Requires
  net462+ for the WebView2 SDK — again 2017–2018 (net46) would need to keep the
  WinForms fallback (they already run a reduced command set per readme).
- **Verdict**: the right experiment if flexibility is the goal. Behind a common
  interface it can coexist with the WinForms backend per Revit year.

### Custom owner-drawn control (DirectWrite/Direct2D or GDI double-buffer)

Maximum theoretical performance, large effort, zero flexibility gain over what a virtual
DataGridView already achieves. Not worth it.

## Recommended path

1. **Measure first.** Add a stopwatch to `CustomGUIs.DataGrid`: extraction ms, index ms,
   form-to-shown ms, per-keystroke filter ms, logged to the diagnostics folder. One
   day's usage tells you where the time really goes; every decision below gets cheaper.
2. **Extract the engine** (`RevitBallet.Core`): filter parser + row filter, sort
   comparers, search index, column model, edit-validation rules — no `System.Windows.Forms`
   references. The WinForms code becomes a thin view. (Also unblocks unit tests —
   see testing.md.)
3. **Do the cheap wins** (lazy index, streamed rows, warm form). Expect this alone to
   deliver the "faster" you're after.
4. **Prototype WebView2 as a second backend** behind the same interface for one command
   (e.g. `SelectByCategoriesInView`). Compare feel side-by-side for a week of real use.
   Keep whichever wins; the loser cost you a prototype, not a rewrite.
