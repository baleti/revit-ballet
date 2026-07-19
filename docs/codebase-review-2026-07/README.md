# Codebase Review — July 2026

Analysis of the revit-ballet codebase (~88,000 lines of C# across 290 files: 282 command
files, the DataGrid subsystem, the Roslyn server, and the installer). Analysis only — no
code was changed. Each document below covers one area; this page is the executive summary
and priority list.

## Documents

| Document | Contents |
|---|---|
| [architecture.md](architecture.md) | Structure, namespaces, static state, organizational findings |
| [code-smells.md](code-smells.md) | Concrete antipatterns with file references |
| [duplication.md](duplication.md) | Where code is copy-pasted and what to extract |
| [datagrid-ui.md](datagrid-ui.md) | WinForms DataGrid: performance analysis and framework alternatives |
| [testing.md](testing.md) | Unit-test strategy for internal logic (no Revit VM needed) |
| [security.md](security.md) | Server / network-command security notes |

## Executive summary

The codebase is in better shape than "organically grown over years while learning" would
suggest. The fundamentals are sound: multi-version build via a single parameterized
csproj is clean, git hygiene is good (no binaries tracked, 4.7 MB pack), CLAUDE.md
conventions are unusually thorough, the DataGrid already uses virtual mode with a search
index, and the Roslyn server design (ExternalEvent marshalling, heartbeat registry) is
correct for the Revit threading model.

The main problems are the predictable residue of copy-paste-driven growth:

1. **Namespace chaos** — 25 different namespaces including `YourNamespace`,
   `MyRevitCommands`, `YourCompany.YourAddin` (tutorial/AI-output leftovers), plus many
   types in the global namespace.
2. **Silent exception swallowing** — 168 bare `catch {` blocks plus 26 `catch (Exception)`
   with no logging. When something breaks in the field, there is no trail.
3. **Massive duplication in the `InNetwork` command family** — each of the 5+ network
   commands re-implements session discovery, HTTPS client construction, certificate
   bypass (10 copies), and token reading.
4. **The `Select*` scope families** (`InView`/`InDocument`/`InSession`/`InNetwork`)
   share 70–80% of their logic per family but are fully independent files.
   `InputResolver.cs` exists to fix part of this but only 8 files have adopted it —
   the April 2026 `command-standardization-plan.org` stalled at DRAFT.
5. **God-object UI layer** — `CustomGUIs` is a static partial class spread over 10+
   files (~8,800 lines) with ~20 static mutable fields; it is a global singleton that
   cannot be tested, is not reentrant, and leaks state between invocations.
6. **No CI** — nothing catches a per-Revit-year compile break before `Build.sh` is run
   manually. This is the cheapest, highest-value fix available: the Revit API comes from
   NuGet and `EnableWindowsTargeting` is already set, so a GitHub Actions matrix build
   (2017–2026) needs no Windows machine and no Revit install.

## On the DataGrid / WinForms question

Short answer to "is there a framework that could outperform the current implementation":
**no framework will beat a virtual-mode `DataGridView` at raw grid rendering** — it is
already one of the fastest grid surfaces on Windows. The latency you feel is dominated by
Revit API data extraction, upfront search-index construction, and first-open form/JIT
cost — none of which a framework swap fixes. **Flexibility** is a different story:
WebView2 + a JS grid is the strongest candidate there, and can be made to *feel* faster
than today via a pre-warmed hidden instance. Full analysis, including Avalonia and WPF,
in [datagrid-ui.md](datagrid-ui.md). The recommended first step is the same regardless of
framework: extract the filter/sort/column engine out of WinForms into a UI-agnostic core,
which also makes it unit-testable.

## Priority list

Ordered by value-to-effort ratio. Items 1–3 are mechanical and low-risk.

| # | Action | Effort | Payoff |
|---|---|---|---|
| 1 | Add GitHub Actions CI: matrix compile across all 10 Revit years | Hours | Every push verified; catches per-year breaks |
| 2 | Namespace unification to `RevitBallet.*` (scripted rename) | Hours | Discoverability, no tutorial leftovers, safe scripting-server reflection |
| 3 | Central `Log.Warn(ex)` helper; replace bare `catch {}` incrementally | Hours–days | Field failures become diagnosable |
| 4 | Extract `NetworkClient` + `SessionRegistry`; collapse InNetwork duplication | Days | −2,000+ lines, one place to fix TLS/token/discovery |
| 5 | Finish InputResolver adoption per the existing standardization plan | Days | −~3,000 lines across Select*/Filter* families |
| 6 | Extract UI-agnostic `RevitBallet.Core` (filter engine, sorting, column model) | 1–2 weeks | Unit tests on Linux CI; DataGrid backend becomes swappable |
| 7 | Convert `CustomGUIs` static state to an instance session object | With #6 | Reentrancy, testability |
| 8 | DataGrid perceived-latency work (profile first; warm form cache, streamed rows) | After #6 | Faster feel without framework risk |
| 9 | Optional: WebView2 grid prototype behind the Core interface | Exploratory | Answers the flexibility question empirically |

Items the review deliberately skipped: full integration testing (needs Revit VMs, per
your instruction), installer deep-dive, and per-command functional review of all 282
commands (spot-checked representative families instead).
