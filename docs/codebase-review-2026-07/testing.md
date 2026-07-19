# Testing Strategy (Internal Logic Only)

Scope per your instruction: no Revit-in-a-VM integration testing yet. The question is
whether unit tests on internal logic are worth having now. **Yes — but only after a
small extraction step**, because today almost nothing is testable in isolation: logic
lives in static classes referencing `System.Windows.Forms` and `Autodesk.Revit.DB`,
and Revit API types (`UIDocument`, `Element`) cannot be instantiated outside Revit.

## What is genuinely unit-testable (and worth it)

These are pure or nearly-pure logic with real bug surface, currently trapped inside
UI/static classes:

| Logic | Current location | Why it deserves tests |
|---|---|---|
| Search/filter query parser (`$col` syntax, OR groups, exact-match, negation) | `DataGrid/DataGrid.Filtering.cs` | It's a small language; regressions here break every picker subtly |
| Row filtering against parsed groups | same | Core of daily UX |
| Natural sort comparer | `DataGrid/DataGrid.Sorting.cs` | Classic off-by-edge-case territory (numbers-in-strings, mixed case) |
| Kebab-case / header formatting | `DataGrid/DataGrid.Main.cs` | Used for history keys — silent corruption if wrong |
| Search-history store | `DataGrid/DataGrid.SearchHistory.cs` | File-format round-trip |
| `documents` CSV registry parse + heartbeat filtering | duplicated in InNetwork files (→ `SessionRegistry`, see duplication.md) | Comma-in-title bug is *known-shaped*; a test pins the fix |
| SelectionStorage SQLite round-trip | `SelectionStorage.cs` | Already path-parameterizable; UniqueId grouping logic |
| Column-edit validators | `DataGrid/DataGrid.Validation.cs` | Pure string/number checks |
| Installer shortcut merge/dedup | `installer/_installer-*.py` + Installer.cs | Corrupting `KeyboardShortcuts.xml` hurts |

Explicitly **not** worth unit-testing now: anything needing `Document`/`UIDocument`
(InputResolver, ElementDataHelper, all commands). Mocking the Revit API is a famous
time sink (interfaces don't exist; wrapping everything in shims costs more than it
returns). That's integration-test territory, later.

## The prerequisite: `RevitBallet.Core`

Create a `netstandard2.0` class library (consumable by every TFM in your matrix,
net46 → net8.0-windows) containing the logic above, with zero references to WinForms or
RevitAPI.dll. The commands project references it. This is the same extraction
recommended in datagrid-ui.md and architecture.md — three goals, one refactor:

```
core/RevitBallet.Core.csproj        (netstandard2.0)
  Filtering/FilterQuery.cs          ← parser, moved from DataGrid.Filtering.cs
  Filtering/RowFilter.cs
  Sorting/NaturalComparer.cs
  Text/NameConventions.cs           ← kebab-case, header formatting
  Network/SessionRegistry.cs        ← CSV parse (file IO injected as string/lines)
  Storage/SelectionStore.cs         ← SQLite path injected
tests/RevitBallet.Core.Tests.csproj (net8.0, xunit)
```

Key mechanical rule for the extraction: methods take data in, return data out. E.g. the
filter engine's signature becomes
`RowFilter.Apply(IReadOnlyList<Dictionary<string, object>> rows, FilterQuery query)`
and the WinForms layer keeps only the `DataGridView` binding. Most of the existing
filtering code already has this shape internally — it's a move, not a rewrite.

Payoff beyond safety: **tests run on Linux**, so the same GitHub Actions workflow that
matrix-compiles the ten Revit years (architecture.md finding 6) runs the suite on every
push — your first-ever automated verification, with no Windows or Revit anywhere.

## Example of the kind of test this enables

```csharp
[Theory]
[InlineData("door 900",            "Door-900x2100", true)]   // implicit AND
[InlineData("door||window",        "Window-600",    true)]   // OR groups
[InlineData("$category wall door", "Door-900x2100", false)]  // scoped column miss
public void FilterQuery_matches_expected_rows(string query, string cellValue, bool expected)
{
    var q = FilterQuery.Parse(query);
    var row = new Dictionary<string, object> { ["Name"] = cellValue, ["Category"] = "Doors" };
    Assert.Equal(expected, RowFilter.Matches(row, q));
}
```

(Adjust cases to the real query syntax — writing these tests will double as the first
precise documentation of what the syntax actually is, which today exists only as code.)

## Later: integration testing without a VM farm

Noting for the future, not now: you already own a better integration harness than most
Revit teams — the Roslyn server. A `tests/smoke/` directory of C# scripts POSTed to
`/roslyn` on the production machine (open a known test model, run a command's core via
the API, assert on element state, report `OK`/`FAIL`) gives real-Revit coverage with
zero new infrastructure. The `tests/` directory's existing shell scripts
(`health-check.sh`, `test-invoke-addin-command.sh`) are already embryonic versions of
exactly this — formalizing them into a runnable suite is the natural next step when
you're ready.
