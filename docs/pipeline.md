# Shell-Style Command Pipelines

Commands are designed to eventually be composable into pipelines, like PowerShell cmdlets. The syntax:

```
SelectByCategoriesInDocument --filter "view" | DataGrid --filter "$name structu" --pattern "{} suffix"
```

## Design Principles (not yet implemented)

- **Same classes, no separate code paths.** Every `IExternalCommand` doubles as a pipeline stage. With no args it runs interactively (shows pickers/DataGrid as today). With args it runs non-interactively, applying flags and passing results through.
- **Objects in the pipe, not text.** The pipe carries a live `ICollection<ElementId>` (or `SelectionStorage` for cross-document), not serialised strings. Revit elements are already objects.
- **`--filter`** pre-applies a search/predicate (e.g. DataGrid's search box, category filter). `$property` syntax matches element properties.
- **`--pattern`** applies a string transform to a property (e.g. rename, suffix).
- A future `RunPipeline` command will accept a pipeline string, parse `|`-separated stages, resolve each to a command class via reflection, and thread the selection through.
- `CommandMeta` Input already describes what each stage expects — this becomes the typed port in the pipeline.

## Analogy

Unix pipes pass text; PowerShell passes objects. We pass Revit elements. The interactive/non-interactive duality (no args = show UI, with args = headless) is the same pattern PowerShell cmdlets use. Vim-style macros record at the command level (stable), not the click level (brittle like Excel VBA or Revit journal files).
