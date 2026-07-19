# Duplication Analysis

A crude token metric (identical non-trivial source lines appearing >10 times across
`commands/*.cs`) finds ~176 such lines — consistent with heavy template copy-paste.
The meaningful duplication clusters into four groups, ordered by extraction value.

## 1. The InNetwork family — highest value

`SelectByCategoriesInNetwork.cs` (754 lines), `SelectByFamilyTypesInNetwork.cs` (868),
`SelectByWorksetsInNetwork.cs` (679), `CopyTypeParametersInNetwork.cs` (1,327),
`SynchronizeDocumentsInNetwork.cs`, `SwitchDocumentInNetwork.cs` each independently
re-implement:

- **Session discovery**: parsing the `documents` CSV registry — a private
  `DocumentInfo` class is declared in 4 separate files.
- **HTTPS client construction**: `new HttpClientHandler` + cert-bypass callback
  (10 copies) + `new HttpClient` (18 sites) + token header attachment.
- **Roslyn script POST**: build C# source as a string, POST to `/roslyn`, parse the
  `{Success, Output, Error}` JSON, split `Output` on pipe-delimited lines.
- **Diagnostics**: a private `WriteDiagnostic` helper exists in 5 files.

**Recommendation** — one infrastructure file, roughly:

```csharp
namespace RevitBallet.Infrastructure
{
    public record SessionInfo(string DocumentTitle, string DocumentPath,
        string SessionId, int Port, string Hostname, int ProcessId);

    public static class SessionRegistry
    {
        public static IReadOnlyList<SessionInfo> GetLiveSessions();      // parses documents CSV once, filters dead heartbeats
    }

    public static class NetworkClient   // single static HttpClient, cert policy in ONE place
    {
        public static RoslynResult ExecuteScript(SessionInfo session, string csharpSource);
    }
}
```

Estimated reduction: 2,000+ lines, and every future InNetwork command becomes
~100 lines of actual logic. This should happen *before* writing more InNetwork
commands (the CLAUDE.md scope roadmap says more are planned).

## 2. The scope-family pattern (`InView` / `InDocument` / `InSession` / `InNetwork`)

Example: `SelectByCategories*` — 461 + 305 + 383 + 754 lines. Diffing the InView and
InDocument variants shows the differences are exactly: (a) which element collector runs,
(b) the InView variant's 40-line "views from selection or active view" preamble that
`InputResolver.ResolveViews` already implements. Everything else — category grouping,
DataGrid invocation, selection application — is identical.

The same holds for `SelectByFamilyTypes*`, `SelectByWorksets*`, `SelectByMaterial*`,
`OpenSheets*`, `CloseViews*`, `SwitchView*`, `OpenPreviousViews*`.

**Recommendation**: parameterize the core by an element source instead of duplicating
per scope:

```csharp
public abstract class SelectByCategoriesBase : IExternalCommand
{
    protected abstract IEnumerable<(Document doc, Element el)> CollectElements(UIApplication app);
    public Result Execute(...)  { /* shared: group, DataGrid, select */ }
}

[CommandMeta("")] public class SelectByCategoriesInView : SelectByCategoriesBase { ... 10 lines ... }
[CommandMeta("")] public class SelectByCategoriesInDocument : SelectByCategoriesBase { ... 5 lines ... }
```

This is also precisely the "commands as nodes with typed inputs" direction of
`docs/architecture-vision.org` — a scope is just an element-source input. Doing this
family-by-family (start with SelectByCategories as the reference, as CLAUDE.md already
designates it) converges the codebase toward the vision doc without a big-bang rewrite.

## 3. Selection-or-picker-fallback preambles

The pattern "use selection if compatible, else active view, else DataGrid picker" is
hand-rolled inline in dozens of commands predating `InputResolver.cs` (only 8 adopters).
`SelectByCategoriesInView.cs:18-57` is a representative inline copy. The two recent
commits ("universal selection-first + picker-fallback...", "...11 more commands") show
this migration is active — finish it, then delete the inline copies. The
`command-standardization-plan.org` rename catalogue can ride along command-by-command.

## 4. Small utilities declared per-file

- `WriteDiagnostic` ×5 → `Log.Diag(operation, text)` in Infrastructure.
- CSV parsing of `documents` ×4 → `SessionRegistry`.
- Kebab-case conversion exists in `DataGrid.Main.cs` and (per the installer scripts)
  the installer's shortcut logic → one string-utils home.
- `GetSharedAuthToken` logic in `Server.cs` is duplicated in client code — expose the
  server's static method and delete the copy.

## Suggested sequencing

1. Create `commands/Infrastructure/` (or `RevitBallet.Core` per testing.md) with
   `SessionRegistry`, `NetworkClient`, `Log`.
2. Migrate one InNetwork command as reference; then the rest mechanically.
3. Pick one Select* family; build its scope-parameterized base; validate the pattern
   feels right in daily use.
4. Roll the pattern across families opportunistically (when a command needs touching
   anyway), rather than as a dedicated big-bang — each family conversion is
   independently shippable.
