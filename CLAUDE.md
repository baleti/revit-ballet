# Revit Ballet - Code Conventions

## Naming Conventions

**IMPORTANT**: Revit Ballet uses **PascalCase** (also called UpperCamelCase) for all naming conventions, following standard C# conventions.

This codebase was partially ported from AutoCAD Ballet which used kebab-case. All kebab-case naming has been migrated to PascalCase to maintain consistency with C# standards.

### Naming Standards:
- **Classes**: PascalCase (e.g., `TransformSelectedElements`, `DataGrid2SearchHistory`)
- **Methods**: PascalCase (e.g., `GetRuntimeFilePath`, `ProcessCommand`)
- **Properties**: PascalCase (e.g., `CurrentView`, `SelectedElements`)
- **Variables**: camelCase (e.g., `appData`, `searchHistory`)
- **File Names**: PascalCase (e.g., `Server.cs`, `SearchboxQueries/`)
- **Directory Names**: PascalCase (e.g., `runtime/SearchboxQueries/`)

### Migration from AutoCAD Ballet:
- ❌ `searchbox-queries` → ✅ `SearchboxQueries`
- ❌ `server-cert.pfx` → ✅ `ServerCert.pfx`
- ❌ `InvokeAddinCommand-history` → ✅ `InvokeAddinCommandHistory`

**Note**: The project name `revit-ballet` itself uses kebab-case as it's a product name/brand identity, not a code identifier.

## Build Limitations

**IMPORTANT**: Do NOT attempt to build the installer using `cd installer && dotnet build -c Release`.

Due to limitations of the underlying 9p filesystem in the isolated environment, builds may take excessive amounts of time (over 5 minutes or may hang indefinitely).

Instead:
- Ask the user to test the build and return results if needed
- Focus on code changes and let the user handle compilation
- If build verification is required, request the user run the build command in their environment

## Testing Across Revit Versions

When making changes that could affect compatibility, it's recommended (though not required) to test builds across all supported Revit versions:

```bash
cd ~/revit-ballet/commands && for i in {2026..2017}; do dotnet build -c Release -p:RevitYear=$i; done
```

This ensures consistency across Revit 2017-2026. Use this when:
- Modifying core functionality
- Changing Revit API interactions
- Making significant refactoring changes
