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
