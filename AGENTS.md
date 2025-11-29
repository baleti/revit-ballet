# Revit Ballet - Agent Development Guide

## Code Conventions for AI Agents

When working on the Revit Ballet codebase, follow these naming conventions:

### Naming Standards (C# PascalCase)

**CRITICAL**: This project uses **PascalCase** for all code identifiers, NOT kebab-case.

- **Classes/Interfaces**: PascalCase (e.g., `TransformSelectedElements`)
- **Methods/Functions**: PascalCase (e.g., `GetRuntimeFilePath`)
- **Properties**: PascalCase (e.g., `CurrentView`)
- **Local Variables**: camelCase (e.g., `appData`, `currentView`)
- **Constants**: PascalCase or UPPER_CASE (e.g., `MaxRetries` or `MAX_RETRIES`)
- **File Names**: PascalCase matching class name (e.g., `Server.cs`)
- **Directory Names**: PascalCase (e.g., `SearchboxQueries/`)

### Migration Note

This codebase was ported from AutoCAD Ballet which used kebab-case. All kebab-case has been standardized to PascalCase.

Examples of corrections made:
- `searchbox-queries/` → `SearchboxQueries/`
- `server-cert.pfx` → `ServerCert.pfx`
- `InvokeAddinCommand-history` → `InvokeAddinCommandHistory`

### Exception

The product name `revit-ballet` remains in kebab-case as it's a brand identifier, not a code element.

## Development Guidelines

When creating or modifying code:
1. Always use PascalCase for new files, classes, and methods
2. Convert any kebab-case to PascalCase when refactoring
3. Follow existing C# conventions in the codebase
4. Update references when renaming files or directories

## Build Limitations

**IMPORTANT**: Do NOT attempt to build the installer yourself using `cd installer && dotnet build -c Release`.

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
