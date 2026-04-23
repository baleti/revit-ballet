# Revit Ballet Installer

## Overview

The installer is a self-contained Windows Forms executable (`installer.exe`) that deploys the Revit Ballet plugin to all detected Revit versions on the machine. It is built on Linux and deployed to Windows via SCP.

## Build & Deploy Pipeline

```bash
# 1. Build commands for target Revit version(s)
cd ~/revit-ballet/commands
dotnet build -c Release -p:RevitYear=2024    # single version
# or all versions:
for i in {2026..2017}; do dotnet build -c Release -p:RevitYear=$i; done

# 2. Build installer (embeds built DLLs as resources, deduplicates shared files)
cd ~/revit-ballet/installer
dotnet build -c Release

# 3. SCP to Windows and run
scp -i ~/.ssh/hwo-18 installer/bin/Release/installer.exe daniel@192.168.231.1:'C:\Users\Daniel\Desktop\installer.exe'
ssh daniel@192.168.231.1 -i ~/.ssh/hwo-18 'powershell -NoProfile -Command "C:\Users\Daniel\Desktop\installer.exe --quiet --hot-reload"'
```

> **Always pass `--hot-reload`** so new commands are immediately available via `InvokeAddinCommand` without restarting Revit.

## Command-Line Flags

| Flag | Description |
|------|-------------|
| `--quiet` / `-q` | Run headlessly (no UI dialogs); print results to stdout. Required for SSH use. |
| `--verbosity 0\|1\|2` | Controls output detail in quiet mode (default: 1). |
| `--hot-reload` | Also copy DLLs to `%appdata%\revit-ballet\hot-reload\{year}\` for immediate pickup by `InvokeAddinCommand`. |
| `/uninstall` | Run uninstaller instead of installer. |

### Verbosity levels

- **0** — one-line summary: `updated: 2025; pending restart: 2024`
- **1** — grouped list with restart notice (default)
- **2** — full trace: detected installs, exact paths, per-file operations, registry steps

## What the Installer Does

1. **Detects Revit installations** — scans `%appdata%\Autodesk\Revit\Addins\{year}\` for years 2017–2026
2. **Extracts embedded resources** — DLLs are embedded per-year with deduplication (shared files stored once)
3. **Deploys to Addins folder** — copies `revit-ballet.dll` and dependencies to `Addins\{year}\revit-ballet\`
4. **Handles locked files** — if Revit is running, writes to a `revit-ballet.update\` folder instead; Startup.cs applies it on next Revit launch
5. **Deploys keyboard shortcuts** — merges custom shortcuts into `KeyboardShortcuts.xml` per Revit version
6. **Registers uninstaller** — adds entry to Windows Add/Remove Programs
7. **Registers trusted addins** — adds to Revit's addin trust list
8. **Hot-reload dir** *(with `--hot-reload`)* — copies DLLs to `%appdata%\revit-ballet\hot-reload\{year}\`

## Hot-Reload Mechanism

When Revit is running, the Addins folder DLL is locked by Revit's process. Two mechanisms allow picking up new code without restarting Revit:

### `.update` folder (automatic)
When the Addins DLL is locked, the installer writes to `Addins\{year}\revit-ballet.update\` instead. `Startup.cs` detects this folder on next Revit launch and swaps it in before any commands run.

### Hot-reload dir (immediate, via `--hot-reload`)
The installer also writes to `%appdata%\revit-ballet\hot-reload\{year}\`. `InvokeAddinCommand` checks this directory **first** on every invocation, loading the DLL via `Assembly.Load(bytes)` (no file lock). This means a build-deploy cycle makes new commands available immediately — no Revit restart needed.

```
%appdata%\revit-ballet\
  hot-reload\
    2024\
      revit-ballet.dll    ← InvokeAddinCommand reads this first
      *.dll               ← dependencies
    2025\
      ...
```

## Output Examples

```
# --verbosity 0
updated: 2025; pending restart: 2024

# --quiet (default verbosity 1)
Successfully updated:
  - Revit 2025
Updated (restart Revit to use new version):
  - Revit 2024
Restart Revit to load updates or use commands immediately via InvokeAddinCommand.

# --quiet --verbosity 2
Already installed: True
Loading file mapping...
Detecting Revit installations...
  Found: Revit 2024 -> C:\Users\Daniel\AppData\Roaming\Autodesk\Revit\Addins\2024
  Found: Revit 2025 -> C:\Users\Daniel\AppData\Roaming\Autodesk\Revit\Addins\2025
Target dir: C:\Users\Daniel\AppData\Roaming\revit-ballet
Extracting files...
  [Revit 2024] Updating -> ...\Addins\2024\revit-ballet
  [Revit 2024] Files locked, using update folder...
    update folder: ...\Addins\2024\revit-ballet.update
  [Revit 2024] Hot-reload dir updated -> ...\revit-ballet\hot-reload\2024
  [Revit 2025] Updating -> ...\Addins\2025\revit-ballet
    revit-ballet.dll
    revit-ballet.addin -> ...\Addins\2025\revit-ballet.addin
  [Revit 2025] Hot-reload dir updated -> ...\revit-ballet\hot-reload\2025
...
```

## File Layout (installed)

```
%appdata%\Autodesk\Revit\Addins\{year}\
  revit-ballet.addin              ← Revit plugin manifest
  revit-ballet\
    revit-ballet.dll              ← main plugin DLL
    *.dll                         ← dependencies (Roslyn, SQLite, etc.)
  revit-ballet.update\            ← staging area when files locked (cleaned on next launch)

%appdata%\revit-ballet\
  hot-reload\{year}\              ← hot-reload copies (--hot-reload only)
  uninstaller.exe
  network\token
  documents
  ...
```

## Implementation Notes

- Built as a .NET Framework 4.8 WinForms app (runs on Windows without .NET 8)
- DLLs are embedded as manifest resources during `dotnet build` via a Python pre-build script that deduplicates shared files
- The installer can also function as an uninstaller when renamed to contain "uninstall" or called with `/uninstall`
- Source: `installer/Installer.cs`
