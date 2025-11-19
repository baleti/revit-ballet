# Revit Ballet Installer

This installer packages Revit Ballet and deploys it to all detected Revit installations.

## Quick Start

```bash
# Build the installer (automatically builds all Revit versions)
cd installer
dotnet build -c Release

# Installer will be at: bin/Release/installer.exe
```

## Building the Installer

### Prerequisites
- .NET Framework 4.8 SDK

### Build Steps

The installer project automatically builds all supported Revit versions (2019-2025) before building itself:

```bash
cd installer
dotnet build -c Release
```

The installer executable will be in: `installer/bin/Release/installer.exe`

**How it works**: The installer.csproj includes a `BuildAllRevitVersions` target that runs before the build, automatically compiling revit-ballet.dll for all Revit versions and embedding them into the installer.

## What the Installer Does

1. **Detects Revit Installations**: Scans both locations to detect which Revit versions are installed:
   - `%ProgramData%\Autodesk\Revit\Addins\*` (system-wide installations)
   - `%AppData%\Autodesk\Revit\Addins\*` (current user installations)

   **Note**: The installer detects Revit from both locations but always installs to `%AppData%` (current user) to avoid requiring administrator privileges. Revit loads addins from both locations, so this works seamlessly.

2. **Smart Reinstall Detection**: When already installed, the installer:
   - Checks if keyboard shortcuts are missing
   - Offers to add shortcuts without full reinstall
   - Shows list of shortcuts that will be added
   - Allows full reinstall if needed

3. **Deploys Files** (to `%AppData%\Autodesk\Revit\Addins\YEAR`):
   - Creates `revit-ballet` subfolder in each `Addins\YEAR` directory
   - Copies `revit-ballet.dll` and dependencies (Newtonsoft.Json.dll, clipper_library.dll)
   - Creates `revit-ballet.addin` file with paths pointing to the subfolder

4. **Keyboard Shortcuts**:
   - Checks if `KeyboardShortcuts.xml` exists in `%AppData%\Autodesk\Revit\Autodesk Revit YEAR\`
   - If file doesn't exist, shows a dialog explaining the Revit limitation and how to create it
   - If file exists, adds shortcuts for all External Tools commands from the template
   - Only adds shortcuts that don't already exist
   - Creates backup as `KeyboardShortcuts.xml.bak`

5. **Trusted Addin Registration**:
   - Registers all addin GUIDs in Windows Registry
   - Location: `HKEY_CURRENT_USER\SOFTWARE\Autodesk\Revit\Autodesk Revit [YEAR]\CodeSigning`
   - Prevents "Load Always/Load Once" security prompts
   - Users won't see security warnings when Revit starts

6. **Uninstaller Registration**:
   - Copies installer to `%AppData%\revit-ballet\uninstaller.exe`
   - Registers in Windows Add/Remove Programs

## Revit Keyboard Shortcuts Limitation

Revit only creates `KeyboardShortcuts.xml` after the user modifies at least one shortcut. If the file doesn't exist:

1. Open Revit
2. Go to: View > User Interface > Keyboard Shortcuts
3. Modify any shortcut (e.g., add "ET" to "Type Properties")
4. Click OK
5. Re-run the installer

The installer will detect the file and add all revit-ballet shortcuts.

## Uninstalling

Run `uninstaller.exe` from `%AppData%\revit-ballet\` or use Windows Add/Remove Programs.

The uninstaller will:
- Remove `revit-ballet.addin` from all Revit Addins folders
- Delete `revit-ballet` subfolders with DLLs
- Remove keyboard shortcuts from `KeyboardShortcuts.xml`
- Delete backup files
- Unregister from Windows

## Icon Files

The installer includes:
- `revit-ballet-logo.ico` - Icon for the executable (displayed in Windows Explorer, Add/Remove Programs)
- `revit-ballet-logo.png` - Logo for the installer dialog window (optional, embedded if present)

To replace with your own icon:
1. Replace `revit-ballet-logo.ico` with your .ico file (256x256 or 128x128 recommended)
2. Optionally replace `revit-ballet-logo.png` for the installer dialog
3. Rebuild the installer

The icon appears in:
- Windows Explorer (exe file icon)
- Windows Add/Remove Programs
- Installer window during installation (if PNG exists)

## Files Embedded in Installer

The installer embeds:
- `revit-ballet.addin` - Revit addin manifest
- `revit-ballet.dll` (per Revit version) - Main addin assembly
- `Newtonsoft.Json.dll` (per Revit version) - JSON library dependency
- `clipper_library.dll` (per Revit version) - Clipper library dependency
- `KeyboardShortcuts.xml` - Template with External Tools shortcuts
- `revit-ballet-logo.png` - Logo for installer dialog (optional)
