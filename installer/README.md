# Revit Ballet Installer

Single unified installer supporting all Revit versions (2017-2026) with automatic DLL deduplication.

## Quick Start

```bash
# Build the installer
cd installer
./build.sh    # Linux/Mac
# or
build.ps1     # Windows

# Output: bin/Release/installer.exe (42 MB, supports all versions)
```

## How It Works

The installer uses **checksum-based deduplication** to minimize size:

1. **build.py** - Calculates MD5 checksums of all DLLs across all year directories
2. Identifies duplicates (e.g., `Newtonsoft.Json.dll` identical across 8 years)
3. Creates `resources/` with only unique files + `file-mapping.txt`
4. Installer reads mapping at runtime and copies each DLL to applicable year directories

**Result**: 131 MB → 41 MB (69% space savings)

### Build Steps

```bash
# Automated (recommended)
./build.sh      # Runs build.py + dotnet build

# Manual
python3 build.py                # Generate deduplicated resources/
dotnet build -c Release         # Build installer.exe
```

### Prerequisites
- .NET Framework 4.8 SDK
- Python 3 (for build.py)

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

The installer embeds deduplicated resources from `resources/`:
- `revit-ballet.addin` - Revit addin manifest
- `file-mapping.txt` - Maps resources to Revit years
- `_common.*.dll` - DLLs identical across ALL years (e.g., clipper_library.dll)
- `_YEAR.*.dll` - Year-specific or shared DLLs (e.g., _2017.Newtonsoft.Json.dll for 2017-2024)
- `KeyboardShortcuts.xml` - Template with External Tools shortcuts
- `revit-ballet-logo.png` - Logo for installer dialog (optional)

Total: 28 unique DLL files instead of 90+ without deduplication.

## Directory Structure

```
installer/
├── build.py                  # Deduplication script
├── build.sh / build.ps1      # Build scripts
├── installer.cs              # Main installer code
├── installer.csproj          # Embeds resources/
├── resources/                # Generated by build.py
│   ├── file-mapping.txt     # Resource → years mapping
│   ├── _common.*.dll        # Shared by all years
│   └── _YEAR.*.dll          # Year-specific files
└── bin/Release/
    └── installer.exe        # Final installer (42 MB)
```
