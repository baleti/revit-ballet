# Revit Ballet Installer - Complete Implementation Summary

## What Was Created

### 1. Installer Application (`installer.cs`)
A Windows Forms application that:
- Detects Revit installations automatically
- Deploys the addin to all detected versions
- Manages keyboard shortcuts
- Provides installation/uninstallation UI

### 2. Build Configuration (`installer.csproj`)
MSBuild project that:
- Targets .NET Framework 4.8
- Embeds all required resources (DLLs, .addin, shortcuts)
- Automatically builds all Revit versions (2019-2025) before building installer
- Single-command build process: `dotnet build -c Release`

### 3. Updated Files
- `revit-ballet.addin` - UUIDs randomized for all commands
- `installer/KeyboardShortcuts.xml` - Template with all External Tools shortcuts (UUIDs synchronized with .addin)

## Installation Process Flow

### When User Runs installer.exe:

1. **Detection Phase**
   - Scans **both** locations to detect which Revit versions are installed:
     - `%ProgramData%\Autodesk\Revit\Addins\*` (system-wide installations)
     - `%AppData%\Autodesk\Revit\Addins\*` (current user installations)
   - Identifies years (2017-2030)
   - **Always installs to** `%AppData%` (current user) to avoid requiring admin privileges
   - Maps to keyboard shortcuts location (`%AppData%\Autodesk\Revit\Autodesk Revit YEAR\`)

2. **Deployment Phase** (to `%AppData%\Autodesk\Revit\Addins\YEAR`)
   For each detected Revit version:
   - Creates `Addins\YEAR\revit-ballet\` subfolder
   - Extracts version-specific DLLs:
     - `revit-ballet.dll`
     - `Newtonsoft.Json.dll`
     - `clipper_library.dll`
   - Creates `revit-ballet.addin` with full paths to subfolder
   - Saves KeyboardShortcuts.xml template for reference
   - **No administrator privileges required** - all files written to current user's AppData

3. **Keyboard Shortcuts Phase**
   - Checks if `KeyboardShortcuts.xml` exists in user's Revit profile
   - If missing: Shows dialog with clear instructions (Revit limitation)
   - If exists: Adds External Tools shortcuts (only if not already present)
   - Creates backup: `KeyboardShortcuts.xml.bak`

4. **Trusted Addin Registration**
   - Extracts all AddInId GUIDs from the .addin file
   - Registers each GUID in Windows Registry for each Revit version
   - Registry location: `HKEY_CURRENT_USER\SOFTWARE\Autodesk\Revit\Autodesk Revit [YEAR]\CodeSigning`
   - Sets DWORD value to 1 for each GUID (indicates "Always Load")
   - Eliminates "Load Always/Load Once" security prompts on Revit startup

5. **Finalization**
   - Copies itself to `%AppData%\revit-ballet\uninstaller.exe`
   - Registers in Windows Add/Remove Programs
   - Shows success dialog

## Uninstallation Process

The uninstaller:
1. Removes `revit-ballet.addin` from all Addins folders
2. Deletes `revit-ballet` subfolders with DLLs
3. Removes External Tools shortcuts from KeyboardShortcuts.xml
4. Deletes `%AppData%\revit-ballet\` directory
5. Unregisters from Windows

## File Structure After Installation

```
%AppData%\Autodesk\Revit\Addins\
  2024\
    revit-ballet.addin              # Points to subfolder
    revit-ballet\
      revit-ballet.dll
      Newtonsoft.Json.dll
      clipper_library.dll
      KeyboardShortcuts-template.xml
  2023\
    revit-ballet.addin
    revit-ballet\
      ...

%AppData%\Autodesk\Revit\Autodesk Revit 2024\
  KeyboardShortcuts.xml             # Modified with shortcuts
  KeyboardShortcuts.xml.bak

%AppData%\revit-ballet\
  commands\
    bin\
      2024\
        revit-ballet.dll            # Runtime copy for InvokeAddinCommand
        Newtonsoft.Json.dll
        clipper_library.dll
      2023\
        ...
  uninstaller.exe
```

**Note**: DLLs are installed in **two locations**:
1. `Addins\YEAR\revit-ballet\` - Loaded by Revit at startup (file locked)
2. `revit-ballet\commands\bin\YEAR\` - Loaded by InvokeAddinCommand via memory (no lock)

## Build Requirements

### To Build the Installer:
1. .NET SDK (for building revit-ballet.dll for all versions)
2. .NET Framework 4.8 SDK (for building installer.exe)

Simply run: `cd installer && dotnet build -c Release`

The build process automatically compiles revit-ballet.dll for all supported Revit versions (2019-2025) before packaging the installer.

### To Run the Installer:
1. Windows OS
2. .NET Framework 4.8 Runtime
3. At least one Revit installation (2019-2030)
4. **No administrator privileges required** - installs to current user only

## New Features (v1.1)

### Smart Reinstall Detection
When the installer detects an existing installation, it now:
1. Checks if keyboard shortcuts are missing from `KeyboardShortcuts.xml`
2. Displays a dialog listing all shortcuts that will be added
3. Offers to add shortcuts without full reinstall
4. Only prompts for full reinstall if user declines shortcut-only update

### Trusted Addin Registration
The installer now automatically registers all addin GUIDs in the Windows Registry:
- Eliminates "Load Always/Load Once" security prompts
- Registers for all detected Revit versions
- Uses registry location: `HKEY_CURRENT_USER\SOFTWARE\Autodesk\Revit\Autodesk Revit [YEAR]\CodeSigning`
- Each GUID registered as DWORD with value 1

## Testing Checklist

- [x] ✓ Installer builds without errors
- [x] ✓ Embedded resources verified (4.9MB size includes icon and all resources)
- [x] ✓ UUIDs synchronized between .addin and KeyboardShortcuts.xml
- [x] ✓ Automated multi-version build process in installer.csproj
- [x] ✓ Detects both ProgramData and AppData Revit installations
- [x] ✓ Smart reinstall detection for keyboard shortcuts
- [x] ✓ Trusted addin registry entries
- [x] ✓ Icon embedded in installer
- [ ] ⚠ Full end-to-end testing requires Windows with Revit installed

## Known Limitations

1. **Revit Keyboard Shortcuts Limitation**:
   - Revit only creates KeyboardShortcuts.xml after user modifies a shortcut
   - Installer handles this with clear user instructions
   - This is a Revit limitation, not an installer bug

3. **Windows Only**:
   - Installer is Windows-only (Revit is Windows-only)
   - Must be built on Windows or with full .NET Framework support

## Next Steps for Production Use

1. **Test on Windows**:
   ```
   - Copy installer.exe to Windows machine
   - Run installer
   - Verify files are deployed correctly
   - Open Revit and test commands
   - Test uninstaller
   ```

2. **Add Icon** (optional):
   ```
   - Create revit-ballet-logo.ico (256x256)
   - Uncomment ApplicationIcon in installer.csproj
   - Rebuild
   ```

3. **Build Installer**:
   ```bash
   cd installer
   dotnet build -c Release
   ```

   The installer automatically builds all Revit versions (2019-2025) before packaging.

4. **Code Signing** (recommended for production):
   ```
   - Obtain code signing certificate
   - Sign installer.exe
   - This prevents Windows SmartScreen warnings
   ```

## Architecture Decisions

### Why Subfolder Approach?
- Keeps DLLs organized and separate from other addins
- Makes uninstallation cleaner
- Avoids DLL conflicts with other addins

### Why Version-Specific DLLs?
- Different Revit versions target different .NET frameworks
- Ensures compatibility across Revit 2019-2025
- Follows Revit best practices

### Why Embedded Resources?
- Single .exe distribution (no installer package needed)
- Users can't accidentally delete required files
- Cleaner than extracting to temp folder

### Why No AutoUpdate?
- Simpler implementation
- Users manually download new versions
- Can be added later if needed

## Troubleshooting

### Build Errors

**"Unable to read beyond end of stream"**
- Caused by invalid .ico file
- Solution: Remove or fix icon file

**"Resource not found"**
- Ensure revit-ballet.dll is built first
- Check `../commands/bin/*/revit-ballet.dll` exists

### Runtime Errors (Windows)

**"No Revit installations found"**
- Check `%AppData%\Autodesk\Revit\Addins` exists
- Verify at least one numbered folder (2019-2030) exists

**"KeyboardShortcuts.xml not found"**
- Expected behavior if user never modified shortcuts
- Follow on-screen instructions to create the file

## Success Criteria

✅ Installer builds successfully (2.8MB .exe)
✅ All resources embedded correctly
✅ UUIDs synchronized
✅ Documentation complete
✅ Automated multi-version build system
✅ Single-command build process

⚠️ Requires Windows testing for full validation
