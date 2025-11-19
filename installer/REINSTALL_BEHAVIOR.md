# Reinstallation and DLL Updates

## What Happens When You Rebuild and Reinstall

### Current Behavior

When you make code changes, rebuild, and run the installer:

1. **Installer detects existing installation**
2. **Checks if shortcuts are missing**
   - If yes: Offers to add shortcuts only
   - If no: Asks "Revit Ballet is already installed. Reinstall?"
3. **If user clicks "Yes"**:
   - Overwrites all DLLs in **two locations**:
     - `%AppData%\Autodesk\Revit\Addins\YEAR\revit-ballet\` (Revit's addin location)
     - `%AppData%\revit-ballet\commands\bin\YEAR\` (Runtime location for InvokeAddinCommand)
   - Overwrites `revit-ballet.addin` file
   - Updates keyboard shortcuts
   - Re-registers trusted addins

### Will It Replace Old DLL with New One?

**Yes**, the installer **will correctly replace** the old DLL with the new one in **both locations**.

**No user action required** besides clicking "Yes" to reinstall.

## Why Two Locations?

### Location 1: Revit's Addins Folder
- **Path**: `%AppData%\Autodesk\Revit\Addins\{YEAR}\revit-ballet\`
- **Purpose**: Revit loads addins from here on startup
- **Contains**: revit-ballet.dll, dependencies, .addin file
- **Locked by**: Revit when running (file is loaded into memory)

### Location 2: Runtime Folder (for InvokeAddinCommand)
- **Path**: `%AppData%\revit-ballet\commands\bin\{YEAR}\`
- **Purpose**: InvokeAddinCommand loads commands dynamically from here
- **Contains**: revit-ballet.dll, dependencies
- **Loaded via**: `Assembly.Load(byte[])` from memory (no file locking!)
- **Development benefit**: Can be replaced while Revit is running

### How InvokeAddinCommand Works

InvokeAddinCommand is a special command that:
1. Reads the DLL from the runtime location into a byte array
2. Loads the assembly from memory using `Assembly.Load(assemblyBytes)`
3. This avoids locking the DLL file on disk
4. You can replace the DLL while Revit is still running!

**Code snippet from InvokeAddinCommand.cs:**
```csharp
// Lines 140-141: Load from memory to avoid file locking
byte[] assemblyBytes = File.ReadAllBytes(assemblyPath);
Assembly assembly = Assembly.Load(assemblyBytes);  // No file lock!
```

This means during development:
- Revit keeps the addin location DLL locked (normal Revit behavior)
- InvokeAddinCommand loads from runtime location without locking
- You can update runtime location DLL while Revit is running
- Next InvokeAddinCommand execution will use the new DLL!

## The Revit Locking Problem

### Important: Revit Locks DLLs

**If Revit is running**, Windows will prevent overwriting the DLL because:
- Revit loads the DLL on startup
- Windows locks loaded DLLs to prevent corruption
- Installer will fail with "file in use" error

### Two Solutions:

#### Solution 1: Close Revit Before Reinstalling (Recommended)

```
1. Close all Revit windows
2. Run installer.exe
3. Click "Yes" to reinstall
4. Start Revit to see changes
```

This is the cleanest approach.

#### Solution 2: Add DLL Lock Detection to Installer

We could modify the installer to:
1. Detect if DLL is locked
2. Show message: "Please close Revit before reinstalling"
3. Optionally: Show list of running Revit processes

Example code:
```csharp
private bool IsDllLocked(string dllPath)
{
    if (!File.Exists(dllPath))
        return false;

    try
    {
        using (FileStream stream = File.Open(dllPath, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            return false; // Not locked
        }
    }
    catch (IOException)
    {
        return true; // Locked
    }
}
```

## Recommended Development Workflow

### Option A: Close Revit Each Time
```bash
# 1. Make code changes
# 2. Close Revit
# 3. Build new DLL
dotnet build commands/revit-ballet.csproj /p:RevitYear=2024 -c Release

# 4. Rebuild installer with new DLL
dotnet build installer/installer.csproj -c Release

# 5. Run installer
./installer/bin/Release/installer.exe

# 6. Start Revit and test
```

### Option B: Manual DLL Copy (Faster for Development)
```bash
# 1. Make code changes
# 2. Build new DLL
dotnet build commands/revit-ballet.csproj /p:RevitYear=2024 -c Release

# 3. Copy to RUNTIME location (can do this while Revit is running!)
cp commands/bin/2024/revit-ballet.dll "$env:AppData\revit-ballet\commands\bin\2024\revit-ballet.dll"

# 4. Test by running InvokeAddinCommand in Revit
# No restart needed! InvokeAddinCommand will load the new DLL
```

**Option B is fastest** during development:
- **No need to close Revit!** (if testing via InvokeAddinCommand)
- Skip installer rebuild
- Directly copy to runtime location
- Run InvokeAddinCommand to test new code immediately

**Note**: If testing commands loaded at Revit startup (not via InvokeAddinCommand), you still need to:
```bash
# Close Revit first
cp commands/bin/2024/revit-ballet.dll "$env:AppData\Autodesk\Revit\Addins\2024\revit-ballet\revit-ballet.dll"
# Restart Revit
```

### Option C: Development Build Script
```bash
#!/bin/bash
# dev-update.sh - Quick development update script

YEAR=2024

# Build the DLL
echo "Building revit-ballet.dll for Revit $YEAR..."
dotnet build commands/revit-ballet.csproj /p:RevitYear=$YEAR -c Release

if [ $? -ne 0 ]; then
    echo "Build failed!"
    exit 1
fi

# Copy to RUNTIME location (works even if Revit is running)
RUNTIME_DIR="$APPDATA/revit-ballet/commands/bin/$YEAR"
if [ -d "$RUNTIME_DIR" ]; then
    cp commands/bin/$YEAR/revit-ballet.dll "$RUNTIME_DIR/"
    echo "✓ Runtime DLL updated: $RUNTIME_DIR"
    echo "  You can test changes immediately using InvokeAddinCommand (no restart needed)"
else
    echo "WARNING: Runtime directory not found. Run installer first."
fi

# Also copy to Addins location (only if Revit is NOT running)
if pgrep -x "Revit.exe" > /dev/null; then
    echo "⚠ Revit is running - skipping Addins folder update (file is locked)"
    echo "  To update startup commands, close Revit and run this script again"
else
    INSTALL_DIR="$APPDATA/Autodesk/Revit/Addins/$YEAR/revit-ballet"
    if [ -d "$INSTALL_DIR" ]; then
        cp commands/bin/$YEAR/revit-ballet.dll "$INSTALL_DIR/"
        echo "✓ Addins DLL updated: $INSTALL_DIR"
        echo "  Start Revit to see changes in startup commands"
    fi
fi
```

**This script intelligently handles both scenarios:**
- Always updates runtime location (works with Revit running)
- Only updates Addins location if Revit is closed

## Should We Add "Close Revit" Detection?

**Recommendation**: Add it to avoid confusing errors.

Would you like me to:
1. Add DLL lock detection to the installer?
2. Show a friendly "Please close Revit" message if DLLs are locked?
3. Optionally retry after user closes Revit?

This would make the reinstall experience smoother.

## Current Status

✅ Installer **does** overwrite DLLs correctly
✅ No special "restart" notification needed
⚠️ User must close Revit manually (or we add detection)

## Testing the Update Process

To test that reinstallation works:

1. Install version 1
2. Close Revit
3. Make a visible code change (e.g., change a command name)
4. Rebuild DLL and installer
5. Run installer, click "Yes" to reinstall
6. Start Revit
7. Verify new changes appear

If DLLs weren't replaced, you'd still see old behavior.
