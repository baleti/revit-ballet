# Manual Registry Cleanup Guide

If the PowerShell scripts don't work, you can manually delete the registry entries.

## Option 1: Using Registry Editor (GUI)

1. **Open Registry Editor**:
   - Press `Win+R`
   - Type `regedit`
   - Press Enter (may require admin approval)

2. **Navigate to Revit's registry location**:
   - Expand: `HKEY_CURRENT_USER`
   - Expand: `SOFTWARE`
   - Expand: `Autodesk`
   - Expand: `Revit`

3. **For each Revit version you have installed**:
   - Expand: `Autodesk Revit 2024` (or your version)
   - Click on: `CodeSigning`
   - In the right pane, you'll see GUID entries

4. **Delete revit-ballet GUIDs**:
   - Look for these GUIDs (revit-ballet entries):
     ```
     7c2e91a4-8f3d-4b2e-9a7c-1d8f4e6b3a2f
     4a9e6f2b-7d3c-4e8a-9b6f-2c7d4a9e1b3f
     8b3f7e2a-9c4d-4f1b-8a6e-3d7f2b9a1c4e
     2f6a9d3e-4b7c-4e2a-9d6f-1c8e4a7b2d9f
     5e8b2f7a-3c9d-4a1e-8b7f-2d9a1c6e4b3f
     ... (28 total)
     ```
   - Right-click each GUID → Delete
   - Or select all revit-ballet GUIDs and delete them

5. **Repeat for all Revit versions** (2024, 2023, 2022, etc.)

## Option 2: Using PowerShell Command Line

Open PowerShell and run these commands (one at a time):

```powershell
# For Revit 2024:
$path = "HKCU:\SOFTWARE\Autodesk\Revit\Autodesk Revit 2024\CodeSigning"
Remove-ItemProperty -Path $path -Name "7c2e91a4-8f3d-4b2e-9a7c-1d8f4e6b3a2f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "4a9e6f2b-7d3c-4e8a-9b6f-2c7d4a9e1b3f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "8b3f7e2a-9c4d-4f1b-8a6e-3d7f2b9a1c4e" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "2f6a9d3e-4b7c-4e2a-9d6f-1c8e4a7b2d9f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "5e8b2f7a-3c9d-4a1e-8b7f-2d9a1c6e4b3f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "9a4f7c2e-6b3d-4e9a-7c2f-1d8b4e9a6c3f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "3c7e9a2d-4f6b-4a8e-9c7f-2d1b4a8e9c6f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "6d9f2a3e-7c4b-4e1a-8d9f-3c2e4b7a9d1f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "1e8a3f7b-9d4c-4a2e-7e9f-4c1d3a8b2e7f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "7f3a9d2c-6e4b-4f8a-9d3e-2c7f1b4a8e9f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "4b9e7a3f-2d6c-4e1a-8b9e-3f7d2a4c1e8f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "8e2f9a4b-7c3d-4a6e-9f8e-1d2c4b7a9e3f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "2a7e4f9b-3d6c-4e8a-7f2e-9c4d1a8b3e6f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "9f3a7e2b-6d4c-4a1e-8f9e-3c2d4a7b1e9f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "5c8e2a7f-9b3d-4e6a-7c8f-2d9a1b4e6c9f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "1d9f3a6e-7c4b-4e2a-9f8e-4c1d2a7b3e8f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "6a3e9f7b-2d4c-4a8e-7e9f-1c3d4a2b8e7f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "3f7a2e9d-6c4b-4e1a-8f7e-2d9c1a4b6e3f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "8b4e9f2a-7d3c-4a6e-9b8f-3c1d2a7e4b9f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "2e7f9a3b-4c6d-4e8a-7f2e-9d1c4a8b3e7f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "9a6f3e7b-2d4c-4a1e-8f9e-3c7d1a2b4e8f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "4e8a2f9b-6c3d-4a7e-9f4e-1d2c8a7b3e9f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "7c2e4a9f-3d6b-4e1a-8c7f-2d9a1e4b7c3f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "1f9a6e3b-7c4d-4a2e-9f1e-4c7d3a8b2e6f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "5e3a8f2b-9d4c-4a6e-7e8f-1c2d9a4b7e3f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "8a9f2e3b-6d4c-4a7e-9a8f-3c1d2e7b4a9f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "2d7a9e4f-3c6b-4e1a-7d9e-2f8c1a4b6e7f" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $path -Name "6f3e9a2b-7d4c-4a8e-9f6e-1c3d4a2b8e9f" -ErrorAction SilentlyContinue
```

**For other Revit versions**, change `2024` to your version (2023, 2022, etc.)

## Option 3: Delete Entire CodeSigning Key (Nuclear Option)

⚠️ **Warning**: This removes ALL trusted addins for that Revit version, not just revit-ballet!

```powershell
# For Revit 2024:
Remove-Item -Path "HKCU:\SOFTWARE\Autodesk\Revit\Autodesk Revit 2024\CodeSigning" -Recurse -Force
```

Revit will recreate the CodeSigning key on next startup, but all addins will prompt for trust again.

## Will the Installer Recreate These Entries?

**YES!** When you run the installer again, it will:
1. Call `RegisterTrustedAddins()` function
2. Extract all 28 GUIDs from revit-ballet.addin
3. Register them in the registry for each detected Revit version

So you can safely delete them and reinstall.

## Verification

To verify entries were removed:

```powershell
# Check if revit-ballet entries still exist (should return nothing)
Get-ItemProperty -Path "HKCU:\SOFTWARE\Autodesk\Revit\Autodesk Revit 2024\CodeSigning" -Name "7c2e91a4-8f3d-4b2e-9a7c-1d8f4e6b3a2f" -ErrorAction SilentlyContinue
```

If it returns nothing (or error), the entry is gone.

## After Cleanup

After manually deleting:
1. Run `installer.exe` again
2. Installer will recreate all registry entries
3. Revit Ballet will load without security prompts
