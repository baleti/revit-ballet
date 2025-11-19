#!/bin/bash

# Verification script for Revit Ballet Installer build
# This checks that all required files are present and the installer is built correctly

set -e

echo "Revit Ballet Installer - Build Verification"
echo "============================================"
echo ""

# Check if installer.exe exists
if [ ! -f "bin/Release/installer.exe" ]; then
    echo "❌ ERROR: installer.exe not found. Run 'dotnet build -c Release' first."
    exit 1
fi

echo "✓ installer.exe exists"

# Check installer size
INSTALLER_SIZE=$(stat -f%z "bin/Release/installer.exe" 2>/dev/null || stat -c%s "bin/Release/installer.exe" 2>/dev/null)
if [ "$INSTALLER_SIZE" -lt 1000000 ]; then
    echo "❌ WARNING: installer.exe is smaller than expected ($INSTALLER_SIZE bytes)"
    echo "   Expected at least 1MB. Resources may not be embedded correctly."
else
    echo "✓ installer.exe size is reasonable ($(numfmt --to=iec-i --suffix=B $INSTALLER_SIZE 2>/dev/null || echo "${INSTALLER_SIZE} bytes"))"
fi

# Check source files
echo ""
echo "Checking source files:"
for file in installer.cs installer.csproj KeyboardShortcuts.xml ../revit-ballet.addin; do
    if [ -f "$file" ]; then
        echo "  ✓ $file"
    else
        echo "  ❌ $file MISSING"
    fi
done

# Check for at least one DLL version
echo ""
echo "Checking for built DLLs:"
DLL_COUNT=$(find ../commands/bin -name "revit-ballet.dll" 2>/dev/null | wc -l)
if [ "$DLL_COUNT" -eq 0 ]; then
    echo "  ❌ No revit-ballet.dll found in ../commands/bin/"
    echo "     Build revit-ballet first: dotnet build ../commands/revit-ballet.csproj /p:RevitYear=2024"
else
    echo "  ✓ Found $DLL_COUNT version(s) of revit-ballet.dll:"
    find ../commands/bin -name "revit-ballet.dll" 2>/dev/null | sed 's/^/    /'
fi

# Check dependencies
echo ""
echo "Checking dependencies:"
for year_dir in ../commands/bin/*/; do
    year=$(basename "$year_dir")
    if [ -f "../commands/bin/$year/Newtonsoft.Json.dll" ] && [ -f "../commands/bin/$year/clipper_library.dll" ]; then
        echo "  ✓ Year $year has all dependencies"
    fi
done

echo ""
echo "============================================"
echo "Verification complete!"
echo ""
echo "To distribute:"
echo "  cp bin/Release/installer.exe revit-ballet-installer.exe"
echo ""
echo "Note: This installer requires Windows with .NET Framework 4.8 to run."
