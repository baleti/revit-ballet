#!/bin/bash
# Build all Revit versions and the installer, anchored to this script's location.
# Run from any worktree: bash Build.sh
# Safe from both main tree and git worktrees — paths are always relative to this script.

set -e
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
COMMANDS_DIR="$REPO_ROOT/commands"
INSTALLER_DIR="$REPO_ROOT/installer"

echo "Repo root: $REPO_ROOT"
echo ""

echo "Building commands for Revit 2026..2017..."
FAILED=()
for year in $(seq 2026 -1 2017); do
    result=$(dotnet build "$COMMANDS_DIR" -c Release -p:RevitYear=$year 2>&1)
    errors=$(echo "$result" | grep -oP '\d+(?= Error)' | head -1)
    summary=$(echo "$result" | grep -E '^\s*(Build succeeded|Error\(s\))' | tail -1 | xargs)
    if [ "${errors:-0}" -gt 0 ]; then
        FAILED+=($year)
        echo "Revit $year: FAILED ($errors error(s))"
    else
        echo "Revit $year: OK"
    fi
done

if [ ${#FAILED[@]} -gt 0 ]; then
    echo ""
    echo "Build failed for: ${FAILED[*]}"
    exit 1
fi

echo ""
echo "Building installer..."
dotnet build "$INSTALLER_DIR" -c Release

echo ""
echo "Done. Installer at: $INSTALLER_DIR/bin/Release/installer.exe"
