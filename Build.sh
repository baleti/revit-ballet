#!/bin/bash
# Build all Revit versions and the installer, anchored to this script's location.
# Run from any worktree: bash Build.sh
# Safe from both main tree and git worktrees — paths are always relative to this script.

set -e
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
COMMANDS_DIR="$REPO_ROOT/commands"
INSTALLER_DIR="$REPO_ROOT/installer"


# Optional Authenticode signing (set REVIT_BALLET_SIGN_PFX [+ REVIT_BALLET_SIGN_PASS]; needs osslsigncode).
# Signed DLLs load in Revit without the "unsigned add-in" prompt and trigger fewer antivirus false positives.
# Never commit the PFX. Timestamp server: REVIT_BALLET_SIGN_TS (default DigiCert).
sign_file() {
    [ -n "${REVIT_BALLET_SIGN_PFX:-}" ] || return 0
    command -v osslsigncode >/dev/null || { echo "REVIT_BALLET_SIGN_PFX is set but osslsigncode is not installed"; exit 1; }
    osslsigncode sign -pkcs12 "$REVIT_BALLET_SIGN_PFX" -pass "${REVIT_BALLET_SIGN_PASS:-}" -h sha256 \
        -t "${REVIT_BALLET_SIGN_TS:-http://timestamp.digicert.com}" -n "Revit Ballet" -in "$1" -out "$1.signed" >/dev/null \
        && mv -f "$1.signed" "$1" && echo "  signed $(basename "$1")" || { echo "signing failed: $1"; exit 1; }
}

echo "Repo root: $REPO_ROOT"
echo ""

echo "Building commands for Revit 2026..2020..."
FAILED=()
for year in $(seq 2026 -1 2020); do
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

if [ -n "${REVIT_BALLET_SIGN_PFX:-}" ]; then
    echo ""
    echo "Signing revit-ballet.dll for each Revit year..."
    for year in $(seq 2026 -1 2020); do sign_file "$COMMANDS_DIR/bin/$year/revit-ballet.dll"; done
fi

echo ""
echo "Building installer..."
dotnet build "$INSTALLER_DIR" -c Release

if [ -n "${REVIT_BALLET_SIGN_PFX:-}" ]; then sign_file "$INSTALLER_DIR/bin/Release/installer.exe"; fi

echo ""
echo "Done. Installer at: $INSTALLER_DIR/bin/Release/installer.exe"
