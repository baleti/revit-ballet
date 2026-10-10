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
    if [ "${errors:-0}" -gt 0 ]; then
        FAILED+=("$year")
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

# Script engines (separate components, shipped in engines/<name>/): one build per Revit year each.
ENGINES=(roslyn ironpython2 cpython3)
bash "$REPO_ROOT/agents/cpython3/fetch-python.sh"   # embeddable CPython for the cpython3 engine (no-op when present)
for engine in "${ENGINES[@]}"; do
    echo ""
    echo "Building engine '$engine' for Revit 2026..2020..."
    for year in $(seq 2026 -1 2020); do
        if dotnet build "$REPO_ROOT/agents/$engine" -c Release -p:RevitYear=$year 2>&1 | grep -q -E '^\s*0 Error\(s\)'; then
            echo "Revit $year: OK"
        else
            echo "Revit $year: FAILED"; FAILED+=("$engine-$year")
        fi
    done
done
if [ ${#FAILED[@]} -gt 0 ]; then echo "Build failed for: ${FAILED[*]}"; exit 1; fi

if [ -n "${REVIT_BALLET_SIGN_PFX:-}" ]; then
    echo ""
    echo "Signing revit-ballet.dll for each Revit year..."
    for year in $(seq 2026 -1 2020); do
        sign_file "$COMMANDS_DIR/bin/$year/revit-ballet.dll"
        for engine in "${ENGINES[@]}"; do sign_file "$REPO_ROOT/agents/$engine/bin/$year/revit-ballet-engine-$engine.dll"; done
    done
fi

echo ""
echo "Building installer..."
# The installer project embeds installer/bin/resources/* picked up when the project is evaluated, i.e. BEFORE its own prepare step runs;
# generate them first so a changed set of bundled DLLs does not fail the first build with a missing-resource error.
(cd "$INSTALLER_DIR" && python3 _installer-merge-shortcuts.py >/dev/null && python3 _installer-deduplicate.py >/dev/null)
dotnet build "$INSTALLER_DIR" -c Release

if [ -n "${REVIT_BALLET_SIGN_PFX:-}" ]; then sign_file "$INSTALLER_DIR/bin/Release/installer.exe"; fi

echo ""
echo "Done. Installer at: $INSTALLER_DIR/bin/Release/installer.exe"
