#!/usr/bin/env bash
# The Windows remote command picker, end to end in a VM that already runs Revit 2026 (ci.sh --hold --years 2026, then revit-up.ps1):
#   installer opt-in states -> the real tray program started like the Run key would -> real Win+` hotkey, typed queries, Enter runs a command
#   -> the agent's job list proves the command ran in Revit -> opt-out and uninstall remove everything.
#   tests/vm/remote-picker.sh [--run <run id>] [--no-build] [--no-uninstall]
#   The last step uninstalls Revit Ballet from the VM (the check that the uninstaller removes the picker); --no-uninstall skips it so the VM stays usable.
set -uo pipefail
HERE=$(cd "$(dirname "$0")" && pwd); REPO=$(cd "$HERE/../.." && pwd)
. "$HERE/ci.env"
RUN=""; BUILD=1; UNINSTALL=1
while [ $# -gt 0 ]; do case $1 in --run) RUN=$2; shift;; --no-build) BUILD=0;; --no-uninstall) UNINSTALL=0;; *) echo "unknown option $1"; exit 2;; esac; shift; done
SSHO=(-o BatchMode=yes -o ConnectTimeout=15); [ -n "${CI_HOST2_KEY:-}" ] && SSHO+=(-i "$CI_HOST2_KEY")
h2() { ssh "${SSHO[@]}" "$CI_HOST2" "$@"; }
[ -n "$RUN" ] || RUN=$(h2 "ls -t $CI_VM_DIR/runs | head -1")
RD=$CI_VM_DIR/runs/$RUN; FAILS=0
g() { h2 "$RD/gssh $*"; }
step() { printf '\n== %s\n' "$*"; }
count() { grep -c "^FAIL" <<<"$1" || true; }
report() { printf '%s\n' "$1"; FAILS=$((FAILS + $(grep -c "^FAIL" <<<"$1"))); }

if [ "$BUILD" = 1 ]; then
  step "build the picker and the installer"
  dotnet build "$REPO/remote/RevitBalletRemote" -c Release --nologo -v q 2>&1 | grep -E " error |Error\(s\)"
  ( cd "$REPO/installer" && python3 _installer-merge-shortcuts.py >/dev/null && python3 _installer-deduplicate.py >/dev/null ) && dotnet build "$REPO/installer" -c Release --nologo -v q 2>&1 | grep -E " error |Error\(s\)"
fi
step "push installer and tests"
cat "$REPO/installer/bin/Release/installer.exe" | h2 "cat > $RD/installer.exe"
for f in remote-installer.ps1 remote-picker.ps1 remote-picker-start.ps1 remote-prepare.ps1 q-jobs.ps1; do cat "$HERE/guest/$f" | h2 "cat > $RD/$f"; done
h2 "$RD/gcp $RD/installer.exe $RD/remote-installer.ps1 $RD/remote-picker.ps1 $RD/remote-picker-start.ps1 $RD/remote-prepare.ps1 $RD/q-jobs.ps1 GUEST:C:/ci/" || exit 1
g powershell -NoProfile -Command "Get-Process RevitBalletRemote -ErrorAction SilentlyContinue | Stop-Process -Force" >/dev/null 2>&1 || true

step "prepare: add-in installed, Revit 2026 running"
g 'powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/remote-prepare.ps1' 2>&1 | tail -4

step "installer: default does not add the picker; --remote-picker on installs it"
report "$(g 'powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/remote-installer.ps1 -Step default' 2>&1)"
report "$(g 'powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/remote-installer.ps1 -Step on' 2>&1)"
report "$(g 'powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/remote-installer.ps1 -Step keep' 2>&1)"

step "the program, started the way the Run key starts it, in the interactive session"
g 'powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/remote-picker-start.ps1' >/dev/null
out=""; for i in $(seq 1 60); do sleep 2
  out=$(g "Get-Content C:/ci/remote-picker.txt -ErrorAction SilentlyContinue" 2>/dev/null)
  grep -q "^DONE" <<<"$out" && break; grep -q "^FAIL" <<<"$out" && break
done
report "$out" | grep -v "^DONE"

step "the agent confirms the command Enter picked ran inside Revit"
jobs=$(g 'powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/q-jobs.ps1' 2>/dev/null)
echo "$jobs"
if grep -q "command .*No active document" <<<"$jobs"; then echo "PASS enter-ran-the-command-in-revit"; else echo "FAIL enter-ran-the-command-in-revit: no 'command' job with CloseWorksets' own message in the agent's job list"; FAILS=$((FAILS+1)); fi

step "installer: opt-out, then opt-in + uninstall"
report "$(g 'powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/remote-installer.ps1 -Step off' 2>&1)"
[ "$UNINSTALL" = 1 ] && report "$(g 'powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/remote-installer.ps1 -Step uninstall' 2>&1)"

printf '\n%s\n' "remote picker: $FAILS failure(s)"
exit "$FAILS"
