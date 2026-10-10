#!/usr/bin/env bash
# Quick tier: build ONE Revit year, hot-reload it into a warm VM that is already running Revit (ci.sh --hold), run the checks that need Revit.
# No Revit restart, no VM boot. Typical time: build + a few seconds of reload + the trimmed battery.
#   tests/vm/quick.sh [--year 2026] [--run <run id>] [--full] [--no-unit] [--no-build] [--filter REGEX]
# First time: tests/vm/ci.sh --hold, then start Revit with  ssh <host2> <run>/gssh 'powershell -File C:/ci/revit-up.ps1 -Year 2026'.
# The shell part of the agent (commands/AgentShell.cs) and the installer are NOT reloadable: after changing those, reinstall and restart Revit.
set -euo pipefail
HERE=$(cd "$(dirname "$0")" && pwd); REPO=$(cd "$HERE/../.." && pwd)
. "$HERE/ci.env"
YEAR=2026; RUN=""; FULL=0; UNIT=1; BUILD=1
FILTER='^(info|job-value|compile-error|runtime-error|revit-context|timeout|dialog-win32|reload|value-revit|ui-state)'
while [ $# -gt 0 ]; do case $1 in
  --year) YEAR=$2; shift;; --run) RUN=$2; shift;; --full) FULL=1;; --no-unit) UNIT=0;; --no-build) BUILD=0;; --filter) FILTER=$2; shift;;
  *) echo "unknown option $1"; exit 2;; esac; shift; done
[ "$FULL" = 1 ] && FILTER=""
SSHO=(-o BatchMode=yes -o ConnectTimeout=15); [ -n "${CI_HOST2_KEY:-}" ] && SSHO+=(-i "$CI_HOST2_KEY")
h2() { ssh "${SSHO[@]}" "$CI_HOST2" "$@"; }
t0=$(date +%s); step() { printf '%3ss  %s\n' $(( $(date +%s) - t0 )) "$*"; }

if [ -z "$RUN" ]; then RUN=$(h2 "ls -t $CI_VM_DIR/runs 2>/dev/null | head -1"); fi
RD=$CI_VM_DIR/runs/$RUN
h2 "$RD/gssh echo up" 2>/dev/null | grep -q up || { echo "no warm VM at $RD (start one: tests/vm/ci.sh --hold)"; exit 1; }
PORT=23717

if [ "$UNIT" = 1 ]; then
  ( cd "$REPO" && dotnet test tests/host -c Release --nologo -v q 2>&1 | grep -E 'Passed!|Failed!|error' | sed 's/ - RevitBallet.*//' ) || true
  step "unit tests done"
fi
if [ "$BUILD" = 1 ]; then
  ( cd "$REPO" && dotnet build commands -c Release -p:RevitYear=$YEAR 2>&1 | grep -E ' error |Error\(s\)' ) | grep -v ' 0 Error' && { echo "BUILD FAILED"; exit 1; } || true
  step "built revit-ballet.dll ($YEAR)"
fi
DLL="$REPO/commands/bin/$YEAR/revit-ballet.dll"; [ -f "$DLL" ] || { echo "missing $DLL"; exit 1; }
cat "$DLL" | h2 "cat > $RD/rb.dll"; cat "$REPO/commands/bin/$YEAR/RevitBallet.Core.dll" | h2 "cat > $RD/RevitBallet.Core.dll"
cat "$HERE/guest/v1-tests.ps1" | h2 "cat > $RD/v1-tests.ps1"; cat "$HERE/guest/datagrid-dsl.cs" | h2 "cat > $RD/datagrid-dsl.cs"; cat "$HERE/guest/reload.ps1" | h2 "cat > $RD/reload.ps1"
h2 "$RD/gcp $RD/rb.dll $RD/v1-tests.ps1 $RD/datagrid-dsl.cs $RD/reload.ps1 GUEST:C:/ci/" || exit 1
h2 "$RD/gcp $RD/RevitBallet.Core.dll GUEST:C:/ci/" || exit 1
h2 "$RD/gssh \"New-Item -ItemType Directory -Force \\\$env:APPDATA/revit-ballet/hot-reload/$YEAR | Out-Null; Copy-Item C:/ci/rb.dll \\\$env:APPDATA/revit-ballet/hot-reload/$YEAR/revit-ballet.dll -Force; Copy-Item C:/ci/RevitBallet.Core.dll \\\$env:APPDATA/revit-ballet/hot-reload/$YEAR/RevitBallet.Core.dll -Force\""
step "pushed"
h2 "$RD/gssh powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/reload.ps1 -Port $PORT" || { echo "RELOAD FAILED (is Revit $YEAR running with this build's shell? see runtime\\server.log)"; exit 1; }
step "reloaded"
out=$(h2 "$RD/gssh powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/v1-tests.ps1 -Port $PORT -Year $YEAR -Filter '\"$FILTER\"'" 2>&1) || true
echo "$out" | grep -E '^(PASS|FAIL)' | sed 's/^PASS/  pass/'
p=$(echo "$out" | grep -c '^PASS' || true); f=$(echo "$out" | grep -c '^FAIL' || true)
step "checks: $p passed, $f failed"
[ "$f" = 0 ] && [ "$p" -gt 0 ]
