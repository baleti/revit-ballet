#!/bin/bash
# End-to-end CI against the offline Windows/Revit VM.
#
#   build installer (Build.sh)  ->  fresh qcow2 overlay on the read-only baseline  ->  boot  ->  ssh over vsock
#   ->  trust+sign add-in (throwaway cert)  ->  run the installer silently  ->  per Revit year: start Revit,
#   wait for the revit-ballet server, run tests/vm/checks/*.cs through /roslyn  ->  tear down, delete overlay.
#
# Runs on the dev machine. It drives the VM host over ssh; see docs/vm-testing.md and tests/vm/ci.env.example.
#
# usage: tests/vm/ci.sh [--years "2026 2024"] [--all] [--no-build] [--keep] [--gpu] [--installer FILE]
#   --years   Revit years to test (default "2026 2024")        --all       2020..2026
#   --no-build  reuse installer/bin/Release/installer.exe      --keep      keep the overlay disk and VM host files
#   --gpu     pass the GPU through (same hardware as the baseline; stop the interactive VM first). Default: no GPU,
#             16G/6 vCPU, so the CI VM can run next to the interactive one. Licence checks may differ without the GPU.
# Exit status = number of failed checks (0 = all green). Evidence for failures lands in tests/vm/artifacts/<run-id>/.
set -uo pipefail
HERE=$(cd "$(dirname "$0")" && pwd)
REPO=$(cd "$HERE/../.." && pwd)
[ -f "$HERE/ci.env" ] && . "$HERE/ci.env"
CI_VM_DIR=${CI_VM_DIR:-/srv/vm/win}
CI_BASELINE=${CI_BASELINE:-baseline-before-revit-ballet.qcow2}
CI_GUEST_KEY=${CI_GUEST_KEY:-'~/.ssh/WinDev2407Eval-revit'}
CI_CID=${CI_CID:-4}
YEARS="2026 2024"; BUILD=1; KEEP=0; GPU=0
INSTALLER="$REPO/installer/bin/Release/installer.exe"
while [ $# -gt 0 ]; do
  case $1 in
    --years) YEARS=$2; shift;;
    --all) YEARS="2020 2021 2022 2023 2024 2025 2026";;
    --no-build) BUILD=0;;
    --keep) KEEP=1;;
    --gpu) GPU=1;;
    --installer) INSTALLER=$2; shift;;
    -h|--help) sed -n '2,/^set -uo/p' "$0" | sed '$d;s/^# \{0,1\}//'; exit 0;;
    *) echo "unknown option $1" >&2; exit 64;;
  esac; shift
done

: "${CI_HOST2:?set CI_HOST2 in tests/vm/ci.env (see ci.env.example)}"
ID=$(date +%Y%m%d-%H%M%S)
ART="$HERE/artifacts/$ID"; mkdir -p "$ART"
RD="$CI_VM_DIR/runs/$ID"                     # per-run directory on the VM host
SSHO=(-o BatchMode=yes -o ConnectTimeout=15); [ -n "${CI_HOST2_KEY:-}" ] && SSHO+=(-i "$CI_HOST2_KEY")
h2()  { ssh "${SSHO[@]}" "$CI_HOST2" "$@"; }
gps() { h2 "$RD/gssh powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/$*"; }   # run a staged guest script
log() { printf '%s  %s\n' "$(date +%H:%M:%S)" "$*" | tee -a "$ART/run.log"; }
FAILS=0; : > "$ART/results.txt"
record() { echo "$1" | tee -a "$ART/results.txt"; case $1 in FAIL*) FAILS=$((FAILS+1));; esac; }

cleanup() {
  log "teardown"
  h2 "echo '{\"execute\":\"qmp_capabilities\"}{\"execute\":\"quit\"}' | socat - UNIX-CONNECT:/tmp/qmp-ci-$ID.sock >/dev/null 2>&1; sleep 2; rm -f /tmp/ci-$ID.pid" >/dev/null 2>&1
  if [ "$KEEP" = 1 ]; then log "kept $RD on the VM host (boot it with IMG=runs/$ID/disk.qcow2)"; else h2 "rm -rf $RD" >/dev/null 2>&1; fi
}

# ---- 1. build -------------------------------------------------------------------------------------------------
if [ "$BUILD" = 1 ]; then
  log "building (Build.sh)"; ( cd "$REPO" && bash Build.sh ) > "$ART/build.log" 2>&1 || { tail -20 "$ART/build.log"; log "BUILD FAILED"; exit 1; }
fi
[ -f "$INSTALLER" ] || { log "installer not found: $INSTALLER"; exit 1; }

# ---- 2. overlay + VM ------------------------------------------------------------------------------------------
if [ "$GPU" = 1 ] && h2 "pgrep -fx 'qemu-system-x86_64 -name revit-win.*' >/dev/null"; then
  log "--gpu needs the GPU, but the interactive VM (revit-win) is running. Shut it down first."; exit 1; fi
trap cleanup EXIT
log "run $ID: years=[$YEARS] gpu=$GPU cid=$CI_CID"
h2 "mkdir -p $RD && qemu-img create -q -f qcow2 -b $CI_VM_DIR/$CI_BASELINE -F qcow2 $RD/disk.qcow2" || { log "overlay creation failed"; exit 1; }
h2 "cat > $RD/env.sh" <<EOF
CID=$CI_CID
GKEY=$CI_GUEST_KEY
EOF
h2 "cat > $RD/gssh && chmod +x $RD/gssh" <<'EOF'
#!/bin/bash
. "$(dirname "$0")/env.sh"; GKEY=${GKEY/#\~/$HOME}
exec ssh -i "$GKEY" -o ProxyCommand="socat - VSOCK-CONNECT:$CID:22" -o StrictHostKeyChecking=no -o UserKnownHostsFile=/dev/null \
  -o LogLevel=ERROR -o BatchMode=yes -o ConnectTimeout=10 "user@vsock$CID" "$@"
EOF
h2 "cat > $RD/gcp && chmod +x $RD/gcp" <<'EOF'
#!/bin/bash
# scp wrapper; an argument starting with GUEST: addresses the guest, e.g. gcp -r GUEST:C:/ci/out /tmp/x
. "$(dirname "$0")/env.sh"; GKEY=${GKEY/#\~/$HOME}
args=(); for a in "$@"; do args+=("${a/#GUEST:/user@vsock$CID:}"); done
exec scp -q -i "$GKEY" -o ProxyCommand="socat - VSOCK-CONNECT:$CID:22" -o StrictHostKeyChecking=no -o UserKnownHostsFile=/dev/null \
  -o LogLevel=ERROR -o BatchMode=yes "${args[@]}"
EOF
h2 "cat > $RD/start-vm.sh" < "$HERE/start-vm.sh"
log "booting overlay VM"
h2 "VM_DIR=$CI_VM_DIR NAME=ci-$ID CID=$CI_CID GPU=$GPU MEM=${CI_MEM:-16G} SMP=${CI_SMP:-6} IMG=runs/$ID/disk.qcow2 bash $RD/start-vm.sh" || { log "VM failed to start"; exit 1; }
up=0; for i in $(seq 1 60); do h2 "$RD/gssh echo up" 2>/dev/null | grep -q up && { up=1; break; }; sleep 5; done
[ "$up" = 1 ] || { record "FAIL boot: guest ssh not reachable after 5 min"; exit 1; }
log "guest is up"

# ---- 3. stage + install ---------------------------------------------------------------------------------------
tar -C "$HERE" -cf - guest checks | h2 "mkdir -p $RD/stage && tar -C $RD/stage -xf -"
h2 "cat > $RD/stage/installer.exe" < "$INSTALLER"
h2 "$RD/gssh \"if (Test-Path C:/ci) { Remove-Item C:/ci -Recurse -Force }; New-Item -ItemType Directory C:/ci | Out-Null\""
h2 "$RD/gcp $RD/stage/guest/*.ps1 $RD/stage/installer.exe GUEST:C:/ci/ && $RD/gcp -r $RD/stage/checks GUEST:C:/ci/" || { record "FAIL stage: copying files to the guest"; exit 1; }
Y_CSV=${YEARS// /,}
log "installing"
out=$(gps "install.ps1 -Years $Y_CSV"); rc=$?; echo "$out" > "$ART/install.txt"
if [ $rc -ne 0 ]; then record "FAIL install (exit $rc)"; echo "$out" | grep FAIL | sed 's/^/    /'; else record "PASS install"; fi
log "trust + sign"
out=$(gps "trust-and-sign.ps1 -Years $Y_CSV"); rc=$?; echo "$out" > "$ART/sign.txt"
if [ $rc -ne 0 ]; then record "FAIL sign (exit $rc)"; echo "$out" | grep FAIL | sed 's/^/    /'; else record "PASS sign"; fi

out=$(gps "revit-ini.ps1 -Years $Y_CSV $([ "$GPU" = 0 ] && echo -NoGpu)") || record "FAIL revit-ini setup: $out"

# ---- 4. per-year Revit checks ---------------------------------------------------------------------------------
verdict() {  # stdin = /roslyn JSON response; $1 = expected substring
  python3 -c '
import json,sys
exp=sys.argv[1]
try: r=json.loads(sys.stdin.read())
except Exception as e: print("FAIL bad response: %s" % e); sys.exit(1)
if not r.get("Success"): print("FAIL script error: %s %s" % (r.get("Error"), r.get("Diagnostics"))); sys.exit(1)
o=r.get("Output") or ""
print("PASS" if exp in o else "FAIL expected %r in %r" % (exp, o)); sys.exit(0 if exp in o else 1)' "$1"
}
evidence() {  # $1 = label
  gps "shot.ps1" >/dev/null 2>&1; gps "collect.ps1" >/dev/null 2>&1
  h2 "rm -rf $RD/fetch && mkdir -p $RD/fetch && $RD/gcp -r GUEST:C:/ci/out $RD/fetch/ && tar -C $RD/fetch/out -cf - ." | tar -C "$ART" -xf - 2>/dev/null
  [ -f "$ART/desktop.png" ] && mv -f "$ART/desktop.png" "$ART/desktop-$1.png"
}
for y in $YEARS; do
  log "Revit $y: starting"
  out=$(gps "revit-up.ps1 -Year $y"); rc=$?
  port=$(echo "$out" | sed -n 's/.*PORT=\([0-9]*\).*/\1/p' | head -1)
  if [ -z "$port" ]; then record "FAIL $y startup: $(echo "$out" | tail -1)"; evidence "$y-startup"; gps revit-down.ps1 >/dev/null 2>&1; continue; fi
  log "Revit $y: server on port $port"
  for c in "$HERE"/checks/*.cs; do
    n=$(basename "$c" .cs); exp=$(sed "s/{YEAR}/$y/" "$HERE/checks/$n.expect")
    resp=$(h2 "$RD/gssh powershell -NoProfile -ExecutionPolicy Bypass -File C:/ci/query.ps1 -Port $port -File C:/ci/checks/$n.cs")
    echo "$resp" > "$ART/$y-$n.json"
    v=$(printf '%s' "$resp" | verdict "$exp")
    case $v in PASS*) record "PASS $y $n";; *) record "FAIL $y $n: ${v#FAIL }";; esac
  done
  grep -q "^FAIL $y" "$ART/results.txt" && evidence "$y"
  gps revit-down.ps1 >/dev/null 2>&1
done

echo; log "results ($ART/results.txt):"; cat "$ART/results.txt"
[ "$FAILS" = 0 ] && log "ALL GREEN" || log "$FAILS FAILED"
exit "$FAILS"
