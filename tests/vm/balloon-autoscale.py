#!/usr/bin/env python3
"""
balloon-autoscale.py

Runs on the HOST (not the guest). Monitors a VM's guest-reported free memory
via virtio-balloon stats (no guest agent required) and adjusts the balloon
target to keep guest free memory near a target floor.

Windows guests: needs the virtio-win balloon driver and the blnsvr.exe service (tests/vm/guest/balloon.ps1); it reports the same
stat-available-memory field as Linux.   Per-VM watcher: add --exit-when-gone.

Usage:
    balloon-autoscale.py --socket /tmp/qmp-claude3.sock \
        --min-free-mb 1024 --headroom-mb 512 \
        --min-size-mb 1024 --max-size-mb 10240 \
        --poll-interval 5 --balloon-path /machine/peripheral/balloon0
"""
import argparse
import json
import os
import socket
import time
import sys

def wait_for_socket(path, timeout=30, poll_interval=0.2, quiet=False):
    """Block until the QMP unix socket exists (and is a socket), or timeout."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if os.path.exists(path) and stat_is_socket(path):
            return True
        time.sleep(poll_interval)
    if not quiet:
        print(f"[fatal] timed out waiting for socket {path} after {timeout}s", file=sys.stderr)
    return False

def stat_is_socket(path):
    import stat as stat_module
    try:
        return stat_module.S_ISSOCK(os.stat(path).st_mode)
    except OSError:
        return False

def qmp_connect(path):
    s = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
    s.connect(path)
    f = s.makefile("rwb")
    greeting = json.loads(f.readline())
    send(f, {"execute": "qmp_capabilities"})
    recv(f)  # ack
    return s, f

def send(f, obj):
    f.write((json.dumps(obj) + "\n").encode())
    f.flush()

def recv(f):
    while True:
        line = f.readline()
        if not line:
            raise ConnectionError("QMP socket closed")
        obj = json.loads(line)
        # QEMU pushes unsolicited events (e.g. BALLOON_CHANGE) on the same
        # socket. Skip them and keep waiting for the actual command reply.
        if "event" in obj:
            continue
        return obj

def qom_set(f, path, prop, value):
    send(f, {"execute": "qom-set", "arguments": {"path": path, "property": prop, "value": value}})
    return recv(f)

def qom_get(f, path, prop):
    send(f, {"execute": "qom-get", "arguments": {"path": path, "property": prop}})
    return recv(f)

def query_balloon(f):
    send(f, {"execute": "query-balloon"})
    return recv(f)

def set_balloon(f, target_bytes):
    send(f, {"execute": "balloon", "arguments": {"value": int(target_bytes)}})
    return recv(f)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--socket", required=True, help="path to QMP unix socket")
    ap.add_argument("--balloon-path", default="/machine/peripheral/balloon0",
                     help="QOM path of the balloon device (matches -device id=)")
    ap.add_argument("--min-free-mb", type=int, default=1024,
                     help="target minimum guest free memory (MB)")
    ap.add_argument("--headroom-mb", type=int, default=512,
                     help="extra buffer above min-free before we deflate again (MB)")
    ap.add_argument("--min-size-mb", type=int, default=1024,
                     help="never shrink guest below this total size (MB)")
    ap.add_argument("--max-size-mb", type=int, default=10240,
                     help="never grow guest above this total size (MB), i.e. original -object size")
    ap.add_argument("--poll-interval", type=int, default=5, help="seconds between checks")
    ap.add_argument("--step-mb", type=int, default=512, help="max adjustment step per cycle (MB)")
    ap.add_argument("--socket-wait-timeout", type=int, default=30,
                     help="seconds to wait for the QMP socket to appear before giving up")
    ap.add_argument("--exit-when-gone", action="store_true",
                     help="exit 0 once the QMP socket is gone / refuses connections (the VM quit); for per-VM watchers started with the VM")
    ap.add_argument("--quiet", action="store_true", help="suppress all log output")
    args = ap.parse_args()

    MB = 1024 * 1024

    if not wait_for_socket(args.socket, timeout=args.socket_wait_timeout, quiet=args.quiet):
        sys.exit(1)

    sock, f = qmp_connect(args.socket)
    if not args.quiet:
        print(f"[connected] {args.socket}", file=sys.stderr)

    # Enable periodic guest stats reporting over the balloon stats virtqueue
    r = qom_set(f, args.balloon_path, "guest-stats-polling-interval", args.poll_interval)
    if "error" in r:
        if not args.quiet:
            print(f"[fatal] could not enable guest-stats-polling-interval: {r}", file=sys.stderr)
            print("Check --balloon-path matches your device id (qom-list under /machine/peripheral/ to find it)", file=sys.stderr)
        sys.exit(1)

    last_touch_date = None

    while True:
        today = time.strftime("%Y-%m-%d")
        if today != last_touch_date:
            try:
                os.utime(args.socket, None)  # bump atime, dodges /tmp tmpfiles reaping
            except OSError:
                pass
            last_touch_date = today

        try:
            stats_r = qom_get(f, args.balloon_path, "guest-stats")
            balloon_r = query_balloon(f)
        except (ConnectionError, BrokenPipeError, OSError) as e:
            if not args.quiet:
                print(f"[warn] QMP connection lost ({e}), reconnecting...", file=sys.stderr)
            time.sleep(2)
            if args.exit_when_gone and not stat_is_socket(args.socket):
                sys.exit(0)
            if not wait_for_socket(args.socket, timeout=args.socket_wait_timeout, quiet=args.quiet):
                sys.exit(1)
            try:
                sock, f = qmp_connect(args.socket)
            except OSError:
                if args.exit_when_gone:
                    sys.exit(0)
                raise
            continue

        if "error" in stats_r:
            if not args.quiet:
                print(f"[warn] guest-stats not ready yet: {stats_r['error']}", file=sys.stderr)
            time.sleep(args.poll_interval)
            continue

        stats = stats_r["return"].get("stats", {})
        last_update = stats_r["return"].get("last-update", 0)
        # stat-free-memory excludes page cache, which Linux keeps low on
        # purpose. Use stat-available-memory (free + reclaimable cache) instead
        free_bytes = stats.get("stat-available-memory")
        if free_bytes is None:
            free_bytes = stats.get("stat-free-memory")
        current_actual = balloon_r["return"]["actual"]

        if free_bytes is None:
            if not args.quiet:
                print("[warn] guest hasn't reported stats yet (guest kernel may not support F_STATS_VQ)", file=sys.stderr)
            time.sleep(args.poll_interval)
            continue

        # last-update == 0 means the stats virtqueue hasn't produced a real
        # sample yet (fresh device or just-enabled polling). Also guard
        # against implausible values (free memory can't exceed the max
        # guest size we'd ever configure) which show up as a sentinel
        # before the first genuine sample arrives.
        if last_update == 0 or free_bytes > args.max_size_mb * MB:
            if not args.quiet:
                print(f"[warn] guest-stats not fresh yet (last-update={last_update}, "
                      f"free_bytes={free_bytes}), skipping this cycle", file=sys.stderr)
            time.sleep(args.poll_interval)
            continue

        free_mb = free_bytes / MB
        current_mb = current_actual / MB
        step = args.step_mb * MB

        new_target = None
        if free_mb < args.min_free_mb:
            # guest is running low -> give it more RAM
            new_target = min(current_actual + step, args.max_size_mb * MB)
        elif free_mb > (args.min_free_mb + args.headroom_mb):
            # guest has plenty spare -> reclaim some for host
            new_target = max(current_actual - step, args.min_size_mb * MB)

        if new_target is not None and new_target != current_actual:
            set_balloon(f, new_target)
            if not args.quiet:
                print(f"[adjust] free={free_mb:.0f}MB current={current_mb:.0f}MB -> target={new_target/MB:.0f}MB",
                      file=sys.stderr)
        else:
            if not args.quiet:
                print(f"[ok] free={free_mb:.0f}MB current={current_mb:.0f}MB (no change)", file=sys.stderr)

        time.sleep(args.poll_interval)

if __name__ == "__main__":
    main()
