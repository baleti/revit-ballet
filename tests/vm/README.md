# tests/vm - end-to-end CI against the offline Windows/Revit VM

```
cp ci.env.example ci.env      # describe your VM host once (git-ignored)
./ci.sh                       # build, boot a fresh overlay of the baseline, install, test Revit 2026 + 2024
./ci.sh --all --no-build      # all years 2020-2026, reuse the last build
./ci.sh --keep                # keep the overlay disk for debugging
```

| Path | What |
|---|---|
| `ci.sh` | orchestrator (runs on the dev machine, drives the VM host over ssh) |
| `start-vm.sh` | qemu launcher, runs on the VM host (offline, GPU passthrough optional) |
| `balloon-autoscale.py`, `fetch-balloon-driver.sh` | memory ballooning so several VMs share the host's RAM (docs/vm-testing.md) |
| `guest/*.ps1` | PowerShell helpers copied into the guest at `C:\ci` |
| `checks/` | Roslyn scripts + expected output, run inside every Revit year |
| `guest-setup/` | Windows-side pieces of the baseline image (vsock relay) |
| `artifacts/` | per-run evidence (screenshots, logs, JSON), git-ignored |

Everything is explained in [docs/vm-testing.md](../../docs/vm-testing.md).
