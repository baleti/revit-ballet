# VM testing and CI

Revit Ballet is tested end to end against a real Windows machine with real Revit installs. That machine is a
qemu/KVM virtual machine on a Linux host. This document describes the whole setup so it can be rebuilt, and how
the CI runner in `tests/vm/` uses it.

Supported Revit versions: **2020 to 2026** (2017 to 2019 were dropped; they could not be licensed in the VM and are too old).

## Why a VM, and why offline

- Revit and the Windows image are proprietary and licensed to this setup. **The VM never has a network
  device** (`-nic none`, no netdev ever). Everything goes through vsock, a host-to-guest socket that needs no network.
- Revit needs a real interactive desktop and, for sensible performance, a GPU. The VM gets a GTX 1060 through
  `vfio-pci` passthrough. Hardware acceleration is enabled in every `Revit.ini` (`[Graphics] UseGraphicsHardware=1`);
  with it off, 3D navigation lagged badly.
- Licences (FlexNet) are sensitive to hardware changes. Keep the hardware definition of the baseline stable.

## Layout

```
dev machine (builds, drives CI)
   |  ssh
VM host  (Linux, qemu/KVM, vfio GPU)         images live in one directory, here /srv/vm/win
   |  vsock CID 3 (interactive) / CID 4 (CI)
guest    (Windows 11, Revit 2017-2026 vanilla, OpenSSH, RDP)
```

| File on the VM host | Role |
|---|---|
| `baseline-before-revit-ballet.qcow2` | Golden image: Windows + vanilla Revit installs, nothing of revit-ballet. **Read-only** (see below). |
| `WinDev2407Eval-revit.qcow2` | The interactive dev/test VM. Can be written to freely; revit-ballet builds are deployed here by hand. |
| `runs/<run-id>/disk.qcow2` | Per-CI-run overlay whose backing file is the baseline. Deleted after the run. |
| `start-vm.sh` | Launcher (`tests/vm/start-vm.sh` is the source of truth). |

### Protecting the baseline

The baseline is the reference that CI trusts, so it is protected in layers:

1. `chmod a-w` (read-only for everyone).
2. `chattr +i` on the btrfs file: even root cannot modify, rename or delete it until someone runs `chattr -i`.
   (btrfs read-only properties apply to subvolumes, not single files, so `chattr +i` is the file-level tool.)
3. qemu opens backing files read-only, so overlays can never write into it.
4. A second copy is kept on another machine. Record `sha256sum` of the file when it is taken.

To update the baseline deliberately: boot the live image, make the change, shut down, `chattr -i`, copy (btrfs
`cp --reflink=always` is instant), re-protect. Never boot the baseline file directly.

## The launcher (`tests/vm/start-vm.sh`)

Plain qemu, no libvirt. Key points (details are in the script header):

- `q35` machine, OVMF firmware, `-nodefaults -display none`, QXL kept only for a console that is read via QMP
  screendump. With the GPU passed through, the real desktop is on the GPU output, so a QMP screendump is blank;
  CI takes desktop screenshots from inside the guest instead (`tests/vm/guest/shot.ps1`).
- Hyper-V enlightenments (`hv_*`), HPET off + PIT lost-tick delay + RTC drift slew (fixes idle CPU and clock),
  a dedicated iothread for virtio-blk. Tested and kept: those. Tested and rejected: SMT topology + pinning.
- `IMG=` picks the image, `NAME=`/`CID=` allow a second concurrent VM, `GPU=0` starts without passthrough,
  `SNAP=1` uses a throwaway `-snapshot` (CI uses overlays instead because a failed run's disk can be kept and booted).
- Only one VM can own the GPU. The host needs `vfio-pci` bound to the GPU and its audio function, IOMMU on,
  and `vhost_vsock` loaded.

## Reaching the guest

- **ssh**: OpenSSH server in the guest, reached over vsock: from the VM host
  `ssh -o ProxyCommand="socat - VSOCK-CONNECT:<CID>:22" user@vsock<CID>`. The default shell is PowerShell.
  Windows has no native vsock-to-TCP bridge, so a small relay runs at logon
  (`tests/vm/guest-setup/vsock-relay.py`: listens on a vsock port, forwards to `127.0.0.1`; one instance for
  port 22 and one for 3389 with the constants edited).
- **RDP** for interactive work: relay on vsock 3389, then a TCP forwarder on the host and an ssh tunnel to the
  dev machine.
- GUI programs cannot be started from an ssh session (no desktop). `tests/vm/guest/it.ps1` starts anything in the
  logged-in desktop through a one-shot scheduled task with an interactive principal. The CI uses it for Revit and
  for screenshots.

## What the baseline contains (Windows side)

Nothing here is revit-ballet specific. To rebuild the baseline from a fresh Windows 11 dev image:

- Revit 2017 to 2026 vanilla installs, licensed, hardware acceleration on, first-run dialogs dismissed.
- Auto-logon of the `User` account so an interactive session always exists.
- OpenSSH server (key auth) and the two vsock relays above in the user's Startup folder.
- A second loopback network adapter (Microsoft KM-TEST) so SMB works locally and PsExec `-i -s` works.
- Sysinternals in `C:\Tools\Sysinternals`; clipboard redirection for RDP.
- No guest agents that need a network or a balloon (SPICE agent, QEMU guest agent and virtio balloon were removed;
  balloon cannot work with GPU passthrough).

## CI runner (`tests/vm/ci.sh`)

```
tests/vm/ci.sh [--hold] [--years "2026 2024"] [--all] [--no-build] [--keep] [--gpu] [--installer FILE]
```

One run does, in order:

1. **Build** with `Build.sh` (all Revit years and the installer; plain Linux, `dotnet`, Revit API from NuGet).
2. **Overlay**: `qemu-img create -b <baseline> -F qcow2 runs/<id>/disk.qcow2`. Every run starts from the pristine baseline.
3. **Boot** the overlay (default: no GPU, 16 GB, 6 vCPU, CID 4, so it can run next to the interactive VM) and wait for ssh.
4. **Stage** the guest scripts, checks and installer into `C:\ci`.
5. **Install**: `installer.exe /quiet --verbosity 2`; verifies exit code, `.addin` and `revit-ballet.dll` per year, uninstaller.
6. **Trust + sign**: a throwaway self-signed certificate is created inside the guest, trusted machine-wide
   (LocalMachine Root and Trusted Publishers), and the installed `revit-ballet.dll` files are signed. This stands in for
   a real code-signing certificate so Revit does not show its "unsigned add-in" prompt. The certificate never leaves the VM.
7. **Revit.ini fixes** (`guest/revit-ini.ps1`) so no modal dialog can appear: an empty `FamilyTemplatePath`
   (Revit 2020 in the baseline has no template content) gets an existing directory, and in no-GPU mode
   `[Graphics] UseGraphicsHardware=0` (otherwise a "Hardware Acceleration disabled" dialog blocks every script).
8. **Per Revit year**: start Revit on the desktop, wait until the revit-ballet server listens (port 23717 up),
   run every `tests/vm/checks/*.cs` through `/roslyn` and compare to its `.expect`, then run the `/v1` agent API battery
   (`guest/v1-tests.ps1`: info, jobs, errors, sessions, queue order, cancel, timeouts, SSE, non-ASCII bodies, the legacy shim, and for the
   `ironpython2` and `cpython3` engines the same Python checks), then stop Revit. After all years the installer's `--engines` flag is
   exercised (`guest/engines-flag.ps1`).
9. **Evidence on failure**: desktop screenshot, `server.log`, tail of the newest Revit journals, per-check JSON,
   all under `tests/vm/artifacts/<run-id>/` (git-ignored).
10. **Tear down**: QMP `quit`, delete the overlay (`--keep` keeps it).

The exit status is the number of failed checks. A typical single-year run takes about two minutes; all seven years
take roughly fifteen.

### Setup on a new dev machine

Copy `tests/vm/ci.env.example` to `tests/vm/ci.env` (git-ignored) and fill in the VM host, ssh key, image directory
and guest key. Requirements: `ssh`, `python3`, `dotnet` on the dev machine; `qemu-img`, `socat` and qemu on the VM host.

### Writing checks

A check is `NN-name.cs` plus `NN-name.expect` in `tests/vm/checks/`. The script runs inside Revit with `UIApp`,
`UIDoc`, `Doc`; ASCII only; the response `Output` must contain the expect text (`{YEAR}` is substituted). Keep
checks cheap and deterministic; they all run in one Revit session per year. Current checks: Revit year matches,
revit-ballet assembly loaded with a plausible number of commands, application object usable.

### Known pitfalls

- **Modal dialogs freeze `/roslyn`**: scripts run on the UI thread via an external event, so any open dialog turns
  every check into a "Script execution timeout". The failure screenshot shows the dialog; fix it in the
  baseline/setup rather than clicking through it.
- **Unsigned add-in prompt**: the registry values the installer writes under `...\CodeSigning\<AddInId>` do not
  suppress the prompt in Revit 2026. Choosing "Always Load" is only persisted when Revit exits cleanly, and we did not
  find where. A signed DLL whose signer is trusted loads silently (verified on Revit 2024 and 2026).
- **`Stop-Process` on Revit** discards its settings; use a window close when settings must persist.
- **Runtime path**: server data is under `%APPDATA%\revit-ballet\runtime\` (token at `runtime\network\token`).
- **Licences** after hardware changes: if a no-GPU CI run reports licence trouble, use `--gpu` after shutting the
  interactive VM down, which reproduces the baseline's hardware exactly.

## GPU passthrough or not

Measured 2026-10-10 on Revit 2026 with the 38k-element Snowdon Towers sample, same agent benchmark on both (`tests/bench`):

| | GPU VM | no-GPU VM |
|---|---|---|
| hardware | GTX 1060 passed through, 10 vCPU | QXL display, hardware acceleration off (`UseGraphicsHardware=0`), 8 vCPU |
| agent tasks (30 task/engine pairs) | baseline | median **1.22x** slower, which is the vCPU ratio (1.25x): the GPU does not matter for compute, API loops, output or queueing |
| writes to the model (set a parameter on 2,000 elements, roll back) | 0.2 to 0.5 s | **1.4 to 1.6 s (3 to 7x)** |
| 3D view export, 1920 px | not measured (the GPU VM was busy) | 3.7 s per image |
| Revit UI | full speed | works, full UI and shaded 3D view render, display is basic 1366x768 |

- **Why writes are slower**: a transaction makes Revit redraw the open view, and without acceleration that redraw is software rendering. Tasks that
  modify the model with a graphical view open are the ones that feel the missing GPU.
- **What does not change**: starting Revit, opening a model (36 s against 30 s, mostly the CPU count), every read-only agent task, installers, signing, RDP
  (RDP renders through its own virtual display).
- **Memory ballooning**: not possible with passthrough (QEMU treats VFIO devices as incompatible with ballooning); possible without it.
- **virtio-gpu**: stock Windows support is 2D only; a 3D driver is experimental. QXL is the practical virtual display for Windows.
- **Not measured**: the GPU side of 3D export and viewport navigation. Run `tests/bench/gfx.cs` (through `job.ps1`) on both VMs when the GPU VM is free.

Recommendation: CI and agent tests do not need the GPU (they already run without it). Keep passthrough only for a VM used for interactive 3D work.

### SPICE as a console

`SPICE=1 tests/vm/start-vm.sh` serves the console over SPICE on a host-only unix socket (`/tmp/spice-NAME.sock`; the guest stays offline). Findings from the
SPICE documentation: a viewer shows the one **console session**, it does not create a second session; several viewers at once need the experimental
`SPICE_MC=1` mode (`SPICE_DEBUG_ALLOW_MC`, QEMU itself calls it crashy; all viewers share input). So SPICE helps to watch or share one VM, not to scale
interactive access across tests: automated runs need no viewer at all (vsock ssh, in-guest screenshots). Multiple VMs scale by RAM and, as measured,
mostly by CPU: seven Revits at once are slower than four. Not yet verified live: multi-viewer behaviour (a viewer is installed on the dev machine as the
`org.virt_manager.virt-viewer` flatpak).

## Signing policy

- **CI**: self-signed per-run certificate inside the VM (above). Tests the "trusted publisher" load path.
- **Releases**: sign `revit-ballet.dll` for each year and `installer.exe` with a real code-signing certificate
  from a public CA. This removes Revit's prompt for users and reduces antivirus false positives. `Build.sh` signs
  automatically when `REVIT_BALLET_SIGN_PFX` (and optionally `REVIT_BALLET_SIGN_PASS`) is set and `osslsigncode` is
  installed; the key must never be committed. (This hook is written but has not been exercised yet: `osslsigncode` is not installed on the dev machine.) GitHub Actions can hold the PFX as a secret and run the same script.

## Related

- `tests/vm/README.md`: quick reference for the runner.
- `tests/README.md`: older libvirt/virtiofs notes, superseded by this setup.
- `CLAUDE.md`: how to query a running Revit session through `/roslyn`.

## Quick tier (warm VM, one year, no Revit restart)

`tests/vm/quick.sh` is the inner loop. Start a VM once with `tests/vm/ci.sh --hold` and Revit 2026 in it (`revit-up.ps1`); after that every
run is: unit tests (about 15 s, `--no-unit` to skip), build one year, copy `revit-ballet.dll` into the guest's hot-reload folder, `POST /v1/admin/reload`,
then only the checks that need Revit (about 10 s). About 30 s end to end.

    tests/vm/quick.sh                       # 2026, unit tests + trimmed battery
    tests/vm/quick.sh --year 2024 --full    # another year, every check
    tests/vm/quick.sh --filter 'dialog'     # only checks whose name matches

Not reloadable, so reinstall and restart Revit after changing them: the shell (`commands/AgentShell.cs`, `Server.cs`), the installer, the
engine DLLs, and anything that adds a dependency DLL. The full matrix (`ci.sh`, all years, parallel) stays the gate before a push to main.
