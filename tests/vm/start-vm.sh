#!/bin/bash
# Launcher for the offline Windows + multi-Revit test VM (plain qemu/KVM, no libvirt). Runs on the VM host.
# Source of truth lives in the revit-ballet repo (tests/vm/start-vm.sh); the host keeps a copy next to the images.
#
# MUST STAY OFFLINE (proprietary software): -nic none, no netdev, ever.
# Reach the guest over vsock: ssh -o ProxyCommand="socat - VSOCK-CONNECT:<CID>:22" user@vsock<CID>
# Console only via QMP screendump/sendkey (QXL kept for that); no SPICE/USB/serial/floppy/CD-ROM.
#
# Environment switches (all optional):
#   IMG=<file>     image to boot, relative to VM_DIR or absolute (default WinDev2407Eval-revit.qcow2)
#   SNAP=1         -snapshot: scratch overlay in TMPDIR, image untouched (CI uses per-run qcow2 overlays instead)
#   NAME=<name>    qemu -name, QMP socket (/tmp/qmp-NAME.sock) and pidfile (/tmp/NAME.pid)   (default revit-win)
#   CID=<n>        vsock guest CID (default 3). Use a different CID for a second, concurrent VM.
#   GPU=0          start without GTX 1060 passthrough (default 1 = pass 1d:00.0 + 1d:00.1; only one VM can own it)
#   SMP=10 MEM=32G vCPUs / RAM
#   MACHINE=q35    (default; GPU sits behind a pcie-root-port) or pc
#   CLOCKFIX=1     hpet off + PIT lost-tick delay + rtc drift slew (default on)
#   IOTHREAD=1     virtio-blk on a dedicated iothread (default on)
#   SPICE=1        also serve the console over SPICE on a host-only unix socket /tmp/spice-NAME.sock (no guest network involved: the VM stays offline).
#                  SPICE_MC=1 allows several viewers at once (SPICE_DEBUG_ALLOW_MC: experimental, QEMU calls it crashy; viewers share one console session)
#   TOPO=0         1 = guest sees SMT pairs and pin.sh pins vCPU pairs (tested, no gain, default off)
set -e
VM_DIR=${VM_DIR:-$(cd "$(dirname "$0")" && pwd)}
cd "$VM_DIR"
export TMPDIR=${TMPDIR:-/srv/vm/tmp}
NAME=${NAME:-revit-win}
CID=${CID:-3}
N=${SMP:-10}
SNAPFLAG=""; [ "${SNAP:-0}" = 1 ] && SNAPFLAG="-snapshot"
MACHINE=${MACHINE:-q35}
MACHOPT=$MACHINE; RTC="base=localtime"; EXTRA=""
if [ "${CLOCKFIX:-1}" = 1 ]; then MACHOPT="$MACHINE,hpet=off"; RTC="base=localtime,driftfix=slew"; EXTRA="-global kvm-pit.lost_tick_policy=delay"; fi
CPUX=""; SMPARG="$N"
if [ "${TOPO:-0}" = 1 ]; then CPUX=",topoext=on"; SMPARG="$N,sockets=1,cores=$((N/2)),threads=2"; fi
IMGF=${IMG:-WinDev2407Eval-revit.qcow2}
if [ "${IOTHREAD:-1}" = 1 ]; then
  DISK="-object iothread,id=io0 -drive file=$IMGF,if=none,id=d0,format=qcow2,cache=none,aio=native,discard=unmap -device virtio-blk-pci,drive=d0,iothread=io0"
else
  DISK="-drive file=$IMGF,if=virtio,cache=none,aio=native,discard=unmap"
fi
GPUFLAGS=""
if [ "${GPU:-1}" = 1 ]; then
  if [ "$MACHINE" = q35 ]; then
    GPUFLAGS="-device pcie-root-port,id=rp0,chassis=1,slot=0 -device vfio-pci,host=0000:1d:00.0,bus=rp0,addr=0.0,multifunction=on -device vfio-pci,host=0000:1d:00.1,bus=rp0,addr=0.1"
  else
    GPUFLAGS="-device vfio-pci,host=0000:1d:00.0,multifunction=on -device vfio-pci,host=0000:1d:00.1"
  fi
fi
SPICEFLAGS=""
if [ "${SPICE:-0}" = 1 ]; then
  SPICEFLAGS="-spice unix=on,addr=/tmp/spice-$NAME.sock,disable-ticketing=on,gl=off"
  [ "${SPICE_MC:-0}" = 1 ] && export SPICE_DEBUG_ALLOW_MC=1
fi
qemu-system-x86_64 -name "$NAME" -machine $MACHOPT -nodefaults -display none -enable-kvm \
  -cpu ${CPU:-host,hv_vendor_id=nvidiagpu,hv_relaxed,hv_vapic,hv_spinlocks=0x1fff,hv_vpindex,hv_runtime,hv_time,hv_synic,hv_stimer,hv_tlbflush,hv_ipi}$CPUX \
  -smp $SMPARG -m ${MEM:-32G} -nic none -rtc $RTC $EXTRA -global PIIX4_PM.disable_s3=1 -global ICH9-LPC.disable_s3=1 $SNAPFLAG \
  $DISK -bios /usr/share/ovmf/OVMF.fd \
  -device qxl-vga,vgamem_mb=32 -device vhost-vsock-pci,guest-cid=$CID \
  $GPUFLAGS $SPICEFLAGS \
  -qmp unix:/tmp/qmp-$NAME.sock,server=on,wait=off \
  -daemonize -pidfile /tmp/$NAME.pid
if [ "${TOPO:-0}" = 1 ]; then sleep 1; "$VM_DIR/pin.sh" $N; fi
