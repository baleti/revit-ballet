#!/bin/bash
# One-time, on the VM HOST (not the guest, which stays offline): download the virtio-win ISO and extract the Windows 11 balloon driver
# into $CI_BALLOON_DIR (default ~/virtio-win/balloon/Balloon/w11/amd64: balloon.inf/.cat/.sys + blnsvr.exe). ci.sh then copies it into each
# CI guest over vsock. Needs curl and 7z. Source: Fedora's virtio-win (Red Hat, WHQL-signed). See docs/vm-testing.md, "Memory ballooning".
set -euo pipefail
ISO_URL=${ISO_URL:-https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/stable-virtio/virtio-win.iso}
BASE=${BASE:-$HOME/virtio-win}
mkdir -p "$BASE"; cd "$BASE"
[ -f virtio-win.iso ] || curl -fL -o virtio-win.iso "$ISO_URL"
rm -rf balloon; 7z x -y -oballoon virtio-win.iso "Balloon/w11/amd64" '-x!*.pdb' >/dev/null
ls -l balloon/Balloon/w11/amd64
sha256sum virtio-win.iso
