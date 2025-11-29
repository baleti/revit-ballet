#!/bin/bash

# VM Launcher for Revit Ballet Testing
# Launches Windows VM with Revit via QEMU/KVM and establishes vsock SSH connection

set -euo pipefail

# Configuration
VM_IMAGE="WinDev2407Eval-revit.qcow2"
VM_CID=3  # vsock guest CID
VM_CPUS=8
VM_RAM="8G"
SPICE_PORT=5900
SSH_KEY="$HOME/.ssh/WinDev2407Eval-revit"
SHARED_DIR="$HOME/revit-ballet"  # Shared with guest via virtiofsd
VIRTIOFSD_SOCKET="/tmp/revit-ballet-virtiofs.sock"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

log() {
    echo -e "${GREEN}[$(date +'%Y-%m-%d %H:%M:%S')]${NC} $*"
}

error() {
    echo -e "${RED}[ERROR]${NC} $*" >&2
}

warn() {
    echo -e "${YELLOW}[WARN]${NC} $*"
}

# Check if VM is already running
check_vm_running() {
    if pgrep -f "qemu.*$VM_IMAGE" > /dev/null; then
        warn "VM appears to be already running"
        return 0
    fi
    return 1
}

# Setup shared directory
setup_shared_dir() {
    log "Setting up shared directory..."

    # Create shared directory structure if it doesn't exist
    mkdir -p "$SHARED_DIR/runtime/network"
    mkdir -p "$SHARED_DIR/runtime/SearchboxQueries"
    mkdir -p "$SHARED_DIR/runtime/LogViewChanges"
    mkdir -p "$SHARED_DIR/runtime/screenshots"
    mkdir -p "$SHARED_DIR/commands"

    log "Shared directory ready: $SHARED_DIR"
}

# Start virtiofsd daemon
start_virtiofsd() {
    log "Starting virtiofsd daemon..."

    # Clean up old socket if it exists
    rm -f "$VIRTIOFSD_SOCKET"

    # Find virtiofsd executable
    local virtiofsd_path=""
    for path in /usr/lib/qemu/virtiofsd /usr/libexec/virtiofsd /usr/bin/virtiofsd; do
        if [ -x "$path" ]; then
            virtiofsd_path="$path"
            break
        fi
    done

    if [ -z "$virtiofsd_path" ]; then
        error "virtiofsd not found. Please install QEMU with virtiofs support"
        error "On Debian/Ubuntu: apt install qemu-system-x86 virtiofsd"
        error "On Fedora/RHEL: dnf install qemu-system-x86 virtiofsd"
        return 1
    fi

    # Start virtiofsd in background
    "$virtiofsd_path" \
        --socket-path="$VIRTIOFSD_SOCKET" \
        --shared-dir="$SHARED_DIR" \
        --cache=auto \
        --sandbox=none \
        > /tmp/revit-ballet-virtiofsd.log 2>&1 &

    local vfsd_pid=$!
    echo $vfsd_pid > /tmp/revit-ballet-virtiofsd.pid

    # Wait for socket to be created
    local attempts=0
    while [ ! -S "$VIRTIOFSD_SOCKET" ] && [ $attempts -lt 10 ]; do
        sleep 0.5
        attempts=$((attempts + 1))
    done

    if [ ! -S "$VIRTIOFSD_SOCKET" ]; then
        error "virtiofsd failed to create socket"
        error "Check logs: /tmp/revit-ballet-virtiofsd.log"
        return 1
    fi

    log "virtiofsd started with PID: $vfsd_pid"
}

# Launch VM in background
launch_vm() {
    log "Launching Windows VM with Revit..."

    if [ ! -f "$VM_IMAGE" ]; then
        error "VM image not found: $VM_IMAGE"
        error "Please ensure the VM image is in the current directory"
        exit 1
    fi

    # Setup shared directory and virtiofsd
    setup_shared_dir
    start_virtiofsd || {
        warn "virtiofsd failed to start - continuing without shared filesystem"
        warn "You can still use SSH-based orchestration"
    }

    # Build QEMU command with optional virtiofs
    local qemu_cmd=(
        qemu-system-x86_64
        -enable-kvm
        -smp $VM_CPUS
        -m $VM_RAM
        -nic none
        -drive file="$VM_IMAGE",if=virtio
        -bios /usr/share/ovmf/OVMF.fd
        -usb
        -device usb-tablet
        -device virtio-serial-pci
        -device virtserialport,chardev=spicechannel0,name=com.redhat.spice.0
        -chardev spicevmc,id=spicechannel0,name=vdagent
        -spice port=$SPICE_PORT,addr=127.0.0.1,disable-ticketing=on
        -device qxl-vga,vgamem_mb=32
        -device vhost-vsock-pci,guest-cid=$VM_CID
    )

    # Add virtiofs device if socket exists
    if [ -S "$VIRTIOFSD_SOCKET" ]; then
        qemu_cmd+=(
            -chardev socket,id=char0,path="$VIRTIOFSD_SOCKET"
            -device vhost-user-fs-pci,queue-size=1024,chardev=char0,tag=revit-ballet
        )
        log "Shared filesystem enabled (tag: revit-ballet)"
    fi

    # Launch QEMU in background
    "${qemu_cmd[@]}" > /tmp/revit-ballet-vm.log 2>&1 &

    local vm_pid=$!
    echo $vm_pid > /tmp/revit-ballet-vm.pid

    log "VM launched with PID: $vm_pid"
    log "VM logs: /tmp/revit-ballet-vm.log"
    log "SPICE display available at: spice://127.0.0.1:$SPICE_PORT"
    if [ -S "$VIRTIOFSD_SOCKET" ]; then
        log "Shared directory: $SHARED_DIR (mounted in guest)"
    fi
}

# Wait for SSH to be available via vsock
wait_for_ssh() {
    log "Waiting for SSH service via vsock..."

    if [ ! -f "$SSH_KEY" ]; then
        error "SSH private key not found: $SSH_KEY"
        error "Please ensure SSH key is properly configured"
        exit 1
    fi

    local max_attempts=60
    local attempt=0

    while [ $attempt -lt $max_attempts ]; do
        if ssh -i "$SSH_KEY" \
               -o StrictHostKeyChecking=no \
               -o UserKnownHostsFile=/dev/null \
               -o ConnectTimeout=5 \
               -o BatchMode=yes \
               "vsock/$VM_CID" "exit 0" 2>/dev/null; then
            log "SSH connection established"
            return 0
        fi

        attempt=$((attempt + 1))
        echo -n "."
        sleep 2
    done

    echo ""
    error "Failed to establish SSH connection after $max_attempts attempts"
    return 1
}

# Get VM information via SSH
get_vm_info() {
    log "Gathering VM information..."

    ssh -i "$SSH_KEY" \
        -o StrictHostKeyChecking=no \
        -o UserKnownHostsFile=/dev/null \
        "vsock/$VM_CID" << 'EOF'
        Write-Host "=== System Information ==="
        $os = Get-CimInstance Win32_OperatingSystem
        Write-Host "OS: $($os.Caption)"
        Write-Host "Version: $($os.Version)"
        Write-Host ""

        Write-Host "=== CPU and Memory ==="
        $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
        Write-Host "CPU: $($cpu.Name)"
        Write-Host "Cores: $($cpu.NumberOfCores)"
        $mem = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
        $freeMem = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
        Write-Host "RAM: $freeMem GB free / $mem GB total"
        Write-Host ""

        Write-Host "=== Installed Revit Versions ==="
        $revitKeys = Get-ChildItem "HKLM:\SOFTWARE\Autodesk" -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Revit*" }
        if ($revitKeys) {
            foreach ($key in $revitKeys) {
                $versions = Get-ChildItem "Registry::$($key.Name)" -ErrorAction SilentlyContinue
                foreach ($ver in $versions) {
                    $verNum = $ver.PSChildName
                    if ($verNum -match '^\d+\.') {
                        $year = 2000 + [int]($verNum.Split('.')[0])
                        Write-Host "  - Revit $year"
                    }
                }
            }
        } else {
            Write-Host "  No Revit installations detected"
        }
        Write-Host ""

        Write-Host "=== Revit Ballet Installation ==="
        $balletPath = "$env:APPDATA\revit-ballet"
        if (Test-Path $balletPath) {
            Write-Host "  Installed at: $balletPath"

            # Check for active sessions
            $sessionFile = "$balletPath\runtime\network\sessions"
            if (Test-Path $sessionFile) {
                $sessions = Get-Content $sessionFile | Where-Object { $_ -match '\S' }
                Write-Host "  Active sessions: $($sessions.Count)"
            }
        } else {
            Write-Host "  Not installed"
        }
EOF
}

# Main execution
main() {
    log "Revit Ballet VM Launcher"
    log "========================"
    echo ""

    if check_vm_running; then
        log "Connecting to existing VM..."
    else
        launch_vm
    fi

    if wait_for_ssh; then
        echo ""
        get_vm_info
        echo ""
        log "VM is ready for testing"
        log "Connect via: ssh -i $SSH_KEY vsock/$VM_CID"
        log "Stop VM: kill \$(cat /tmp/revit-ballet-vm.pid)"
    else
        error "Failed to connect to VM"
        exit 1
    fi
}

main "$@"
