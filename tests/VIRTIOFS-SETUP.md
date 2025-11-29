# Setting Up virtiofs Shared Directory

This guide explains how to set up the shared `revit-ballet` directory between the Linux host and Windows guest VM using virtiofs.

## Why virtiofs?

**Benefits:**
- ✅ **Direct file access** - Host can read session registry without SSH
- ✅ **Real-time log monitoring** - Tail server logs directly on host
- ✅ **Faster** - No SSH overhead for file operations
- ✅ **Simpler tests** - Read session info with `cat ~/revit-ballet/runtime/network/sessions`
- ✅ **Better debugging** - Access all runtime files from host

**Architecture:**
```
Host: ~/revit-ballet/
  ↓ (virtiofsd daemon)
  ↓ (QEMU virtio-fs device, tag: "revit-ballet")
  ↓ (WinFsp driver on Windows)
  ↓ (Mounted to drive letter, e.g., Z:\revit-ballet)
  ↓ (Directory Junction)
Guest: %APPDATA%\revit-ballet\ → Z:\revit-ballet
```

## Setup Steps

### 1. Linux Host Setup

**Install virtiofsd** (if not already installed):

```bash
# Debian/Ubuntu
sudo apt install qemu-system-x86 virtiofsd

# Fedora/RHEL
sudo dnf install qemu-system-x86 virtiofsd

# Arch
sudo pacman -S qemu-base virtiofsd

# Verify
which virtiofsd
# Should output: /usr/lib/qemu/virtiofsd or similar
```

**Create shared directory:**

```bash
mkdir -p ~/revit-ballet/runtime/network
mkdir -p ~/revit-ballet/runtime/SearchboxQueries
mkdir -p ~/revit-ballet/runtime/LogViewChanges
mkdir -p ~/revit-ballet/runtime/screenshots
mkdir -p ~/revit-ballet/commands
```

**The `vm-launcher.sh` script handles the rest automatically:**
- Starts virtiofsd daemon
- Configures QEMU with virtiofs device
- Tags the share as "revit-ballet"

### 2. Windows Guest Setup

#### Step 1: Install WinFsp

WinFsp provides FUSE-like functionality for Windows and is required for virtiofs support.

**Download and install:**
```powershell
# Download from https://github.com/winfsp/winfsp/releases
# Latest stable version as of writing: 2.0

# Silent installation
Start-Process -Wait -FilePath "winfsp-2.0.23075.msi" -ArgumentList "/quiet"

# Verify installation
Get-Service WinFsp.Launcher
# Should show: Running
```

#### Step 2: Identify the virtiofs Mount Point

After VM starts with virtiofsd, Windows should auto-mount the share. Find it:

```powershell
# Check for WinFsp mounts
Get-PSDrive | Where-Object { $_.Provider -like "*WinFsp*" }

# Or check common locations
$locations = @(
    "\\?\VirtioFs\revit-ballet",
    "Z:\revit-ballet",
    "Y:\revit-ballet"
)

foreach ($loc in $locations) {
    if (Test-Path $loc) {
        Write-Host "Found virtiofs share at: $loc"
    }
}
```

**Note:** The exact mount point depends on your WinFsp configuration. Common patterns:
- `\\?\VirtioFs\<tag>` (where tag is "revit-ballet")
- Drive letter assigned by Windows (Z:, Y:, etc.)

#### Step 3: Create Directory Junction

Once you've found the mount point, create a junction to `%APPDATA%\revit-ballet`:

```powershell
# Replace with your actual virtiofs mount point
$source = "Z:\revit-ballet"  # Adjust this!
$target = "$env:APPDATA\revit-ballet"

# Backup existing directory if present
if (Test-Path $target) {
    if ((Get-Item $target).Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
        # It's a junction, just remove it
        (Get-Item $target).Delete()
    } else {
        # It's a real directory, back it up
        $backup = "$env:APPDATA\revit-ballet.backup"
        Move-Item $target $backup -Force
        Write-Host "Backed up existing directory to: $backup"
    }
}

# Create junction
New-Item -ItemType Junction -Path $target -Target $source -Force

# Verify
Get-Item $target | Select-Object Target, LinkType
# Should show:
# Target                 LinkType
# ------                 --------
# Z:\revit-ballet        Junction
```

#### Step 4: Test the Junction

```powershell
# Write a test file from Windows
"Hello from Windows" | Set-Content "$env:APPDATA\revit-ballet\test.txt"

# From Linux host, verify:
cat ~/revit-ballet/test.txt
# Should output: Hello from Windows

# Write from Linux
echo "Hello from Linux" > ~/revit-ballet/test2.txt

# From Windows, verify:
Get-Content "$env:APPDATA\revit-ballet\test2.txt"
# Should output: Hello from Linux
```

#### Step 5: Automate Junction Creation on Startup

Create a startup script to recreate the junction on every boot.

**Create script** (`C:\Scripts\mount-revit-ballet.ps1`):

```powershell
$ErrorActionPreference = "Stop"

# Wait for WinFsp to be ready
$maxWait = 30
$waited = 0
while (((Get-Service WinFsp.Launcher -ErrorAction SilentlyContinue).Status -ne 'Running') -and ($waited -lt $maxWait)) {
    Start-Sleep -Seconds 1
    $waited++
}

if ($waited -ge $maxWait) {
    Write-Error "WinFsp service did not start in time"
    exit 1
}

# Find virtiofs mount
$possibleSources = @(
    "\\?\VirtioFs\revit-ballet",
    "Z:\revit-ballet",
    "Y:\revit-ballet"
)

$source = $null
foreach ($possible in $possibleSources) {
    if (Test-Path $possible) {
        $source = $possible
        Write-Host "Found virtiofs share at: $source"
        break
    }
}

if ($null -eq $source) {
    Write-Error "Could not find virtiofs share"
    exit 1
}

# Create junction
$target = "$env:APPDATA\revit-ballet"

if (Test-Path $target) {
    $item = Get-Item $target
    if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
        $item.Delete()
    } else {
        $backup = "$target.backup-$(Get-Date -Format 'yyyyMMdd')"
        Move-Item $target $backup -Force
    }
}

New-Item -ItemType Junction -Path $target -Target $source -Force | Out-Null
Write-Host "Junction created: $target -> $source"

# Verify
if (Test-Path "$target\runtime") {
    Write-Host "Mount successful"
    exit 0
} else {
    Write-Error "Mount verification failed"
    exit 1
}
```

**Add to Task Scheduler:**

```powershell
# Create startup task
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-ExecutionPolicy Bypass -NoProfile -File C:\Scripts\mount-revit-ballet.ps1"

$trigger = New-ScheduledTaskTrigger -AtStartup

$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName "Mount Revit Ballet Share" `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Mounts virtiofs share for Revit Ballet testing" `
    -Force
```

**Verify task:**

```powershell
Get-ScheduledTask -TaskName "Mount Revit Ballet Share"

# Test run manually
Start-ScheduledTask -TaskName "Mount Revit Ballet Share"
```

## Using the Shared Directory in Tests

### From Linux Host (Direct Access)

```bash
# Read session registry directly
cat ~/revit-ballet/runtime/network/sessions

# Get first session's port
PORT=$(head -n1 ~/revit-ballet/runtime/network/sessions | cut -d',' -f2)

# Read auth token
TOKEN=$(cat ~/revit-ballet/runtime/network/token | tr -d '\r\n')

# Tail server logs in real-time
tail -f ~/revit-ballet/runtime/server.log

# Count active sessions
wc -l < ~/revit-ballet/runtime/network/sessions
```

### Test Script Example

```bash
#!/bin/bash
SHARED_DIR="$HOME/revit-ballet"

# Direct file access - no SSH needed!
if [ -f "$SHARED_DIR/runtime/network/sessions" ]; then
    session=$(head -n1 "$SHARED_DIR/runtime/network/sessions")
    port=$(echo "$session" | cut -d',' -f2)
    token=$(cat "$SHARED_DIR/runtime/network/token" | tr -d '\r\n')

    # Make HTTP request to Revit Ballet server
    # (still need SSH for curl since Windows is the guest, but reading files is local!)
fi
```

## Troubleshooting

### virtiofsd fails to start

```bash
# Check virtiofsd logs
cat /tmp/revit-ballet-virtiofsd.log

# Check if socket exists
ls -la /tmp/revit-ballet-virtiofs.sock

# Try starting manually
/usr/lib/qemu/virtiofsd \
    --socket-path=/tmp/test-virtiofs.sock \
    --shared-dir=$HOME/revit-ballet \
    --cache=auto
```

### Windows can't find the virtiofs share

```powershell
# Check if WinFsp is running
Get-Service WinFsp.*

# Check QEMU logs for virtiofs errors
# (on host: cat /tmp/revit-ballet-vm.log)

# Look for virtio-fs device in Device Manager
# Should appear under "Storage controllers" or similar
```

### Junction not working

```powershell
# Check junction target
Get-Item "$env:APPDATA\revit-ballet" | Select-Object Target, LinkType

# Delete and recreate
Remove-Item "$env:APPDATA\revit-ballet" -Force
New-Item -ItemType Junction -Path "$env:APPDATA\revit-ballet" -Target "Z:\revit-ballet"

# Verify with fsutil
fsutil reparsepoint query "$env:APPDATA\revit-ballet"
```

### Performance issues

If virtiofs is slow:

1. **Try different cache modes** in virtiofsd:
   ```bash
   # More aggressive caching (faster, but may have stale reads)
   virtiofsd --cache=always ...

   # Less caching (slower, but more consistent)
   virtiofsd --cache=none ...
   ```

2. **Check QEMU queue size:**
   ```bash
   # In vm-launcher.sh, increase queue-size
   -device vhost-user-fs-pci,queue-size=2048,...  # Default is 1024
   ```

3. **Disable sandboxing** (already done in vm-launcher.sh):
   ```bash
   virtiofsd --sandbox=none ...
   ```

## Fallback: SSH-based Access

If virtiofs doesn't work, all tests still work via SSH:

```bash
# Get session info via SSH instead of direct file access
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 << 'EOF'
    $session = Get-Content "$env:APPDATA\revit-ballet\runtime\network\sessions" | Select-Object -First 1
    $port = ($session -split ',')[1]
    $token = Get-Content "$env:APPDATA\revit-ballet\runtime\network\token" -Raw
    Write-Host "$port|$token"
EOF
```

The test scripts in this directory automatically fall back to SSH if virtiofs is not available.
