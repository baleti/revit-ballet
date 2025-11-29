# Revit Ballet Integration Testing

This directory contains integration and end-to-end tests for Revit Ballet, designed to run on a Linux hypervisor with Windows VMs via vsock.

## Architecture

```
Linux Hypervisor (Host)
├── ~/revit-ballet/ (shared via virtiofsd)
│   ├── runtime/network/sessions
│   ├── runtime/network/token
│   └── runtime/server.log
├── QEMU/KVM VM (Windows with Revit 2017-2026)
│   ├── %APPDATA%/revit-ballet/ (mounted from host via virtio-fs)
│   ├── Revit processes running revit-ballet addin
│   ├── revit-ballet servers (HTTPS on ports 23717+)
│   └── SSH server (accessible via vsock)
└── Test scripts (bash + curl)
    ├── Read session registry directly from ~/revit-ballet/
    ├── Connect via vsock/SSH
    ├── Send HTTP requests to servers
    └── Verify responses
```

### Shared Filesystem Architecture

The `revit-ballet` runtime directory is shared between host and guest using **virtiofsd** (virtio-fs):

- **Host path**: `~/revit-ballet/` (Linux)
- **Guest mount**: Maps to `%APPDATA%/revit-ballet/` (Windows)
- **Benefits**:
  - Host can read session registry without SSH
  - Host can read server logs in real-time
  - Host can read/write test data directly
  - Faster than scp or SSH file operations
  - Enables direct file-based orchestration

## Prerequisites

### On Linux Host

1. **QEMU/KVM** with vsock support
2. **SSH client** with vsock support
3. **virtiofsd** for filesystem sharing
4. **Standard tools**: `curl`, `jq`, `bash`

#### Setting up virtiofsd

```bash
# Create shared directory on host
mkdir -p ~/revit-ballet/runtime/network
mkdir -p ~/revit-ballet/runtime/SearchboxQueries
mkdir -p ~/revit-ballet/runtime/LogViewChanges
mkdir -p ~/revit-ballet/runtime/screenshots

# virtiofsd is typically included with QEMU
# Verify installation
which virtiofsd
# or
ls /usr/lib/qemu/virtiofsd

# The vm-launcher.sh script will include virtiofsd configuration
```

### On Windows VM

1. **Revit** versions 2017-2026 installed
2. **revit-ballet** installed (via `installer.exe /q`)
3. **SSH server** running with PowerShell as default shell
4. **SSH key** configured at `~/.ssh/WinDev2407Eval-revit` (private key on host)
5. **vsock** configured with guest CID=3
6. **virtio-fs** drivers installed (WinFsp or VirtIO-Win drivers)
7. **Shared folder** mounted to `%APPDATA%/revit-ballet`

#### Setting up virtio-fs on Windows

**Option 1: Using WinFsp + virtio-fs (Recommended)**

WinFsp provides FUSE-like functionality for Windows and supports virtio-fs:

```powershell
# Download and install WinFsp
# https://github.com/winfsp/winfsp/releases

# Install WinFsp (GUI installer or silent)
Start-Process -Wait -FilePath "winfsp-*.msi" -ArgumentList "/quiet"

# The virtio-fs driver will be available after QEMU starts with virtiofsd
# Mount will happen automatically if configured in QEMU command line
```

**Option 2: Manual mounting with VirtIO-Win drivers**

```powershell
# Download VirtIO-Win drivers
# https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/

# After VM starts with virtio-fs device, mount it
# The device will appear as a network share or drive

# Create symbolic link from mounted share to %APPDATA%/revit-ballet
$source = "Z:\revit-ballet"  # Mounted virtio-fs share
$target = "$env:APPDATA\revit-ballet"

# Remove existing directory if it exists
if (Test-Path $target) {
    Remove-Item $target -Recurse -Force
}

# Create directory junction (works like symlink for directories)
New-Item -ItemType Junction -Path $target -Target $source
```

**Option 3: Direct AppData mounting (Advanced)**

If you can configure the virtio-fs mount point in Windows to be directly at `%APPDATA%\revit-ballet`:

```powershell
# This requires configuring the mount point in Windows registry or via custom script
# The exact method depends on your virtio-fs implementation

# Verify the mount
Get-Item "$env:APPDATA\revit-ballet" | Select-Object Target, LinkType
```

**Verification:**

```powershell
# Test that the shared folder works
Set-Content "$env:APPDATA\revit-ballet\test.txt" -Value "Hello from Windows"

# From Linux host, verify:
# cat ~/revit-ballet/test.txt
# Should show "Hello from Windows"
```

### SSH Server Setup on Windows

Install OpenSSH Server and configure PowerShell:

```powershell
# Install OpenSSH Server
Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0

# Start SSH service
Start-Service sshd
Set-Service -Name sshd -StartupType 'Automatic'

# Configure PowerShell as default shell
New-ItemProperty -Path "HKLM:\SOFTWARE\OpenSSH" -Name DefaultShell `
    -Value "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -PropertyType String -Force

# Restart SSH service
Restart-Service sshd
```

Configure SSH key authentication and copy your public key to `~/.ssh/authorized_keys` on the VM.

### Mounting virtiofs Share on Windows Startup

To automatically mount the virtiofs share to `%APPDATA%\revit-ballet` on Windows startup, create a startup script:

**Create mount script** (`C:\Scripts\mount-revit-ballet.ps1`):

```powershell
# Mount virtiofs share to %APPDATA%\revit-ballet
# This script runs at Windows startup

$ErrorActionPreference = "Stop"

# virtiofs share appears as a virtio-fs device with tag "revit-ballet"
# With WinFsp installed, it should auto-mount to a drive letter or UNC path

# Wait for WinFsp service to be ready
$maxWait = 30
$waited = 0
while (-not (Get-Service WinFsp.Launcher -ErrorAction SilentlyContinue).Status -eq 'Running' -and $waited -lt $maxWait) {
    Start-Sleep -Seconds 1
    $waited++
}

# The virtiofs share should be available as \\?\VirtioFs\revit-ballet
# or similar, depending on WinFsp configuration

# Target: %APPDATA%\revit-ballet
$target = "$env:APPDATA\revit-ballet"

# Remove existing directory/junction if it exists
if (Test-Path $target) {
    $item = Get-Item $target
    if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
        # It's a junction/symlink, remove it
        $item.Delete()
    } else {
        # It's a regular directory, move it to backup
        $backup = "$env:APPDATA\revit-ballet.backup-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        Move-Item $target $backup -Force
        Write-Host "Existing directory backed up to: $backup"
    }
}

# Try to find the mounted virtiofs share
# This depends on your virtiofs/WinFsp configuration
# Common locations:
$possibleSources = @(
    "\\?\VirtioFs\revit-ballet",
    "Z:\revit-ballet",  # If auto-mounted to drive Z:
    "\\vboxsvr\revit-ballet"  # Alternative mounting
)

$source = $null
foreach ($possible in $possibleSources) {
    if (Test-Path $possible) {
        $source = $possible
        break
    }
}

if ($null -eq $source) {
    Write-Error "Could not find virtiofs share. Expected one of: $($possibleSources -join ', ')"
    exit 1
}

# Create directory junction (like symlink for directories)
New-Item -ItemType Junction -Path $target -Target $source -Force | Out-Null

Write-Host "Mounted virtiofs share from $source to $target"

# Verify
if (Test-Path "$target\runtime") {
    Write-Host "Mount successful - runtime directory accessible"
} else {
    Write-Error "Mount appears to have failed - runtime directory not found"
    exit 1
}
```

**Add to startup** (Task Scheduler method - recommended):

```powershell
# Create scheduled task to run at startup
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-ExecutionPolicy Bypass -File C:\Scripts\mount-revit-ballet.ps1"

$trigger = New-ScheduledTaskTrigger -AtStartup

$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName "Mount Revit Ballet Share" `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Mounts virtiofs share for Revit Ballet testing"
```

**Alternative: Manual mounting for testing**

```powershell
# Quick test without startup script
$target = "$env:APPDATA\revit-ballet"
if (Test-Path $target) { Remove-Item $target -Recurse -Force }

# Replace Z:\revit-ballet with your actual virtiofs mount point
New-Item -ItemType Junction -Path $target -Target "Z:\revit-ballet" -Force

# Verify
Get-ChildItem $target
```

**Important Notes:**

1. **Windows doesn't support mounting to arbitrary paths** like Linux. You must use:
   - Drive letters (e.g., `Z:\`)
   - UNC paths (e.g., `\\?\VirtioFs\...`)
   - Then create a Junction to `%APPDATA%\revit-ballet`

2. **Directory Junctions work perfectly** for this use case:
   - They're transparent to applications
   - revit-ballet will work exactly as if files were local
   - No performance penalty

3. **WinFsp is required** for virtiofs on Windows:
   - Download: https://github.com/winfsp/winfsp/releases
   - Install silently: `msiexec /i winfsp-*.msi /quiet`

4. **Verify the mount** before running tests:
   ```powershell
   # Check if junction exists and points to correct target
   Get-Item "$env:APPDATA\revit-ballet" | Select-Object Target, LinkType

   # Test file access
   Get-ChildItem "$env:APPDATA\revit-ballet\runtime\network"
   ```

## Test Scripts

### `vm-launcher.sh`

Launches the Windows VM and establishes vsock SSH connection.

```bash
./vm-launcher.sh
```

Features:
- Sets up shared directory (`~/revit-ballet/`)
- Starts virtiofsd daemon for filesystem sharing
- Launches QEMU VM with proper vsock and virtiofs configuration
- Waits for SSH to become available
- Displays system information
- Reports installed Revit versions
- Shows active revit-ballet sessions

The shared directory allows the host to directly access:
- Session registry (`~/revit-ballet/runtime/network/sessions`)
- Auth token (`~/revit-ballet/runtime/network/token`)
- Server logs (`~/revit-ballet/runtime/server.log`)
- Test data and results

### `health-check.sh`

Comprehensive health check of the testing environment.

```bash
./health-check.sh
```

Checks:
- ✓ SSH connectivity via vsock
- ✓ Revit Ballet installation
- ✓ Running Revit processes
- ✓ Active server sessions
- ✓ Server HTTP endpoints responding
- ✓ System resources (CPU, RAM)
- ✓ Recommended parallelism level

Exit codes:
- `0` - All checks passed
- `1` - One or more checks failed

### `test-invoke-addin-command.sh`

Integration test for the `InvokeAddinCommand` functionality.

```bash
./test-invoke-addin-command.sh
```

Tests:
1. **Command Discovery**: Verifies that commands can be discovered via reflection
2. **Command Instantiation**: Verifies commands can be instantiated
3. **DataGrid Data**: Verifies data preparation for UI display

This test does NOT test the actual WinForms DataGrid (that would require UI automation), but tests the core logic that prepares the data.

## Writing New Tests

### Test Template

```bash
#!/bin/bash
set -euo pipefail

VM_CID=3
SSH_KEY="$HOME/.ssh/WinDev2407Eval-revit"
SHARED_DIR="$HOME/revit-ballet"  # If using virtiofs
TEST_NAME="YourTestName"

# Get server info (Option 1: Direct file access via virtiofs - FASTER)
get_server_info_direct() {
    if [ ! -f "$SHARED_DIR/runtime/network/sessions" ]; then
        echo "ERROR: Shared directory not available" >&2
        return 1
    fi

    local session=$(head -n1 "$SHARED_DIR/runtime/network/sessions")
    local port=$(echo "$session" | cut -d',' -f2)
    local token=$(cat "$SHARED_DIR/runtime/network/token" | tr -d '\r\n')

    echo "$port|$token"
}

# Get server info (Option 2: Via SSH - works without virtiofs)
get_server_info_ssh() {
    ssh -i "$SSH_KEY" -o StrictHostKeyChecking=no "vsock/$VM_CID" << 'EOF'
        $sessionFile = "$env:APPDATA\revit-ballet\runtime\network\sessions"
        $tokenFile = "$env:APPDATA\revit-ballet\runtime\network\token"

        $sessions = Get-Content $sessionFile
        $port = ($sessions[0] -split ',')[1]
        $token = (Get-Content $tokenFile -Raw).Trim()

        Write-Host "$port|$token"
EOF
}

# Get server info (try direct first, fallback to SSH)
get_server_info() {
    get_server_info_direct 2>/dev/null || get_server_info_ssh
}

# Execute test script
execute_script() {
    local port=$1
    local token=$2
    local script=$3

    ssh -i "$SSH_KEY" "vsock/$VM_CID" << EOF
        \$response = Invoke-WebRequest -Uri "https://localhost:$port/roslyn" \
            -Method POST \
            -Headers @{"X-Auth-Token" = "$token"} \
            -Body (@{script = @"
$script
"@} | ConvertTo-Json) \
            -ContentType "application/json" \
            -SkipCertificateCheck

        \$response.Content
EOF
}

# Main test
main() {
    local server_info=$(get_server_info)
    local port=$(echo "$server_info" | cut -d'|' -f1)
    local token=$(echo "$server_info" | cut -d'|' -f2)

    local result=$(execute_script "$port" "$token" "return 1 + 1;")
    local value=$(echo "$result" | jq -r '.result')

    if [ "$value" = "2" ]; then
        echo "✓ Test passed"
        exit 0
    else
        echo "✗ Test failed"
        exit 1
    fi
}

main "$@"
```

### C# Script Examples

**Access Revit API:**
```csharp
// UIApp, UIDoc, Doc are available as globals
return new {
    documentTitle = Doc.Title,
    viewName = UIDoc.ActiveView.Name,
    elementCount = new FilteredElementCollector(Doc)
        .WhereElementIsNotElementType()
        .ToElements()
        .Count
};
```

**Test command execution:**
```csharp
// Load command assembly
var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
var year = UIApp.Application.VersionNumber;
var assemblyPath = Path.Combine(appData, "revit-ballet", "commands", "bin", year, "revit-ballet.dll");
var assembly = Assembly.LoadFrom(assemblyPath);

// Find and instantiate command
var cmdType = assembly.GetTypes()
    .First(t => t.Name == "SomeCommand");
var instance = (IExternalCommand)Activator.CreateInstance(cmdType);

// Execute (requires proper context - be careful!)
// var result = instance.Execute(commandData, message, elements);
```

## Workflow

### 1. Setup

```bash
# Launch VM
cd /home/user/revit-ballet/tests
./vm-launcher.sh

# In another terminal, verify health
./health-check.sh
```

### 2. Install/Update Revit Ballet

```bash
# From host, copy installer to VM (if needed)
scp -i ~/.ssh/WinDev2407Eval-revit installer.exe vsock/3:C:/Users/User/Downloads/

# Install silently
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 \
    "C:/Users/User/Downloads/installer.exe /q"
```

### 3. Start Revit

```bash
# Start Revit 2024 with a test project
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 \
    'Start-Process "C:\Program Files\Autodesk\Revit 2024\Revit.exe" -ArgumentList "C:\Projects\test.rvt"'

# Wait for server to start
sleep 30

# Verify
./health-check.sh
```

### 4. Run Tests

```bash
# Run all tests
for test in test-*.sh; do
    echo "Running $test..."
    ./"$test" || echo "FAILED: $test"
done

# Or run specific test
./test-invoke-addin-command.sh
```

### 5. Cleanup

```bash
# Stop Revit
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 "Stop-Process -Name Revit -Force"

# Stop VM
kill $(cat /tmp/revit-ballet-vm.pid)
```

## Orchestration Considerations

### Sequential vs Parallel

**Sequential** (safer for now):
```bash
for version in 2024 2025 2026; do
    start_revit "$version"
    run_tests "$version"
    stop_revit "$version"
done
```

**Parallel** (requires resource monitoring):
```bash
# Check available resources first
./health-check.sh | grep "Recommended parallel instances"

# Start multiple Revit instances if resources allow
start_revit 2024 &
start_revit 2025 &
wait

# Run tests against both
./test-invoke-addin-command.sh  # Will test all active sessions
```

### Resource Monitoring

The `health-check.sh` script recommends parallelism based on:
- **CPU**: Each Revit instance benefits from ~2 cores
- **Memory**: Each Revit instance uses ~1.5GB RAM
- **Current load**: Takes into account existing processes

## Debugging

### View server logs

```bash
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 \
    'Get-Content "$env:APPDATA\revit-ballet\runtime\server.log" -Tail 50'
```

### Manual HTTP request

```bash
# Get port and token
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 << 'EOF'
$port = ((Get-Content "$env:APPDATA\revit-ballet\runtime\network\sessions")[0] -split ',')[1]
$token = (Get-Content "$env:APPDATA\revit-ballet\runtime\network\token" -Raw).Trim()
Write-Host "Port: $port"
Write-Host "Token: $token"
EOF

# Send test request
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 << 'EOF'
Invoke-WebRequest -Uri "https://localhost:PORT/roslyn" `
    -Method POST `
    -Headers @{"X-Auth-Token" = "TOKEN"} `
    -Body '{"script":"return Doc.Title;"}' `
    -ContentType "application/json" `
    -SkipCertificateCheck
EOF
```

### Connect to SPICE display

```bash
# From another machine on the network
remote-viewer spice://YOUR_HOST_IP:5900
```

## Limitations

### What These Tests Can Do
- ✓ Test server HTTP endpoints
- ✓ Test script compilation and execution
- ✓ Test command discovery and instantiation
- ✓ Test Revit API interactions
- ✓ Verify session management
- ✓ Check resource usage

### What These Tests Cannot Do
- ✗ Test WinForms UI directly (DataGrid, dialogs)
- ✗ Test keyboard shortcuts
- ✗ Test user interactions (clicks, selections)
- ✗ Test visual rendering

For UI testing, you would need additional tools like:
- **UI Automation** (Windows Automation API)
- **SikuliX** (image-based UI testing)
- **TestStack.White** (.NET UI automation library)

## Future Enhancements

1. **Test Framework**: Consider using BATS (Bash Automated Testing System) for better test organization
2. **CI/CD Integration**: GitHub Actions with self-hosted runner on the Linux host
3. **Test Coverage**: Measure which commands have integration tests
4. **Performance Tests**: Track execution time trends
5. **Snapshot Testing**: Compare results across Revit versions
6. **Parallel Execution**: Implement proper parallel test runner with resource management
