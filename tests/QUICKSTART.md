# Quick Start Guide

## First Time Setup

### 1. Prepare Windows VM

On your Windows VM with Revit installed:

```powershell
# Install OpenSSH Server
Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0
Start-Service sshd
Set-Service -Name sshd -StartupType 'Automatic'

# Set PowerShell as default shell
New-ItemProperty -Path "HKLM:\SOFTWARE\OpenSSH" -Name DefaultShell `
    -Value "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
    -PropertyType String -Force

Restart-Service sshd
```

Copy your SSH public key to `C:\Users\<YourUser>\.ssh\authorized_keys`

### 2. Configure SSH on Linux Host

```bash
# Generate SSH key if you don't have one
ssh-keygen -t ed25519 -f ~/.ssh/WinDev2407Eval-revit

# Copy public key to VM (one-time, from VM side - see above)
```

### 3. Test SSH Connection

```bash
# Test connection
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 "echo 'Connection works!'"
```

## Daily Workflow

### Option 1: Automatic (Recommended)

```bash
cd /home/user/revit-ballet/tests

# Full workflow: deploy, install, start Revit, run tests
./deploy-and-test.sh --all
```

### Option 2: Manual (Step by Step)

```bash
# 1. Launch VM
./vm-launcher.sh

# 2. Deploy and install
./deploy-and-test.sh --deploy --install

# 3. Start Revit manually via SPICE or:
./deploy-and-test.sh --start

# 4. Check health
./health-check.sh

# 5. Run tests
./test-invoke-addin-command.sh
# ... or all tests:
./deploy-and-test.sh --test

# 6. Cleanup
./deploy-and-test.sh --cleanup
```

## Common Commands

```bash
# Just run tests (assumes everything is running)
./deploy-and-test.sh --test

# Deploy new version and test
./deploy-and-test.sh --deploy --install --test

# Health check only
./health-check.sh

# Connect to VM
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3

# View server logs on VM
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 \
    'Get-Content "$env:APPDATA\revit-ballet\runtime\server.log" -Tail 50'

# Stop VM
kill $(cat /tmp/revit-ballet-vm.pid)
```

## Writing a New Test

1. Copy the template from `README.md`
2. Name it `test-your-feature.sh`
3. Make it executable: `chmod +x test-your-feature.sh`
4. Run it: `./test-your-feature.sh`

## Troubleshooting

### "Cannot connect via SSH"

```bash
# Check if VM is running
pgrep -f qemu

# Check VM logs
tail /tmp/revit-ballet-vm.log

# Restart VM
./vm-launcher.sh
```

### "No active sessions"

```bash
# Verify Revit is running
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 "Get-Process Revit"

# Start Revit
./deploy-and-test.sh --start

# Or manually via SPICE display
remote-viewer spice://localhost:5900
```

### "Server not responding"

```bash
# Check server logs
ssh -i ~/.ssh/WinDev2407Eval-revit vsock/3 \
    'Get-Content "$env:APPDATA\revit-ballet\runtime\server.log" -Tail 50'

# Restart Revit
./deploy-and-test.sh --cleanup --start
```

### "Installer not found"

```bash
# Build installer first
cd ../installer
dotnet build -c Release
cd ../tests

# Or use build flag
./deploy-and-test.sh --build --deploy --install
```

## Tips

- **Run health check first** - It tells you what's missing
- **Use --all sparingly** - It rebuilds everything, use specific flags for faster iteration
- **Check resource usage** - health-check recommends parallelism level
- **Keep VM snapshots** - Before major changes, snapshot your VM
- **Monitor logs** - Add `Write-Host` to PowerShell scripts for debugging
- **Use jq** - Parse JSON responses easily: `echo "$result" | jq '.'`
