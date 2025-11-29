#!/bin/bash

# Health Check for Revit Ballet Testing Environment
# Verifies VM state, Revit processes, and server availability

set -euo pipefail

# Configuration
VM_CID=3
SSH_KEY="$HOME/.ssh/WinDev2407Eval-revit"
TOKEN_FILE="%APPDATA%/revit-ballet/runtime/network/token"
SESSIONS_FILE="%APPDATA%/revit-ballet/runtime/network/sessions"

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

log() {
    echo -e "${GREEN}[✓]${NC} $*"
}

error() {
    echo -e "${RED}[✗]${NC} $*"
}

warn() {
    echo -e "${YELLOW}[!]${NC} $*"
}

info() {
    echo -e "${BLUE}[i]${NC} $*"
}

# Execute PowerShell command on VM via SSH
run_ps() {
    ssh -i "$SSH_KEY" \
        -o StrictHostKeyChecking=no \
        -o UserKnownHostsFile=/dev/null \
        -o ConnectTimeout=5 \
        "vsock/$VM_CID" "$@" 2>/dev/null
}

# Check SSH connectivity
check_ssh() {
    info "Checking SSH connectivity..."
    if run_ps "exit 0" > /dev/null 2>&1; then
        log "SSH connection OK"
        return 0
    else
        error "Cannot connect via SSH"
        return 1
    fi
}

# Check Revit processes
check_revit_processes() {
    info "Checking for Revit processes..."

    local result=$(run_ps 'Get-Process Revit -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count')

    if [ "$result" -eq 0 ]; then
        warn "No Revit processes running"
        return 1
    else
        log "Found $result Revit process(es)"
        return 0
    fi
}

# Get detailed process information
get_process_details() {
    info "Getting Revit process details..."

    run_ps << 'EOF'
        $processes = Get-Process Revit -ErrorAction SilentlyContinue
        if ($processes) {
            foreach ($proc in $processes) {
                Write-Host "  PID: $($proc.Id)"
                Write-Host "  CPU: $([math]::Round($proc.CPU, 2))s"
                Write-Host "  Memory: $([math]::Round($proc.WorkingSet64 / 1MB, 2)) MB"
                Write-Host "  Started: $($proc.StartTime)"
                Write-Host ""
            }
        }
EOF
}

# Check Revit Ballet installation
check_installation() {
    info "Checking Revit Ballet installation..."

    local installed=$(run_ps 'Test-Path "$env:APPDATA\revit-ballet"')

    if [ "$installed" = "True" ]; then
        log "Revit Ballet is installed"

        # Get version info if available
        local dll_count=$(run_ps 'Get-ChildItem "$env:APPDATA\revit-ballet\commands\bin" -Recurse -Filter "revit-ballet.dll" -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count')
        info "Found $dll_count version-specific DLLs"

        return 0
    else
        error "Revit Ballet is not installed"
        return 1
    fi
}

# Check server sessions
check_sessions() {
    info "Checking active server sessions..."

    run_ps << 'EOF'
        $sessionFile = "$env:APPDATA\revit-ballet\runtime\network\sessions"
        if (Test-Path $sessionFile) {
            $sessions = Get-Content $sessionFile | Where-Object { $_ -match '\S' }
            if ($sessions.Count -gt 0) {
                Write-Host "Active sessions: $($sessions.Count)"
                foreach ($session in $sessions) {
                    $parts = $session -split ','
                    if ($parts.Count -ge 4) {
                        Write-Host "  Port: $($parts[1]) | Docs: $($parts[3]) | PID: $($parts[2])"
                    }
                }
                exit 0
            } else {
                Write-Host "No active sessions"
                exit 1
            }
        } else {
            Write-Host "Session file not found"
            exit 2
        }
EOF

    local exit_code=$?
    if [ $exit_code -eq 0 ]; then
        log "Active sessions found"
        return 0
    elif [ $exit_code -eq 1 ]; then
        warn "No active sessions"
        return 1
    else
        error "Session file not found"
        return 2
    fi
}

# Test server connectivity
test_server() {
    info "Testing server connectivity..."

    run_ps << 'EOF'
        $sessionFile = "$env:APPDATA\revit-ballet\runtime\network\sessions"
        $tokenFile = "$env:APPDATA\revit-ballet\runtime\network\token"

        if (-not (Test-Path $sessionFile)) {
            Write-Host "No sessions file"
            exit 1
        }

        if (-not (Test-Path $tokenFile)) {
            Write-Host "No token file"
            exit 1
        }

        $token = Get-Content $tokenFile -Raw
        $token = $token.Trim()

        $sessions = Get-Content $sessionFile | Where-Object { $_ -match '\S' }
        $testedPorts = @()

        foreach ($session in $sessions) {
            $parts = $session -split ','
            if ($parts.Count -ge 2) {
                $port = $parts[1]

                if ($testedPorts -contains $port) {
                    continue
                }
                $testedPorts += $port

                try {
                    # Simple ping test
                    $uri = "https://localhost:$port/roslyn"
                    $body = @{script = "return 1 + 1;"} | ConvertTo-Json

                    $response = Invoke-WebRequest -Uri $uri `
                        -Method POST `
                        -Headers @{"X-Auth-Token" = $token} `
                        -Body $body `
                        -ContentType "application/json" `
                        -SkipCertificateCheck `
                        -TimeoutSec 5 `
                        -ErrorAction Stop

                    $result = $response.Content | ConvertFrom-Json

                    if ($result.result -eq 2) {
                        Write-Host "✓ Server on port $port is responding correctly"
                    } else {
                        Write-Host "✗ Server on port $port returned unexpected result: $($result.result)"
                        exit 1
                    }
                } catch {
                    Write-Host "✗ Failed to connect to server on port $port : $($_.Exception.Message)"
                    exit 1
                }
            }
        }

        Write-Host "All servers responding"
        exit 0
EOF

    local exit_code=$?
    if [ $exit_code -eq 0 ]; then
        log "All servers responding correctly"
        return 0
    else
        error "Server connectivity test failed"
        return 1
    fi
}

# Check system resources
check_resources() {
    info "Checking system resources..."

    run_ps << 'EOF'
        $os = Get-CimInstance Win32_OperatingSystem
        $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1

        $cpuLoad = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction SilentlyContinue).CounterSamples.CookedValue
        $memUsedPct = [math]::Round((($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize) * 100, 1)
        $memFreeGB = [math]::Round($os.FreePhysicalMemory / 1MB, 2)

        Write-Host "  CPU Load: $([math]::Round($cpuLoad, 1))%"
        Write-Host "  Memory: $memUsedPct% used ($memFreeGB GB free)"

        # Check if resources are constrained
        if ($cpuLoad -gt 90) {
            Write-Host "  WARNING: High CPU usage"
        }

        if ($memUsedPct -gt 90) {
            Write-Host "  WARNING: High memory usage"
        }
EOF

    log "Resource check complete"
}

# Determine recommended parallel test count
recommend_parallelism() {
    info "Analyzing system capacity for parallel testing..."

    run_ps << 'EOF'
        $os = Get-CimInstance Win32_OperatingSystem
        $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1

        $totalCores = $cpu.NumberOfLogicalProcessors
        $memTotalGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
        $memFreeGB = [math]::Round($os.FreePhysicalMemory / 1MB, 2)

        # Estimate: each Revit instance uses ~1.5GB RAM and benefits from 2 cores
        $maxByMem = [math]::Floor($memFreeGB / 1.5)
        $maxByCpu = [math]::Floor($totalCores / 2)

        $recommended = [math]::Min($maxByMem, $maxByCpu)
        $recommended = [math]::Max(1, $recommended)

        Write-Host "  Total Cores: $totalCores"
        Write-Host "  Available Memory: $memFreeGB GB"
        Write-Host "  Max by CPU: $maxByCpu instances"
        Write-Host "  Max by Memory: $maxByMem instances"
        Write-Host ""
        Write-Host "Recommended parallel instances: $recommended"
EOF
}

# Main health check
main() {
    echo "=================================="
    echo " Revit Ballet Health Check"
    echo "=================================="
    echo ""

    local failed=0

    check_ssh || ((failed++))
    echo ""

    check_installation || ((failed++))
    echo ""

    check_revit_processes || ((failed++))
    if [ $? -eq 0 ]; then
        get_process_details
    fi
    echo ""

    check_sessions || warn "Consider starting Revit with revit-ballet addin loaded"
    echo ""

    if check_sessions > /dev/null 2>&1; then
        test_server || ((failed++))
        echo ""
    fi

    check_resources
    echo ""

    recommend_parallelism
    echo ""

    if [ $failed -eq 0 ]; then
        log "All health checks passed"
        exit 0
    else
        error "$failed health check(s) failed"
        exit 1
    fi
}

main "$@"
