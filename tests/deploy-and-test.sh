#!/bin/bash

# Deploy and Test - Complete workflow for testing Revit Ballet
# 1. Build installer (optional)
# 2. Deploy to VM
# 3. Install silently
# 4. Start Revit
# 5. Run tests

set -euo pipefail

# Configuration
VM_CID=3
SSH_KEY="$HOME/.ssh/WinDev2407Eval-revit"
INSTALLER_PATH="../installer/bin/Release/installer.exe"
REVIT_VERSIONS=("2024" "2025" "2026")  # Versions to test

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

# Run PowerShell command on VM
run_ps() {
    ssh -i "$SSH_KEY" \
        -o StrictHostKeyChecking=no \
        -o UserKnownHostsFile=/dev/null \
        -o ConnectTimeout=10 \
        "vsock/$VM_CID" "$@"
}

# Copy installer to VM
deploy_installer() {
    info "Deploying installer to VM..."

    if [ ! -f "$INSTALLER_PATH" ]; then
        error "Installer not found: $INSTALLER_PATH"
        error "Please build the installer first"
        return 1
    fi

    # Create temp directory on VM
    run_ps "New-Item -ItemType Directory -Force -Path C:\\Temp | Out-Null"

    # Copy installer
    scp -i "$SSH_KEY" \
        -o StrictHostKeyChecking=no \
        -o UserKnownHostsFile=/dev/null \
        "$INSTALLER_PATH" "vsock/$VM_CID:C:/Temp/revit-ballet-installer.exe"

    log "Installer deployed to C:\\Temp\\revit-ballet-installer.exe"
}

# Install Revit Ballet silently
install_revit_ballet() {
    info "Installing Revit Ballet (silent mode)..."

    run_ps << 'EOF'
        $installer = "C:\Temp\revit-ballet-installer.exe"

        if (-not (Test-Path $installer)) {
            Write-Host "ERROR: Installer not found"
            exit 1
        }

        # Run installer in quiet mode
        $process = Start-Process -FilePath $installer -ArgumentList "/q" -Wait -PassThru

        if ($process.ExitCode -eq 0) {
            Write-Host "Installation completed successfully"
            exit 0
        } else {
            Write-Host "ERROR: Installation failed with exit code $($process.ExitCode)"
            exit 1
        }
EOF

    if [ $? -eq 0 ]; then
        log "Revit Ballet installed successfully"
        return 0
    else
        error "Installation failed"
        return 1
    fi
}

# Start Revit with a test project
start_revit() {
    local version=$1
    info "Starting Revit $version..."

    run_ps << EOF
        \$revitPath = "C:\Program Files\Autodesk\Revit $version\Revit.exe"

        if (-not (Test-Path \$revitPath)) {
            Write-Host "ERROR: Revit $version not found at \$revitPath"
            exit 1
        }

        # Check if already running
        \$existing = Get-Process Revit -ErrorAction SilentlyContinue | Where-Object { \$_.MainWindowTitle -match "$version" }
        if (\$existing) {
            Write-Host "Revit $version already running (PID: \$(\$existing.Id))"
            exit 0
        }

        # Create empty test project if needed
        \$testProject = "C:\Temp\test-$version.rvt"
        if (-not (Test-Path \$testProject)) {
            # Start Revit, it will create new project
            Start-Process -FilePath \$revitPath
            Write-Host "Started Revit $version (creating new project)"
        } else {
            Start-Process -FilePath \$revitPath -ArgumentList \$testProject
            Write-Host "Started Revit $version with test project"
        }

        exit 0
EOF

    if [ $? -eq 0 ]; then
        log "Revit $version started"
        return 0
    else
        error "Failed to start Revit $version"
        return 1
    fi
}

# Wait for Revit Ballet server to be ready
wait_for_server() {
    info "Waiting for Revit Ballet server to start..."

    local max_attempts=30
    local attempt=0

    while [ $attempt -lt $max_attempts ]; do
        local session_count=$(run_ps << 'EOF'
            $sessionFile = "$env:APPDATA\revit-ballet\runtime\network\sessions"
            if (Test-Path $sessionFile) {
                $sessions = Get-Content $sessionFile | Where-Object { $_ -match '\S' }
                Write-Host $sessions.Count
            } else {
                Write-Host "0"
            }
EOF
)

        if [ "$session_count" -gt 0 ]; then
            log "Server ready ($session_count active session(s))"
            return 0
        fi

        attempt=$((attempt + 1))
        echo -n "."
        sleep 2
    done

    echo ""
    error "Server did not start within timeout"
    return 1
}

# Stop all Revit processes
stop_revit() {
    info "Stopping all Revit processes..."

    run_ps "Stop-Process -Name Revit -Force -ErrorAction SilentlyContinue"

    log "Revit processes stopped"
}

# Run all tests
run_tests() {
    info "Running integration tests..."

    local failed=0

    for test in test-*.sh; do
        if [ -f "$test" ]; then
            echo ""
            info "Running: $test"
            if ./"$test"; then
                log "$test passed"
            else
                error "$test failed"
                ((failed++))
            fi
        fi
    done

    echo ""
    if [ $failed -eq 0 ]; then
        log "All tests passed ✓"
        return 0
    else
        error "$failed test(s) failed"
        return 1
    fi
}

# Main workflow
main() {
    local build=false
    local deploy=false
    local install=false
    local start=false
    local test=false
    local cleanup=false

    # Parse arguments
    if [ $# -eq 0 ]; then
        # Default: full workflow
        deploy=true
        install=true
        start=true
        test=true
    else
        for arg in "$@"; do
            case $arg in
                --build) build=true ;;
                --deploy) deploy=true ;;
                --install) install=true ;;
                --start) start=true ;;
                --test) test=true ;;
                --cleanup) cleanup=true ;;
                --all)
                    build=true
                    deploy=true
                    install=true
                    start=true
                    test=true
                    ;;
                *)
                    echo "Unknown option: $arg"
                    echo "Usage: $0 [--build] [--deploy] [--install] [--start] [--test] [--cleanup] [--all]"
                    exit 1
                    ;;
            esac
        done
    fi

    echo "========================================="
    echo " Revit Ballet - Deploy and Test"
    echo "========================================="
    echo ""

    # Check VM connectivity first
    info "Checking VM connectivity..."
    if ! run_ps "exit 0" > /dev/null 2>&1; then
        error "Cannot connect to VM via SSH"
        error "Please start the VM first: ./vm-launcher.sh"
        exit 1
    fi
    log "VM connection OK"
    echo ""

    # Build (if requested)
    if [ "$build" = true ]; then
        info "Building installer..."
        cd ../installer
        dotnet build -c Release
        cd ../tests
        log "Build complete"
        echo ""
    fi

    # Deploy
    if [ "$deploy" = true ]; then
        deploy_installer || exit 1
        echo ""
    fi

    # Install
    if [ "$install" = true ]; then
        install_revit_ballet || exit 1
        echo ""
    fi

    # Start Revit
    if [ "$start" = true ]; then
        # Start first available version
        local started=false
        for version in "${REVIT_VERSIONS[@]}"; do
            if start_revit "$version"; then
                started=true
                break
            fi
        done

        if [ "$started" = false ]; then
            error "Could not start any Revit version"
            exit 1
        fi

        echo ""
        wait_for_server || exit 1
        echo ""

        # Give it a moment to stabilize
        sleep 5
    fi

    # Run health check
    info "Running health check..."
    ./health-check.sh
    echo ""

    # Run tests
    if [ "$test" = true ]; then
        run_tests || exit 1
    fi

    # Cleanup
    if [ "$cleanup" = true ]; then
        stop_revit
    fi

    echo ""
    log "Workflow complete!"
}

main "$@"
