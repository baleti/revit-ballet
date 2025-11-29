#!/bin/bash

# Integration Test: InvokeAddinCommand
# Tests that the command discovery and DataGrid display works via server

set -euo pipefail

# Configuration
VM_CID=3
SSH_KEY="$HOME/.ssh/WinDev2407Eval-revit"
TEST_NAME="InvokeAddinCommand"

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

info() {
    echo -e "${BLUE}[i]${NC} $*"
}

# Get server details from VM
get_server_info() {
    ssh -i "$SSH_KEY" \
        -o StrictHostKeyChecking=no \
        -o UserKnownHostsFile=/dev/null \
        -o ConnectTimeout=5 \
        "vsock/$VM_CID" << 'EOF'
        $sessionFile = "$env:APPDATA\revit-ballet\runtime\network\sessions"
        $tokenFile = "$env:APPDATA\revit-ballet\runtime\network\token"

        if (-not (Test-Path $sessionFile) -or -not (Test-Path $tokenFile)) {
            Write-Host "ERROR:No session or token file found"
            exit 1
        }

        $sessions = Get-Content $sessionFile | Where-Object { $_ -match '\S' }
        if ($sessions.Count -eq 0) {
            Write-Host "ERROR:No active sessions"
            exit 1
        }

        # Use first session
        $parts = $sessions[0] -split ','
        $port = $parts[1]
        $token = (Get-Content $tokenFile -Raw).Trim()

        Write-Host "$port|$token"
EOF
}

# Execute script on Revit server
execute_script() {
    local port=$1
    local token=$2
    local script=$3

    ssh -i "$SSH_KEY" \
        -o StrictHostKeyChecking=no \
        -o UserKnownHostsFile=/dev/null \
        "vsock/$VM_CID" << EOF
        \$uri = "https://localhost:$port/roslyn"
        \$body = @{script = @"
$script
"@} | ConvertTo-Json

        try {
            \$response = Invoke-WebRequest -Uri \$uri \
                -Method POST \
                -Headers @{"X-Auth-Token" = "$token"} \
                -Body \$body \
                -ContentType "application/json" \
                -SkipCertificateCheck \
                -TimeoutSec 30 \
                -ErrorAction Stop

            \$response.Content
        } catch {
            Write-Host "ERROR:\$(\$_.Exception.Message)"
            exit 1
        }
EOF
}

# Test: Discover commands via reflection
test_command_discovery() {
    local port=$1
    local token=$2

    info "Test 1: Command discovery via reflection"

    local script='
// Simulate what InvokeAddinCommand does
var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
var year = UIApp.Application.VersionNumber;
var assemblyPath = Path.Combine(appData, "revit-ballet", "commands", "bin", year, "revit-ballet.dll");

if (!File.Exists(assemblyPath)) {
    return new { error = "Assembly not found: " + assemblyPath };
}

var assembly = Assembly.LoadFrom(assemblyPath);
var commandTypes = assembly.GetTypes()
    .Where(t => typeof(Autodesk.Revit.UI.IExternalCommand).IsAssignableFrom(t) && !t.IsAbstract)
    .ToList();

return new {
    assemblyPath = assemblyPath,
    commandCount = commandTypes.Count,
    sampleCommands = commandTypes.Take(5).Select(t => t.Name).ToArray()
};
'

    local result=$(execute_script "$port" "$token" "$script")

    if echo "$result" | grep -q "ERROR:"; then
        error "Script execution failed: $(echo "$result" | sed 's/ERROR://')"
        return 1
    fi

    # Parse JSON response
    local command_count=$(echo "$result" | jq -r '.result.commandCount // 0')

    if [ "$command_count" -gt 0 ]; then
        log "Discovered $command_count commands"

        local sample_commands=$(echo "$result" | jq -r '.result.sampleCommands[]' 2>/dev/null)
        if [ -n "$sample_commands" ]; then
            info "Sample commands:"
            echo "$sample_commands" | head -5 | sed 's/^/    /'
        fi

        return 0
    else
        error "No commands discovered (expected 200+)"
        echo "$result" | jq '.' || echo "$result"
        return 1
    fi
}

# Test: Verify command can be instantiated
test_command_instantiation() {
    local port=$1
    local token=$2

    info "Test 2: Command instantiation"

    local script='
var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
var year = UIApp.Application.VersionNumber;
var assemblyPath = Path.Combine(appData, "revit-ballet", "commands", "bin", year, "revit-ballet.dll");

var assembly = Assembly.LoadFrom(assemblyPath);

// Find a simple command to test
var commandType = assembly.GetTypes()
    .Where(t => typeof(Autodesk.Revit.UI.IExternalCommand).IsAssignableFrom(t) && !t.IsAbstract)
    .FirstOrDefault(t => t.Name.Contains("InvokeAddinCommand"));

if (commandType == null) {
    return new { error = "InvokeAddinCommand not found" };
}

// Try to instantiate
try {
    var instance = Activator.CreateInstance(commandType);
    return new {
        commandType = commandType.FullName,
        instantiated = true
    };
} catch (Exception ex) {
    return new { error = "Failed to instantiate: " + ex.Message };
}
'

    local result=$(execute_script "$port" "$token" "$script")

    if echo "$result" | grep -q "ERROR:"; then
        error "Script execution failed: $(echo "$result" | sed 's/ERROR://')"
        return 1
    fi

    local instantiated=$(echo "$result" | jq -r '.result.instantiated // false')

    if [ "$instantiated" = "true" ]; then
        local command_type=$(echo "$result" | jq -r '.result.commandType')
        log "Successfully instantiated: $command_type"
        return 0
    else
        error "Failed to instantiate command"
        echo "$result" | jq '.' || echo "$result"
        return 1
    fi
}

# Test: Check if DataGrid would show commands
test_datagrid_data() {
    local port=$1
    local token=$2

    info "Test 3: DataGrid data preparation"

    local script='
var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
var year = UIApp.Application.VersionNumber;
var assemblyPath = Path.Combine(appData, "revit-ballet", "commands", "bin", year, "revit-ballet.dll");

var assembly = Assembly.LoadFrom(assemblyPath);

var commandTypes = assembly.GetTypes()
    .Where(t => typeof(Autodesk.Revit.UI.IExternalCommand).IsAssignableFrom(t) && !t.IsAbstract)
    .OrderBy(t => t.Name)
    .ToList();

// Simulate DataGrid data
var gridData = commandTypes.Select(t => new {
    CommandName = t.Name,
    Namespace = t.Namespace
}).ToList();

return new {
    totalCommands = gridData.Count,
    uniqueNamespaces = gridData.Select(d => d.Namespace).Distinct().Count(),
    firstFive = gridData.Take(5)
};
'

    local result=$(execute_script "$port" "$token" "$script")

    if echo "$result" | grep -q "ERROR:"; then
        error "Script execution failed: $(echo "$result" | sed 's/ERROR://')"
        return 1
    fi

    local total=$(echo "$result" | jq -r '.result.totalCommands // 0')
    local namespaces=$(echo "$result" | jq -r '.result.uniqueNamespaces // 0')

    if [ "$total" -gt 0 ]; then
        log "DataGrid would show $total commands across $namespaces namespaces"

        info "First 5 commands that would appear:"
        echo "$result" | jq -r '.result.firstFive[] | "    \(.CommandName) (\(.Namespace))"' 2>/dev/null

        return 0
    else
        error "No data for DataGrid"
        return 1
    fi
}

# Main test execution
main() {
    echo "========================================="
    echo " Test: $TEST_NAME"
    echo "========================================="
    echo ""

    info "Getting server information..."
    local server_info=$(get_server_info)

    if echo "$server_info" | grep -q "ERROR:"; then
        error "Failed to get server info: $(echo "$server_info" | sed 's/ERROR://')"
        exit 1
    fi

    local port=$(echo "$server_info" | cut -d'|' -f1)
    local token=$(echo "$server_info" | cut -d'|' -f2)

    log "Connected to server on port $port"
    echo ""

    local failed=0

    test_command_discovery "$port" "$token" || ((failed++))
    echo ""

    test_command_instantiation "$port" "$token" || ((failed++))
    echo ""

    test_datagrid_data "$port" "$token" || ((failed++))
    echo ""

    echo "========================================="
    if [ $failed -eq 0 ]; then
        log "All tests passed! ✓"
        exit 0
    else
        error "$failed test(s) failed"
        exit 1
    fi
}

main "$@"
