# Revit Ballet - AI Agent Integration & Code Conventions

> **Documentation Size Limit**: This file must stay under 20KB. Keep changes concise and avoid approaching this limit to allow for future additions.

This file provides guidance to AI agents (like Claude Code) when working with code in this repository and querying Revit sessions.

## Project Overview

Revit Ballet is a collection of custom commands for Revit that provides enhanced productivity tools. The project consists of two main components:

1. **Commands** (`/commands/`) - C# plugin with custom Revit commands
2. **Installer** (`/installer/`) - Windows Forms installer application

## Development Environment

**IMPORTANT**: This codebase is located in the same directory as `%APPDATA%/revit-ballet`, which means:
- The `runtime/` folder in this repository reflects **live usage** of the plugin
- You can examine runtime logs and data to verify code behavior and debug issues
- **ALL runtime data MUST be stored in `runtime/` subdirectory** - includes logs, network data, screenshots

## Architecture

### Multi-Version Revit Support
The project supports Revit versions 2017-2026 through conditional compilation:
- Uses `RevitYear` property to target specific versions
- Different .NET Framework targets based on Revit version:
  - 2017-2018: .NET Framework 4.6
  - 2019-2020: .NET Framework 4.7
  - 2021-2024: .NET Framework 4.8
  - 2025-2026: .NET 8.0 Windows
- Revit.NET package versions are automatically selected based on target year

### Roslyn Server for AI Agents

Revit Ballet provides a Roslyn compiler-as-a-service that allows AI agents to execute C# code within Revit session context.

**Purpose:**
- Query and analyze current Revit session programmatically
- Perform validation tasks (element checks, parameter validation)
- Inspect document state and element properties dynamically
- Support automated BIM quality checks and audits
- Enable cross-session network queries for multi-Revit workflows

**Peer-to-Peer Network Architecture:**
- Each Revit instance runs its own HTTPS server
- Servers auto-start when Revit loads (no manual command needed)
- Port allocation: 23717-23817 (automatic first-available selection)
- Shared token authentication across all sessions
- File-based service discovery via `%appdata%/revit-ballet/runtime/network/`

**Network Registry Structure:**
```
%appdata%/revit-ballet/runtime/network/
  ├── sessions      # CSV: SessionId,Port,Hostname,ProcessId,RegisteredAt,LastHeartbeat,Documents
  └── token         # Shared 256-bit authentication token
```

**Design Decision: Shared Token**
- All Revit sessions in the network share a single authentication token
- First session to start generates the token and stores it in `network/token`
- Subsequent sessions read and reuse this token
- Token regenerates when all sessions close and a new one starts
- Simplifies authentication: only one token to manage
- Enables seamless cross-session communication

**Endpoints:**
- **`/roslyn`** - Execute C# scripts in Revit context
- **`/screenshot`** - Capture screenshot of Revit window, saved to `runtime/screenshots/`

**Implementation:**
- Auto-initialized by `RevitBallet` extension application in `commands/RevitBallet.cs`
- Background thread with UI marshalling via `ExternalEvent`
- Network registry heartbeat every 30 seconds, 2-minute timeout for dead sessions

## Querying and Testing Revit Sessions

**When you need to query Revit state, test commands, or inspect models:**

1. **Use the Roslyn Server** - Every Revit session runs an HTTPS server that executes C# scripts
2. **Location**: Token at `runtime/network/token`, sessions at `runtime/network/sessions`
3. **Find active port**: `grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2`

**IMPORTANT - Query File Pattern:**

For anything beyond the simplest single-line queries, **ALWAYS use a file-based approach** to avoid bash escaping issues:

```bash
# 1. Write your C# query to a file in /tmp
cat > /tmp/query-$(uuidgen).cs << 'EOF'
var collector = new FilteredElementCollector(Doc);
var levels = collector.OfClass(typeof(Level)).Cast<Level>();
Console.WriteLine("Total Levels: " + levels.Count());
foreach (var level in levels.OrderBy(l => l.Elevation).Take(5))
{
    Console.WriteLine("  " + level.Name + ": " + level.Elevation);
}
EOF

# 2. Send the file with --data-binary
TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)
curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  --data-binary @/tmp/query-*.cs | jq -r '.Output'
```

**CRITICAL - Claude Code Bash Tool Limitation & Workaround:**

Claude Code's Bash tool has a bug where **multiple command substitutions `$(...)` in one call fail with escaping errors**.

**Workaround:** Write the bash script to a file and execute it:

```bash
# Tool call 1: Write script to file
cat > /tmp/query-revit.sh << 'EOF'
#!/bin/bash
TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)
curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  --data-binary @/tmp/query-*.cs | jq -r '.Output'
EOF

# Tool call 2: Execute the script
bash /tmp/query-revit.sh
```

**Alternative:** Use only one `$(...)` per Bash tool call with hardcoded values:
```bash
# Hardcode port 23717 and use token substitution
curl -k -s -X POST https://127.0.0.1:23717/roslyn \
  -H "X-Auth-Token: $(cat runtime/network/token)" \
  --data-binary @/tmp/query-test.cs | jq -r '.Output'
```

**Recommended Helper Script Pattern:**

For frequent queries, create a reusable helper script:

```bash
# Create helper once
cat > /tmp/revit-query.sh << 'EOF'
#!/bin/bash
TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)
curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  --data-binary @"$1" | jq
EOF
chmod +x /tmp/revit-query.sh

# Use it for any query
cat > /tmp/my-query.cs << 'EOF'
Console.WriteLine("Document: " + Doc.Title);
EOF

bash /tmp/revit-query.sh /tmp/my-query.cs
```

**Why use files?**
- Avoids bash escaping nightmares with quotes, operators (!=, ?, etc.)
- Allows proper multi-line C# formatting
- Enables complex queries with conditionals and loops
- Supports comments in your C# code

**Simple inline queries (single-line, no special characters):**
```bash
TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)
curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  -d 'Console.WriteLine("Hello from Revit");' | jq -r '.Output'
```

**Note:** Use `-k` or `--insecure` flag with curl to accept self-signed SSL certificates (localhost only).

## Available Context

Scripts have access to these global variables:

- **`UIApp`** - UIApplication instance
- **`UIDoc`** - Active UIDocument
- **`Doc`** - Active Document

Pre-imported namespaces:
- `System`, `System.Linq`, `System.Collections.Generic`
- `Autodesk.Revit.DB`
- `Autodesk.Revit.UI`

**Response Format:**
```json
{
  "Success": true,
  "Output": "Hello from Revit!\n",
  "Error": null,
  "Diagnostics": []
}
```

## Common Query Patterns

**NOTE:** All examples below use the file-based approach to avoid escaping issues. For helper function to simplify queries, see the Integration section below.

### Listing Information

**Get All Levels:**
```bash
cat > /tmp/query-$(uuidgen).cs << 'EOF'
var collector = new FilteredElementCollector(Doc);
var levels = collector.OfClass(typeof(Level)).Cast<Level>();
Console.WriteLine("Total Levels: " + levels.Count());
foreach (var level in levels.OrderBy(l => l.Elevation).Take(5))
{
    Console.WriteLine("  " + level.Name + ": " + level.Elevation);
}
EOF

TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)
curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  --data-binary @/tmp/query-*.cs | jq -r '.Output'
```

**Count Elements by Category:**
```bash
cat > /tmp/query-$(uuidgen).cs << 'EOF'
var collector = new FilteredElementCollector(Doc);
var elements = collector.WhereElementIsNotElementType().ToElements();
var typeCounts = elements
    .GroupBy(e => e.Category?.Name ?? "None")
    .OrderByDescending(g => g.Count())
    .Take(5);

foreach (var group in typeCounts)
{
    Console.WriteLine(group.Key + ": " + group.Count());
}
EOF

TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)
curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  --data-binary @/tmp/query-*.cs | jq -r '.Output'
```

### Validation Queries

**Check for Unplaced Rooms:**
```bash
cat > /tmp/query-$(uuidgen).cs << 'EOF'
var collector = new FilteredElementCollector(Doc);
var rooms = collector.OfCategory(BuiltInCategory.OST_Rooms).Cast<Room>();
var unplaced = rooms.Where(r => r.Area == 0 || r.Location == null).ToList();

if (unplaced.Count > 0)
{
    Console.WriteLine("Found " + unplaced.Count + " unplaced rooms:");
    foreach (var room in unplaced)
    {
        Console.WriteLine("  - Room " + room.Number + ": " + room.Name);
    }
}
else
{
    Console.WriteLine("All rooms are placed!");
}
EOF

TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)
curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  --data-binary @/tmp/query-*.cs | jq -r '.Output'
```

**Check for Elements Without Parameters:**
```bash
cat > /tmp/query-$(uuidgen).cs << 'EOF'
var collector = new FilteredElementCollector(Doc);
var walls = collector.OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType();
var missingMark = new List<Element>();

foreach (var wall in walls)
{
    var mark = wall.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
    if (mark == null || string.IsNullOrEmpty(mark.AsString()))
    {
        missingMark.Add(wall);
    }
}

Console.WriteLine("Walls missing Mark parameter: " + missingMark.Count);
EOF

TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)
curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  --data-binary @/tmp/query-*.cs | jq -r '.Output'
```

### Statistical Queries

**Get Document Statistics:**
```bash
cat > /tmp/query-$(uuidgen).cs << 'EOF'
var collector = new FilteredElementCollector(Doc);
var allElements = collector.WhereElementIsNotElementType().ToElements();
var categories = allElements
    .Where(e => e.Category != null)
    .GroupBy(e => e.Category.Name)
    .OrderByDescending(g => g.Count())
    .Take(10);

Console.WriteLine("Top 10 Categories by Count:");
foreach (var cat in categories)
{
    Console.WriteLine("  " + cat.Key + ": " + cat.Count());
}

var levels = new FilteredElementCollector(Doc).OfClass(typeof(Level)).ToElements().Count;
var views = new FilteredElementCollector(Doc).OfClass(typeof(View)).ToElements().Count;
Console.WriteLine("Total Levels: " + levels);
Console.WriteLine("Total Views: " + views);
EOF

TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)
curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  --data-binary @/tmp/query-*.cs | jq -r '.Output'
```

### Screenshot Capture

**Capture Revit Window Screenshot:**
```bash
curl --insecure -X POST https://127.0.0.1:23717/screenshot -H "X-Auth-Token: $(cat revit-ballet/runtime/network/token)" | jq
```

**Response:**
```json
{
  "Success": true,
  "Output": "C:\\Users\\YourName\\AppData\\Roaming\\revit-ballet\\runtime\\screenshots\\20231116-143022-456-session-id.png",
  "Error": null,
  "Diagnostics": []
}
```

**Features:**
- Captures entire Revit window including ribbon, properties panel, and view area
- Screenshots saved to `%appdata%/revit-ballet/runtime/screenshots/`
- Filename format: `{timestamp}-{sessionId}.png`
- Auto-cleanup: Keeps last 20 screenshots, older ones deleted automatically
- Returns absolute file path for immediate use

**Use Cases:**
- Visual debugging of command execution
- Verify view state and element visibility
- Inspect UI state and properties
- Document model state for analysis
- Capture warnings or dialogs

## Error Handling

### Compilation Errors

If code has syntax errors, the server returns them:

```json
{
  "Success": false,
  "Output": "",
  "Error": "Compilation failed",
  "Diagnostics": [
    "(1,23): error CS1002: ; expected"
  ]
}
```

The server **continues listening** after errors, so you can send corrected code.

### Runtime Errors

If code compiles but fails during execution:

```json
{
  "Success": false,
  "Output": "Any output before the error\n",
  "Error": "System.NullReferenceException: Object reference not set...",
  "Diagnostics": []
}
```

## Revit .NET API Best Practices

### Transactions for Modifications
**CRITICAL**: Always use `Transaction` when performing document modifications:

```csharp
using (var trans = new Transaction(Doc, "Description"))
{
    trans.Start();
    // Your modifications here
    trans.Commit();
}
```

**When to use transactions:** Creating/modifying/deleting elements, changing parameters, adding to symbol tables.

### ExternalEvent for Background Operations
For operations initiated from background threads (like the server), use `ExternalEvent`:

```csharp
// Create once during API execution context
private ExternalEvent myEvent;
private MyHandler myHandler;

public void Initialize()
{
    myHandler = new MyHandler();
    myEvent = ExternalEvent.Create(myHandler);
}

// Raise from background thread
myEvent.Raise();
```

### FilteredElementCollector Patterns

**Get all elements by category:**
```csharp
var collector = new FilteredElementCollector(Doc);
var walls = collector.OfCategory(BuiltInCategory.OST_Walls)
    .WhereElementIsNotElementType()
    .ToElements();
```

**Get all of a specific class:**
```csharp
var collector = new FilteredElementCollector(Doc);
var levels = collector.OfClass(typeof(Level)).Cast<Level>();
```

### Output and Return Values

- **Use Console.WriteLine for Output**: All output should go through `Console.WriteLine()`
- **Handle Null Values**: Always check for null when working with Revit objects
- **Format Output Consistently**: Make output easy to parse programmatically
  ```csharp
  Console.WriteLine($"ELEMENT|{element.Id}|{element.Name}|{element.Category?.Name}");
  ```
- **Return Values**: Scripts can return values that will be appended to output
  ```csharp
  return collector.ToElements().Count();
  ```
- **Check Document State**: Verify active document exists before operations
  ```csharp
  if (Doc == null)
  {
      Console.WriteLine("No active document");
      return;
  }
  ```

### Parameter Setting and Unit Conversion

**CRITICAL**: When setting numeric parameters (Double storage type), always use `SetValueString()` instead of `Set()` for user-facing values:

```csharp
// CORRECT: Uses SetValueString - handles unit conversion automatically
param.SetValueString("100");  // In metric project: 100mm, in imperial: 100 feet

// WRONG: Uses Set directly - assumes internal units (feet)
param.Set(100);  // Always 100 feet, regardless of project units
```

**Why this matters:**
- Revit stores all numeric values internally in **imperial units** (feet, square feet, cubic feet)
- `SetValueString()` converts from display units to internal units based on project settings
- `Set()` assumes the value is already in internal units, causing incorrect conversions in metric projects
- Example: Setting "100" for a length in a metric project should be 100mm, not 30480mm (100 feet)

**Implementation pattern:**
```csharp
try
{
    param.SetValueString(userInput);  // Prefer this for user-facing values
}
catch
{
    if (double.TryParse(userInput, out double val))
        param.Set(val);  // Fallback for parameters that don't support SetValueString
}
```

## Integration with Claude Code

When Claude Code needs to query the Revit session:

1. **Read shared token**: Load token from `runtime/network/token`
2. **Discover active sessions**: Read network registry from `runtime/network/sessions`
3. **Construct query**: Write C# code to `/tmp/query-$(uuidgen).cs`
4. **Send request**: POST to `/roslyn` endpoint with `--data-binary` and authentication
5. **Parse JSON response**: Extract `Success`, `Output`, `Error`, and `Diagnostics`
6. **Iterate if needed**: If compilation fails, fix errors and retry

**Recommended Helper Function:**

Add this to your shell for simplified querying:

```bash
# Helper function for querying Revit
query_revit() {
    local query_file="/tmp/query-$(uuidgen).cs"
    echo "$1" > "$query_file"

    local TOKEN=$(cat runtime/network/token 2>/dev/null)
    local PORT=$(grep -v '^#' runtime/network/sessions 2>/dev/null | head -1 | cut -d',' -f2)

    if [ -z "$TOKEN" ] || [ -z "$PORT" ]; then
        echo "Error: Revit Ballet server not running or not accessible"
        return 1
    fi

    curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
      -H "X-Auth-Token: $TOKEN" \
      --data-binary "@$query_file" | jq -r '.Output // .Error'

    rm -f "$query_file"
}

# Usage example:
query_revit 'Console.WriteLine("Hello from Revit: " + Doc.Title);'
```

**Manual Query Workflow:**

```bash
# 1. Write your C# query to a file
cat > /tmp/query-$(uuidgen).cs << 'EOF'
var activeView = Doc.ActiveView;
Console.WriteLine("Current View: " + activeView.Name);
Console.WriteLine("View Type: " + activeView.ViewType);

var collector = new FilteredElementCollector(Doc, activeView.Id);
var elements = collector.WhereElementIsNotElementType().ToElements();
Console.WriteLine("Elements in view: " + elements.Count);
EOF

# 2. Get credentials and send request
TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)

curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  --data-binary @/tmp/query-*.cs | jq -r '.Output'
```

## Security Considerations

- **HTTPS with TLS 1.2/1.3**: All communication is encrypted using self-signed certificates
- **Localhost-only binding**: Servers only listen on 127.0.0.1 - not accessible from network
- **Shared token authentication**: 256-bit random token shared across all sessions
- **Token storage**: `%appdata%/revit-ballet/runtime/network/token` (plain text file)
- **Code execution**: Scripts have **full Revit API access** - can read, modify, and delete model data
- **Intended use**: Local development and automation only
- **Do not**:
  - Expose this service to untrusted networks
  - Share the authentication token with untrusted parties
  - Run Revit with elevated privileges when using the server

## Naming Conventions

**IMPORTANT**: Revit Ballet uses **PascalCase** (also called UpperCamelCase) for all naming conventions, following standard C# conventions.

This codebase was partially ported from AutoCAD Ballet which used kebab-case. All kebab-case naming has been migrated to PascalCase to maintain consistency with C# standards.

### Naming Standards:
- **Classes**: PascalCase (e.g., `TransformSelectedElements`, `DataGrid2SearchHistory`)
- **Methods**: PascalCase (e.g., `GetRuntimeFilePath`, `ProcessCommand`)
- **Properties**: PascalCase (e.g., `CurrentView`, `SelectedElements`)
- **Variables**: camelCase (e.g., `appData`, `searchHistory`)
- **File Names**: PascalCase (e.g., `Server.cs`, `SearchboxQueries/`)
- **Directory Names**: PascalCase (e.g., `runtime/SearchboxQueries/`)

### Migration from AutoCAD Ballet:
- ❌ `searchbox-queries` → ✅ `SearchboxQueries`
- ❌ `server-cert.pfx` → ✅ `ServerCert.pfx`
- ❌ `InvokeAddinCommand-history` → ✅ `InvokeAddinCommandHistory`

**Note**: The project name `revit-ballet` itself uses kebab-case as it's a product name/brand identity, not a code identifier.

## Build Limitations

**IMPORTANT**: Do NOT attempt to build the installer using `cd installer && dotnet build -c Release`.

Due to limitations of the underlying 9p filesystem in the isolated environment, builds may take excessive amounts of time (over 5 minutes or may hang indefinitely).

Instead:
- Ask the user to test the build and return results if needed
- Focus on code changes and let the user handle compilation
- If build verification is required, request the user run the build command in their environment

## Testing Across Revit Versions

When making changes that could affect compatibility, it's recommended (though not required) to test builds across all supported Revit versions:

```bash
cd ~/revit-ballet/commands && for i in {2026..2017}; do dotnet build -c Release -p:RevitYear=$i; done
```

This ensures consistency across Revit 2017-2026. Use this when:
- Modifying core functionality
- Changing Revit API interactions
- Making significant refactoring changes

## Building

```bash
# Build plugin (default 2026)
dotnet build commands/revit-ballet.csproj

# Build for specific version
dotnet build commands/revit-ballet.csproj -p:RevitYear=2024

# Build installer
dotnet build installer/installer.csproj
```

## Key Files

- `commands/revit-ballet.csproj` - Main plugin project
- `commands/RevitBallet.cs` - Extension application entry point (initializes server)
- `commands/Server.cs` - Auto-started Roslyn server for AI agents and network queries
- `commands/PathHelper.cs` - Runtime path management

## File Naming Conventions

- **Command Files**: PascalCase matching class name (e.g., `DataGrid2EditMode.cs`)
- **One Command Per File**: Complete implementation
- **Internal/Utility Files**: PascalCase (e.g., `Server.cs`, `PathHelper.cs`)
