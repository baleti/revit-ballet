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
- **ALL runtime data MUST be stored in `runtime/` subdirectory** - includes logs, network data, screenshots, diagnostics

### Diagnostic Logging Convention

**When to add diagnostics:** For non-trivial operations where bugs could arise from unit conversion, coordinate transformations, or API behavioral differences.

**Location:** All diagnostic files MUST be saved to `runtime/diagnostics/` directory.

**Naming:** Use descriptive names with timestamps: `{OperationName}-{yyyyMMdd-HHmmss-fff}.txt`

**Content Format:**
```csharp
// Create diagnostic file
string diagnosticPath = System.IO.Path.Combine(
    RevitBallet.Commands.PathHelper.RuntimeDirectory,
    "diagnostics",
    $"MyOperation-{System.DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");

var diagnosticLines = new List<string>();
diagnosticLines.Add($"=== Operation Name at {System.DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ===");
diagnosticLines.Add($"Input Values: ...");
diagnosticLines.Add($"Intermediate Results: ...");
diagnosticLines.Add($"Final Results: ...");

// Always create directory before writing
System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
System.IO.File.WriteAllLines(diagnosticPath, diagnosticLines);
```

**Key principles:**
- Include timestamps for correlation with user actions
- Log both internal units (feet) AND display units (mm/m) for coordinates
- Log intermediate calculation steps, not just final results
- Include element IDs for traceability
- Log both input values and verified output values

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

**NOTE:** Use the file-based approach from "Querying and Testing Revit Sessions" section. Query examples:

```csharp
// Get all levels
var levels = new FilteredElementCollector(Doc).OfClass(typeof(Level)).Cast<Level>();
Console.WriteLine("Total Levels: " + levels.Count());

// Count elements by category
var elements = new FilteredElementCollector(Doc).WhereElementIsNotElementType();
var counts = elements.GroupBy(e => e.Category?.Name ?? "None").OrderByDescending(g => g.Count());

// Check for unplaced rooms
var rooms = new FilteredElementCollector(Doc).OfCategory(BuiltInCategory.OST_Rooms).Cast<Room>();
var unplaced = rooms.Where(r => r.Area == 0 || r.Location == null);
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

### Unit Conversion (Coordinates and Parameters)

**CRITICAL**: Revit stores ALL values internally in **imperial units** (feet, sq ft, cu ft) regardless of project settings.

**Parameters:** Use `SetValueString()` for auto-conversion; `Set()` assumes feet.
**Coordinates:** Use `UnitUtils.ConvertFromInternalUnits()` for display, `UnitUtils.ConvertToInternalUnits()` for input.

```csharp
// Get project units
Units units = Doc.GetUnits();
ForgeTypeId unitTypeId = units.GetFormatOptions(SpecTypeId.Length).GetUnitTypeId();

// Display: internal → user units
double display = UnitUtils.ConvertFromInternalUnits(internalFeet, unitTypeId);

// Input: user units → internal
double internal = UnitUtils.ConvertToInternalUnits(userInput, unitTypeId);
```

## Integration with Claude Code

Use the file-based query pattern from "Querying and Testing Revit Sessions" section. Workflow:
1. Write C# code to `/tmp/query-$(uuidgen).cs`
2. POST to `/roslyn` with token and `--data-binary`
3. Parse JSON response for `Success`, `Output`, `Error`, `Diagnostics`

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

## Command Scopes (View, Document, Session, Network)

Revit Ballet supports multiple command scopes, similar to AutoCAD Ballet's architecture. Commands use suffix naming to indicate their scope:

1. **View Scope (`InViews`)** - Commands operate on the current active view
   - Example: `FilterSelectedInViews`, `SelectByCategoriesInViews`
2. **Document Scope (`InDocument`)** - Commands operate on the current active document/project
   - Example: `FilterSelectedInDocument`, `SelectByCategoriesInDocument`
   - **Note**: Previously used `InProject` suffix, renamed to `InDocument` for consistency with AutoCAD Ballet
3. **Session Scope (`InSession`)** - Commands operate across **all open documents** in the Revit process
   - To be implemented for cross-document workflows
4. **Network Scope (`InNetwork`)** - Commands operate across **all Revit sessions** in the peer-to-peer network
   - To be implemented for multi-session coordination

### Multi-Document (Session Scope) Operations

**VERIFIED**: Revit API fully supports reading and writing across multiple open documents without switching the active document.

**Key Findings:**
- `UIApp.Application.Documents` provides access to all open documents in the session
- `FilteredElementCollector(doc)` can query any document, not just the active one
- `Transaction(doc, "...")` can modify any document, not just the active one
- No need to switch active document to perform operations
- Linked documents (`doc.IsLinked == true`) should be skipped for modification operations

**Example - Session Scope Pattern:**
```csharp
var app = commandData.Application.Application;
var activeDoc = commandData.Application.ActiveUIDocument.Document;

foreach (Document doc in app.Documents)
{
    // Skip linked documents (read-only references)
    if (doc.IsLinked) continue;

    // Read from any document
    var collector = new FilteredElementCollector(doc);
    var elements = collector.WhereElementIsNotElementType().ToElements();

    // Write to any document
    using (var trans = new Transaction(doc, "Session Scope Operation"))
    {
        trans.Start();
        // Modify elements in this document
        trans.Commit();
    }
}
```

**Use Cases:**
- Batch operations across all open projects
- Cross-project analysis and reporting
- Synchronizing settings across documents
- Quality checks spanning multiple models
- Collecting elements by category from all open documents

## Selection Mode Manager

**CRITICAL**: All commands that work with element selection MUST use `SelectionModeManager` extension methods instead of direct Revit UI selection APIs.

Revit Ballet supports two selection modes:
1. **RevitUI** - Standard Revit selection (blue highlight in viewport)
2. **SelectionSet** - Selection stored in a persistent SelectionFilterElement named "temp"

### Why Two Modes?

SelectionSet mode enables:
- Persistent selection across view switches
- Selection of elements not visible in current view
- Selection of linked model elements (stored separately in file)
- More robust workflows for complex multi-view operations

### Implementation Requirements

**ALWAYS use these extension methods on `UIDocument`:**

```csharp
// Get current selection
ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();

// Set selection (replace)
uidoc.SetSelectionIds(elementIds);

// Add to existing selection
uidoc.AddToSelection(elementIds);

// Get selection including linked elements
IList<Reference> references = uidoc.GetReferences();

// Set selection including linked elements
uidoc.SetReferences(references);
```

**NEVER use direct Revit selection APIs:**
```csharp
// ❌ WRONG - bypasses SelectionModeManager
uidoc.Selection.GetElementIds();
uidoc.Selection.SetElementIds(elementIds);

// ✅ CORRECT - uses SelectionModeManager
uidoc.GetSelectionIds();
uidoc.SetSelectionIds(elementIds);
```

### Implementation Files

- `commands/SelectionModeManager.cs` - Core selection mode system
- Mode persisted in `runtime/SelectionMode` file
- Linked references stored in `runtime/SelectionSet-LinkedModelReferences`
- Temp selection set created automatically as needed

### Example Command Pattern

```csharp
[Transaction(TransactionMode.Manual)]
public class MyCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get current selection using SelectionModeManager
        ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();

        // ... command logic ...

        // Update selection using SelectionModeManager
        uidoc.SetSelectionIds(newSelection);

        return Result.Succeeded;
    }
}
```

## DataGrid Automatic Editable Columns

**IMPORTANT**: DataGrid columns are automatically editable based on column NAME - commands just create dictionaries, editing happens automatically.

### How It Works

The DataGrid system uses a **column handler registry** that recognizes column names and automatically enables editing with validation:

```csharp
// Command just creates dictionaries with named columns
var data = new Dictionary<string, object>
{
    ["Name"] = element.Name,
    ["Type Name"] = typeElement?.Name,
    ["Family"] = familySymbol?.FamilyName,
    ["Comments"] = commentsParam?.AsString(),
    ["Mark"] = markParam?.AsString()
};

// Set UIDocument once to enable automatic edit application
CustomGUIs.SetCurrentUIDocument(uidoc);

// Show grid - editing is AUTOMATIC
var selected = CustomGUIs.DataGrid(gridData, propertyNames, false);

// ✅ Edit detection: AUTOMATIC (based on column names)
// ✅ Validation: AUTOMATIC (invalid values rejected)
// ✅ ElementIdObject: AUTOMATIC (injected if missing)
// ✅ Apply edits: AUTOMATIC (when grid closes)
```

### Standard Editable Columns (Pre-Registered)

These column names are automatically editable with validation:

| Column Name | Edits What | Validation | Works On |
|-------------|------------|------------|----------|
| **"Family"** | Family.Name | No empty, no special chars, trimmed | Instances or FamilySymbol types |
| **"Type Name"** | ElementType.Name | No empty, no special chars, trimmed | Instances or ElementType objects |
| **"Comments"** | ALL_MODEL_INSTANCE_COMMENTS | Max 1024 chars | Elements with parameter |
| **"Mark"** | ALL_MODEL_MARK | No special chars, max 256 chars | Elements with parameter |

**Key Point**: Column names are **case-insensitive** ("Family", "family", "FAMILY" all work).

### Requirements for Automatic Editing

1. **Column name must match** a registered handler (e.g., "Family", "Type Name")
2. **`SetCurrentUIDocument(uidoc)` must be called** before showing DataGrid
3. **Row must have `ElementIdObject`** or `Id` field (auto-injected if missing)

That's it! No other code needed.

### Adding Custom Editable Columns

Register new handlers in `RevitBallet.cs` OnStartup or command initialization:

```csharp
CustomGUIs.ColumnHandlerRegistry.Register(new CustomGUIs.ColumnHandler
{
    ColumnName = "Level",
    IsEditable = true,
    Description = "Element level",

    // Validation (optional)
    Validator = CustomGUIs.ColumnValidators.NotEmpty,

    // How to read value
    Getter = (elem, doc) => {
        Parameter levelParam = elem.get_Parameter(BuiltInParameter.LEVEL_PARAM);
        return levelParam?.AsValueString() ?? "";
    },

    // How to write value
    Setter = (elem, doc, newValue) => {
        string levelName = newValue?.ToString() ?? "";
        // Find level by name and set parameter
        // ... implementation ...
        return true; // or false if failed
    }
});
```

Now **ANY command** that creates a "Level" column gets automatic editing!

### Implementation Files

- `DataGrid.ColumnHandlers.cs` - Handler registry and standard handlers
- `DataGrid.Validation.cs` - Validation infrastructure
- `DataGrid.EditMode.cs` - Auto-detection integration
- `DataGrid.EditApply.cs` - Handler-based application
- `DataGrid.Main.cs` - Automatic edit application on close

See `IMPLEMENTATION_SUMMARY.md` in `commands/DataGrid/` for full documentation.

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
