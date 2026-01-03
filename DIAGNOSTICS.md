# Diagnostic System Documentation

## What Gets Logged and When

### **Automatic Diagnostics for SwitchToLastView**

Every time `SwitchToLastView` runs, **TWO** diagnostic files are created:

#### 1. **Command Execution Diagnostic**
**File:** `runtime/diagnostics/SwitchToLastView-{timestamp}.txt`

**Created:** Automatically via `CommandDiagnostics.StartCommand()` using statement

**Contains:**
- ✅ **START section:**
  - Timestamp when command started
  - Call stack (shows if invoked via InvokeAddinCommand)
  - Transaction check: Open transactions at START (⚠ warns if any found)
  - Active document and view

- ✅ **Execution log:**
  - Timestamped entries for key operations
  - `OpenDocumentFile` calls
  - `OpenAndActivateDocument` calls
  - Success/failure messages
  - Errors with full exception details

- ✅ **END section:**
  - Command duration in milliseconds
  - Transaction check: Open transactions at END (⚠⚠⚠ CRITICAL if any found)
  - "Clean exit" confirmation or rollback warning

**Example output:**
```
=== COMMAND START: SwitchToLastView at 2026-01-02 22:15:30.123 ===
Call Stack Depth: 2
Called From: SwitchToLastView → InvokeAddinCommand

⚠ WARNING: Open transactions detected at START:
  • Document 'Project1.rvt' has an open transaction: ...

✓ No open transactions at START

Active Document: Project1.rvt
Active View: Level 1

[0ms] Retrieved 20 entries from view history
[150ms] Calling OpenDocumentFile for Project2.rvt
[300ms] Calling OpenAndActivateDocument for Project2.rvt
[450ms] Document activated: Project2.rvt
[600ms] SUCCESS: Switched to Project2.rvt | Sheet A101

=== COMMAND END: SwitchToLastView ===
Duration: 650ms

✓ No open transactions at END (clean exit)
```

#### 2. **View Switching Details Diagnostic**
**File:** `runtime/diagnostics/SwitchToLastView-{timestamp}.txt` (legacy, different timestamp)

**Created:** Via existing `WriteDiagnostic()` method

**Contains:**
- View history entries
- Search logic for finding previous view
- Detailed step-by-step view switching logic

### **When Diagnostics Detect Issues**

#### **Open Transaction at START**
- ⚠ **WARNING logged** but command continues
- Indicates another command left a transaction open
- Revit will auto-rollback that other command's changes

#### **Open Transaction at END**
- ⚠⚠⚠ **CRITICAL logged**
- THIS command left a transaction open
- Revit will show error: "A transaction or sub-transaction was opened but not closed"
- All uncommitted changes from THIS command will be rolled back

#### **Call Stack Shows Invocation Path**
- If you see: `Called From: SwitchToLastView → InvokeAddinCommand`
- This means you invoked SwitchToLastView via InvokeAddinCommand
- Helps correlate crashes with keyboard shortcuts or custom invocations

### **Other Commands - Not Yet Instrumented**

The following commands also switch documents but **don't have diagnostics yet:**
- `SwitchViewInSession.cs`
- `SwitchViewByHistoryInSession.cs`
- `SwitchDocument.cs`
- `OpenViewsInSession.cs`
- `OpenSheetsInSession.cs`
- `OpenPreviousViewsInSession.cs`

**Recommendation:** Add the same diagnostic pattern to these commands.

## How to Add Diagnostics to Other Commands

```csharp
public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
{
    UIApplication uiApp = commandData.Application;

    // Wrap entire command in diagnostic session
    using (var diagnostics = CommandDiagnostics.StartCommand("YourCommandName", uiApp))
    {
        try
        {
            // Your command logic here
            diagnostics.Log("Doing something important");

            // ...

            diagnostics.Log("SUCCESS: Operation completed");
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            diagnostics.LogError("Operation failed", ex);
            message = ex.Message;
            return Result.Failed;
        }
    } // Diagnostic session auto-closes and writes file
}
```

## Answering Your Questions

### Q: "When will diagnostic files be saved?"

**A:** Every time `SwitchToLastView` runs, whether successful or not. The `using` statement ensures the diagnostic is written even if an exception occurs.

### Q: "Will we detect dangling transactions?"

**A:** Yes! The diagnostic checks **twice**:
1. At START - warns if previous command left one open
2. At END - CRITICAL warning if THIS command left one open

### Q: "What about other commands switching views cross-session?"

**A:** They currently have the `OpenDocumentFile` fix but NO diagnostic logging yet. You should add the same pattern to all of them.

### Q: "Will we detect if SwitchToLastView was called from InvokeAddinCommand?"

**A:** Yes! The diagnostic shows the full call stack:
```
Call Stack Depth: 2
Called From: SwitchToLastView → InvokeAddinCommand
```

This directly answers: "what crashed my revit?" by showing the invocation chain.

## File Cleanup

Diagnostic files accumulate in `runtime/diagnostics/`. You may want to periodically clean old files or implement auto-cleanup like the screenshot feature does (keeps last N files).
