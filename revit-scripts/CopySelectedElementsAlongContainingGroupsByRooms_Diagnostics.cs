using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

public partial class CopySelectedElementsAlongContainingGroupsByRooms
{
    private bool enableDiagnostics = true;
    private bool saveDiagnosticsToFile = true;
    private StringBuilder diagnosticLog = new StringBuilder();
    
    // Timing tracking
    private Stopwatch globalStopwatch = new Stopwatch();
    private Dictionary<string, TimeSpan> timingResults = new Dictionary<string, TimeSpan>();
    private Dictionary<string, int> callCounts = new Dictionary<string, int>();
    private Stack<(string name, Stopwatch sw)> timingStack = new Stack<(string, Stopwatch)>();

    private void InitializeDiagnostics()
    {
        if (!enableDiagnostics) return;
        diagnosticLog.Clear();
        timingResults.Clear();
        callCounts.Clear();
        globalStopwatch.Restart();
        
        diagnosticLog.AppendLine($"=== PERFORMANCE TIMING ANALYSIS - {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
        diagnosticLog.AppendLine();
    }

    // Call this at the start of any method you want to time
    public void StartTiming(string operationName)
    {
        if (!enableDiagnostics) return;
        
        var sw = new Stopwatch();
        sw.Start();
        timingStack.Push((operationName, sw));
        
        if (!callCounts.ContainsKey(operationName))
            callCounts[operationName] = 0;
        callCounts[operationName]++;
    }

    // Call this at the end of the method
    public void EndTiming(string operationName)
    {
        if (!enableDiagnostics) return;
        
        if (timingStack.Count > 0 && timingStack.Peek().name == operationName)
        {
            var (name, sw) = timingStack.Pop();
            sw.Stop();
            
            if (!timingResults.ContainsKey(operationName))
                timingResults[operationName] = TimeSpan.Zero;
            timingResults[operationName] = timingResults[operationName].Add(sw.Elapsed);
        }
    }

    // Modified LogSelectedElements to add timing
    private void LogSelectedElements(List<Element> selectedElements)
    {
        if (!enableDiagnostics) return;
        
        StartTiming("LogSelectedElements");
        
        diagnosticLog.AppendLine($"Selected elements: {selectedElements.Count}");
        diagnosticLog.AppendLine($"Room cache size: {_roomDataCache?.Count ?? 0}");
        diagnosticLog.AppendLine();
        
        EndTiming("LogSelectedElements");
    }

    // Add timing hooks to major operations (call from other files)
    public void LogTimingCheckpoint(string checkpoint)
    {
        if (!enableDiagnostics) return;
        
        TimeSpan elapsed = globalStopwatch.Elapsed;
        diagnosticLog.AppendLine($"[{elapsed:mm\\:ss\\.fff}] Checkpoint: {checkpoint}");
    }

    // Override the existing stubs to track specific operations
    private void LogRoomValidationFailure(Room room, Group candidateGroup,
        List<Group> sameTypeGroups, DetailedValidationResult validationResult, Document doc)
    {
        if (!enableDiagnostics) return;
        // Track validation failures
        StartTiming("RoomValidation");
        // ... existing logic if any ...
        EndTiming("RoomValidation");
    }

    private void LogRoomValidationSuccess(Room room, Group candidateGroup,
        DetailedValidationResult validationResult, Document doc)
    {
        if (!enableDiagnostics) return;
        StartTiming("RoomValidation");
        // ... existing logic if any ...
        EndTiming("RoomValidation");
    }

    private void LogElementContainmentCheck(Element element, Group group, bool isContained, string roomName, Document doc)
    {
        if (!enableDiagnostics) return;
        StartTiming("ElementContainmentCheck");
        // Track containment checks
        EndTiming("ElementContainmentCheck");
    }

    private void LogFinalSummary(int selectedCount, int foundInGroups, int totalCopied)
    {
        if (!enableDiagnostics) return;
        
        globalStopwatch.Stop();
        
        diagnosticLog.AppendLine($"\n=== TIMING SUMMARY ===");
        diagnosticLog.AppendLine($"Total execution time: {globalStopwatch.Elapsed:mm\\:ss\\.fff}");
        diagnosticLog.AppendLine();
        
        // Sort operations by total time
        var sortedTimings = timingResults.OrderByDescending(kvp => kvp.Value).ToList();
        
        diagnosticLog.AppendLine("Operation Timings (sorted by total time):");
        diagnosticLog.AppendLine("─────────────────────────────────────────");
        
        foreach (var kvp in sortedTimings)
        {
            string operation = kvp.Key;
            TimeSpan totalTime = kvp.Value;
            int count = callCounts.ContainsKey(operation) ? callCounts[operation] : 1;
            TimeSpan avgTime = TimeSpan.FromMilliseconds(totalTime.TotalMilliseconds / count);
            
            diagnosticLog.AppendLine($"{operation,-40} │ Total: {totalTime:mm\\:ss\\.fff} │ Calls: {count,6} │ Avg: {avgTime.TotalMilliseconds,8:F2}ms");
        }
        
        diagnosticLog.AppendLine();
        diagnosticLog.AppendLine("=== RESULTS ===");
        diagnosticLog.AppendLine($"Elements selected: {selectedCount}");
        diagnosticLog.AppendLine($"Elements in groups: {foundInGroups}");
        diagnosticLog.AppendLine($"Elements copied: {totalCopied}");
        
        // Identify bottlenecks
        diagnosticLog.AppendLine();
        diagnosticLog.AppendLine("=== BOTTLENECK ANALYSIS ===");
        
        var topBottlenecks = sortedTimings.Take(5);
        double totalMs = globalStopwatch.Elapsed.TotalMilliseconds;
        
        foreach (var kvp in topBottlenecks)
        {
            double percentage = (kvp.Value.TotalMilliseconds / totalMs) * 100;
            diagnosticLog.AppendLine($"• {kvp.Key}: {percentage:F1}% of total time");
        }
    }

    // Stub implementations
    private void TestFloorDetection(Document doc, XYZ testPoint) { }
    private void TestRoomContainmentWithF2F(Element element, XYZ point) { }
    private void DiagnoseElementsNotInGroups(List<Element> selectedElements,
        Dictionary<ElementId, List<Group>> elementsInGroups, Document doc) { }
    private void TrackStat(string statName) { }
    private void LogNearbyRoomsForElement(Element element, Document doc) { }
    private void LogFloorToFloorCheck(XYZ point, RoomData roomData, Document doc) { }

    private void SaveDiagnostics(bool forceWrite = false)
    {
        if (!enableDiagnostics || !saveDiagnosticsToFile) return;

        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string fileName = $"RevitPerformance_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
        string filePath = System.IO.Path.Combine(desktopPath, fileName);

        try
        {
            System.IO.File.WriteAllText(filePath, diagnosticLog.ToString());
            if (forceWrite)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Performance Report", 
                    $"Execution time: {globalStopwatch.Elapsed:mm\\:ss}\nReport saved to:\n{filePath}");
            }
        }
        catch { }
    }
}
