using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

public partial class CopySelectedElementsAlongContainingGroupsByRooms
{
    // Simplified diagnostic configuration
    private bool enableDiagnostics = true; // Set to false to disable all diagnostics
    private bool saveDiagnosticsToFile = true; // Save to desktop

    // Focused diagnostic tracking
    private StringBuilder diagnosticLog = new StringBuilder();
    private Dictionary<string, int> summaryStats = new Dictionary<string, int>();

    // Initialize diagnostics
    private void InitializeDiagnostics()
    {
        if (!enableDiagnostics) return;

        diagnosticLog.Clear();
        summaryStats.Clear();

        diagnosticLog.AppendLine($"=== ROOM SIMILARITY DIAGNOSTICS - {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
        diagnosticLog.AppendLine($"Thresholds: Position={ROOM_POSITION_TOLERANCE}ft, Area={ROOM_AREA_TOLERANCE:P0}, MinRatio={ROOM_VALIDATION_THRESHOLD:P0}");
        diagnosticLog.AppendLine();
    }

    // Log selected elements summary
    private void LogSelectedElements(List<Element> selectedElements)
    {
        if (!enableDiagnostics) return;

        diagnosticLog.AppendLine($"SELECTED ELEMENTS: {selectedElements.Count}");
        foreach (Element elem in selectedElements)
        {
            string elemDesc = elem.Name ?? $"{elem.Category?.Name}";
            diagnosticLog.AppendLine($"  - {elemDesc} (ID: {elem.Id})");
        }
        diagnosticLog.AppendLine();
    }

    // Log room validation failure - this is the key diagnostic
    private void LogRoomValidationFailure(Room room, Group candidateGroup,
        List<Group> sameTypeGroups, DetailedValidationResult validationResult, Document doc)
    {
        if (!enableDiagnostics) return;

        GroupType gt = doc.GetElement(candidateGroup.GetTypeId()) as GroupType;
        string groupTypeName = gt?.Name ?? "Unknown";

        diagnosticLog.AppendLine($"ROOM VALIDATION FAILED:");
        diagnosticLog.AppendLine($"  Room: {room.Name ?? "Unnamed"} #{room.Number} (Area: {room.Area:F0} sqft)");
        diagnosticLog.AppendLine($"  Group Type: {groupTypeName} (Instance ID: {candidateGroup.Id})");
        diagnosticLog.AppendLine($"  Result: {validationResult.InstancesWithRoom}/{validationResult.InstancesChecked} instances have similar room");
        diagnosticLog.AppendLine($"  Reason: {validationResult.PrimaryFailureReason}");

        // Show first few mismatches for debugging
        if (validationResult.MismatchDetails.Count > 0)
        {
            diagnosticLog.AppendLine("  Details:");
            foreach (var detail in validationResult.MismatchDetails.Take(3))
            {
                diagnosticLog.AppendLine($"    - {detail}");
            }
        }
        diagnosticLog.AppendLine();
    }

    // Log room validation success
    private void LogRoomValidationSuccess(Room room, Group candidateGroup,
        DetailedValidationResult validationResult, Document doc)
    {
        if (!enableDiagnostics) return;

        GroupType gt = doc.GetElement(candidateGroup.GetTypeId()) as GroupType;
        string groupTypeName = gt?.Name ?? "Unknown";

        // Only log brief success info
        diagnosticLog.AppendLine($"ROOM VALIDATED: {room.Name ?? "Unnamed"} #{room.Number} for {groupTypeName} ({validationResult.InstancesWithRoom}/{validationResult.InstancesChecked} matches)");
    }

    // Log element containment check
    private void LogElementContainmentCheck(Element element, Group group, bool isContained, string roomName, Document doc)
    {
        if (!enableDiagnostics) return;

        // Only log failures for element containment
        if (!isContained)
        {
            GroupType gt = doc.GetElement(group.GetTypeId()) as GroupType;
            string groupTypeName = gt?.Name ?? "Unknown";
            string elemDesc = element.Name ?? $"{element.Category?.Name}";

            diagnosticLog.AppendLine($"ELEMENT NOT IN ROOM: {elemDesc} not found in any room of {groupTypeName} (Group ID: {group.Id})");
        }
    }

    // Log final summary
    private void LogFinalSummary(int selectedCount, int foundInGroups, int totalCopied)
    {
        if (!enableDiagnostics) return;

        diagnosticLog.AppendLine();
        diagnosticLog.AppendLine("=== SUMMARY ===");
        diagnosticLog.AppendLine($"Selected elements: {selectedCount}");
        diagnosticLog.AppendLine($"Elements found in groups: {foundInGroups}");
        diagnosticLog.AppendLine($"Elements copied: {totalCopied}");

        if (summaryStats.Count > 0)
        {
            diagnosticLog.AppendLine();
            diagnosticLog.AppendLine("Key Issues:");
            foreach (var stat in summaryStats.OrderByDescending(x => x.Value))
            {
                diagnosticLog.AppendLine($"  - {stat.Key}: {stat.Value}");
            }
        }
    }

    // Track summary statistic
    private void TrackStat(string statName)
    {
        if (!enableDiagnostics) return;

        if (!summaryStats.ContainsKey(statName))
            summaryStats[statName] = 0;
        summaryStats[statName]++;
    }

    // Save diagnostics to file (only if there are issues)
    private void SaveDiagnostics(bool forceWrite = false)
    {
        if (!enableDiagnostics || !saveDiagnosticsToFile) return;

        // Only save if there are actual issues or forced
        if (!forceWrite && summaryStats.Count == 0 && diagnosticLog.Length < 500)
            return;

        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string fileName = $"RevitRoomDiagnostics_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
        string filePath = System.IO.Path.Combine(desktopPath, fileName);

        try
        {
            System.IO.File.WriteAllText(filePath, diagnosticLog.ToString());

            if (forceWrite || summaryStats.Count > 0)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Diagnostics Saved",
                    $"Room validation diagnostics saved to:\n{filePath}");
            }
        }
        catch (Exception ex)
        {
            Autodesk.Revit.UI.TaskDialog.Show("Diagnostic Error",
                $"Could not save diagnostics: {ex.Message}");
        }
    }

    // Quick diagnostic for why elements weren't found in groups
    private void DiagnoseElementsNotInGroups(List<Element> selectedElements,
        Dictionary<ElementId, List<Group>> elementsInGroups, Document doc)
    {
        if (!enableDiagnostics) return;

        foreach (Element elem in selectedElements)
        {
            if (!elementsInGroups.ContainsKey(elem.Id) || elementsInGroups[elem.Id].Count == 0)
            {
                string elemDesc = elem.Name ?? $"{elem.Category?.Name}";
                diagnosticLog.AppendLine($"ELEMENT NOT IN ANY GROUP: {elemDesc} (ID: {elem.Id})");

                // Check if element is in any room at all
                LocationPoint locPoint = elem.Location as LocationPoint;
                if (locPoint != null)
                {
                    XYZ point = locPoint.Point;

                    // Quick check: is this element in ANY room?
                    bool foundInAnyRoom = false;
                    foreach (var kvp in _roomDataCache)
                    {
                        Room room = doc.GetElement(kvp.Key) as Room;
                        if (room != null && room.IsPointInRoom(point))
                        {
                            foundInAnyRoom = true;
                            diagnosticLog.AppendLine($"  Element IS in room: {room.Name ?? room.Number}");
                            diagnosticLog.AppendLine($"  But room not associated with any group");
                            TrackStat("Elements in rooms not associated with groups");
                            break;
                        }
                    }

                    if (!foundInAnyRoom)
                    {
                        diagnosticLog.AppendLine($"  Element NOT in any room");
                        TrackStat("Elements not in any room");
                    }
                }
                diagnosticLog.AppendLine();
            }
        }
    }
}
