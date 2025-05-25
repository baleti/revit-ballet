using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO; // Required for file operations
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

// Assuming CustomGUIs namespace is available
// using YourNamespaceContainingCustomGUIs;

namespace FilterDoorsWithWallOffsets
{
    [Transaction(TransactionMode.Manual)]
    public class FilterDoorsWithWallOffsets : IExternalCommand
    {
        // Updated configuration file path and name
        private const string ConfigFileName = "FilterDoorsWithWallOffsets"; // File name itself
        private const string ConfigKeyDrawDimensions = "draw_dimensions";
        private const string ConfigKeyDiagnostics = "diagnostics";
        private static readonly string ConfigFolderPath = Path.Combine( // Parent folder
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "revit-scripts");
        private static readonly string ConfigFilePath = Path.Combine(ConfigFolderPath, ConfigFileName); // Full path to the file

        private static void EnsureConfigFileExists()
        {
            try
            {
                if (!Directory.Exists(ConfigFolderPath))
                {
                    Directory.CreateDirectory(ConfigFolderPath);
                }

                if (!File.Exists(ConfigFilePath))
                {
                    File.WriteAllText(ConfigFilePath, $"{ConfigKeyDrawDimensions}: false\n{ConfigKeyDiagnostics}: false\n");
                }
                else
                {
                    // Check if diagnostics option exists, if not add it
                    var lines = File.ReadAllLines(ConfigFilePath).ToList();
                    bool hasDiagnostics = lines.Any(line => line.Trim().StartsWith(ConfigKeyDiagnostics, StringComparison.OrdinalIgnoreCase));
                    
                    if (!hasDiagnostics)
                    {
                        lines.Add($"{ConfigKeyDiagnostics}: false");
                        File.WriteAllLines(ConfigFilePath, lines);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error ensuring config file exists: {ex.Message}");
            }
        }

        private static bool GetConfigValue(string key)
        {
            EnsureConfigFileExists();
            try
            {
                var lines = File.ReadAllLines(ConfigFilePath);
                foreach (var line in lines)
                {
                    if (line.Trim().StartsWith(key, StringComparison.OrdinalIgnoreCase))
                    {
                        var parts = line.Split(':');
                        if (parts.Length > 1)
                        {
                            string value = parts[1].Trim().ToLowerInvariant();
                            return value == "true" || value == "1";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error reading config file: {ex.Message}");
            }
            return false; // Default to false if not found or error
        }

        private static bool ShouldDrawDimensions()
        {
            return GetConfigValue(ConfigKeyDrawDimensions);
        }

        private static bool ShouldShowDiagnostics()
        {
            return GetConfigValue(ConfigKeyDiagnostics);
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            bool drawDimensionsEnabled = ShouldDrawDimensions();
            bool diagnosticsEnabled = ShouldShowDiagnostics();

            try
            {
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("Error", "Please select one or more doors before running this command.");
                    return Result.Failed;
                }

                List<FamilyInstance> selectedDoors = new List<FamilyInstance>();
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem is FamilyInstance fi && fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
                    {
                        selectedDoors.Add(fi);
                    }
                }

                if (selectedDoors.Count == 0)
                {
                    TaskDialog.Show("Error", "No doors found in the current selection. Please select one or more doors.");
                    return Result.Failed;
                }

                var doorsByLevel = selectedDoors
                    .Where(d => d.LevelId != null && d.LevelId != ElementId.InvalidElementId)
                    .GroupBy(d => d.LevelId)
                    .ToDictionary(g => g.Key, g => g.ToList());

                Dictionary<ElementId, List<WallData>> wallsByLevel = new Dictionary<ElementId, List<WallData>>();
                foreach (var levelGroup in doorsByLevel)
                {
                    ElementLevelFilter levelFilter = new ElementLevelFilter(levelGroup.Key);
                    var levelWalls = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .OfClass(typeof(Wall))
                        .WherePasses(levelFilter)
                        .Cast<Wall>()
                        .Where(w => w.Location is LocationCurve)
                        .Select(w => new WallData
                        {
                            WallId = w.Id,
                            Wall = w,
                            Curve = (w.Location as LocationCurve).Curve,
                            BoundingBox = w.get_BoundingBox(null)
                        })
                        .ToList();
                    wallsByLevel[levelGroup.Key] = levelWalls;
                }

                var wallsWithoutLevel = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(Wall))
                    .Cast<Wall>()
                    .Where(w => (w.LevelId == null || w.LevelId == ElementId.InvalidElementId)
                        && w.Location is LocationCurve)
                    .Select(w => new WallData
                    {
                        WallId = w.Id,
                        Wall = w,
                        Curve = (w.Location as LocationCurve).Curve,
                        BoundingBox = w.get_BoundingBox(null)
                    })
                    .ToList();

                ConcurrentBag<DoorProcessingResult> doorResults = new ConcurrentBag<DoorProcessingResult>();
                ConcurrentBag<List<WallDiagnosticInfo>> allDiagnostics = new ConcurrentBag<List<WallDiagnosticInfo>>();

                Parallel.ForEach(selectedDoors, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, doorInst =>
                {
                    DoorProcessingResult doorResult = new DoorProcessingResult
                    {
                        Door = doorInst,
                    };

                    List<WallDiagnosticInfo> doorDiagnostics = new List<WallDiagnosticInfo>();

                    try
                    {
                        Wall hostWall = doorInst.Host as Wall;
                        if (hostWall == null)
                        {
                            doorResult.NoHostWall = true;
                        }
                        else
                        {
                            DoorOrientation doorOrientation = WallFinding.GetDoorOrientation(doorInst, hostWall);
                            List<WallData> wallsToCheck = new List<WallData>();
                            if (doorInst.LevelId != null && wallsByLevel.ContainsKey(doorInst.LevelId))
                            {
                                wallsToCheck.AddRange(wallsByLevel[doorInst.LevelId]);
                            }
                            wallsToCheck.AddRange(wallsWithoutLevel);

                            if (diagnosticsEnabled)
                            {
                                doorResult.AdjacentWalls = WallFinding.FindAdjacentWallsWithDiagnostics(
                                    doorInst, hostWall, doorOrientation, wallsToCheck, doc, out doorDiagnostics);
                            }
                            else
                            {
                                doorResult.AdjacentWalls = WallFinding.FindAdjacentWallsParallel(
                                    doorInst, hostWall, doorOrientation, wallsToCheck, doc);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        doorResult.Error = ex.Message;
                    }
                    doorResults.Add(doorResult);
                    
                    if (diagnosticsEnabled)
                    {
                        allDiagnostics.Add(doorDiagnostics);
                    }
                });

                // Always calculate distances for all doors, regardless of whether we're drawing dimensions
                Dictionary<int, List<DimensionInfo>> doorDistances = new Dictionary<int, List<DimensionInfo>>();

                // We need a transaction to create temporary dimensions for accurate measurement
                using (Transaction tx = new Transaction(doc, drawDimensionsEnabled ? "Create Door-Wall Dimensions" : "Calculate Door Distances"))
                {
                    tx.Start();

                    foreach (var result in doorResults.Where(r => !r.NoHostWall && string.IsNullOrEmpty(r.Error) && r.AdjacentWalls.Any()))
                    {
                        try
                        {
                            List<DimensionInfo> calculatedDistances = Dimensioning.CalculateDistancesForDoor(
                                doc, uidoc, result.Door, result.AdjacentWalls);
                            if (calculatedDistances.Any())
                            {
                                doorDistances[result.Door.Id.IntegerValue] = calculatedDistances;
                            }
                        }
                        catch (Exception exCalc)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error calculating distances for door {result.Door.Id}: {exCalc.Message}");
                            continue;
                        }
                    }

                    // If dimensions should be drawn, create them now
                    if (drawDimensionsEnabled)
                    {
                        foreach (var kvp in doorDistances)
                        {
                            try
                            {
                                // Create dimensions based on the calculated distances
                                Dimensioning.CreateDimensionsFromCalculatedDistances(
                                    doc, uidoc, kvp.Value);
                            }
                            catch (Exception exDim)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error creating dimension for door {kvp.Key}: {exDim.Message}");
                                continue;
                            }
                        }
                        tx.Commit();
                    }
                    else
                    {
                        // Roll back the transaction so temporary dimensions are not kept
                        tx.RollBack();
                    }
                }

                // Show diagnostics if enabled
                if (diagnosticsEnabled && allDiagnostics.Any())
                {
                    var diagnosticsList = allDiagnostics.ToList();
                    for (int i = 0; i < selectedDoors.Count && i < diagnosticsList.Count; i++)
                    {
                        if (diagnosticsList[i].Any())
                        {
                            Diagnostics.ShowDiagnosticsDialog(diagnosticsList[i], selectedDoors[i]);
                        }
                    }
                }

                ShowResults(doc, uidoc, selectedDoors, doorResults, doorDistances, !drawDimensionsEnabled);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Critical Error", $"An unexpected error occurred: {ex.ToString()}");
                return Result.Failed;
            }
        }

        private void ShowResults(
            Document doc,
            UIDocument uidoc,
            List<FamilyInstance> selectedDoors,
            ConcurrentBag<DoorProcessingResult> doorResults,
            Dictionary<int, List<DimensionInfo>> doorDistances,
            bool dimensionsSkippedBySetting)
        {
            var allOrientationLabels = doorDistances.Values
                .SelectMany(dims => dims.Select(d => d.OrientationLabel))
                .Distinct()
                .ToList();

            // Custom ordering: Offset Left, Offset Right, Offset Left Reverse, Offset Right Reverse
            var orderedOrientationLabels = new List<string>();
            
            // First add non-reverse labels
            var nonReverseLabels = allOrientationLabels
                .Where(label => !label.Contains("Reverse"))
                .OrderBy(label => label)
                .ToList();
            orderedOrientationLabels.AddRange(nonReverseLabels);
            
            // Then add reverse labels
            var reverseLabels = allOrientationLabels
                .Where(label => label.Contains("Reverse"))
                .OrderBy(label => label)
                .ToList();
            orderedOrientationLabels.AddRange(reverseLabels);

            // ─── column setup ────────────────────────────────────────────
            List<string> propertyNames = new List<string>
            {
                "Family Name", "Type Name", "Level",
                "FacingFlipped", "HandFlipped",
                "Width (mm)", "Height (mm)", "Adjacent Walls Count"
            };

            // Offset-distance columns now directly follow "Adjacent Walls Count"
            foreach (string orientationLabel in orderedOrientationLabels)
            {
                propertyNames.Add($"{orientationLabel} (mm)");
            }

            // Then the descriptive columns
            propertyNames.AddRange(new[] { "Group", "Room From", "Room To" });

            // "Mark" column
            propertyNames.Add("Mark");

            // (Wall-ID columns removed)

            propertyNames.Add("Door Element Id");
            // ─────────────────────────────────────────────────────────────

            List<Dictionary<string, object>> doorData = new List<Dictionary<string, object>>();

            foreach (var result in doorResults.OrderBy(r => r.Door.Id.IntegerValue))
            {
                if (result.NoHostWall || !string.IsNullOrEmpty(result.Error))
                    continue;

                Dictionary<string, object> doorProperties = new Dictionary<string, object>();
                ElementType doorType = doc.GetElement(result.Door.GetTypeId()) as ElementType;

                doorProperties["Family Name"] = doorType?.FamilyName ?? "";
                doorProperties["Type Name"] = doorType?.Name ?? "";
                doorProperties["Level"] = doc.GetElement(result.Door.LevelId)?.Name ?? "";

                // Flip flags
                doorProperties["FacingFlipped"] = result.Door.FacingFlipped;
                doorProperties["HandFlipped"] = result.Door.HandFlipped;

                double doorWidthParam = doorType?.get_Parameter(BuiltInParameter.DOOR_WIDTH)?.AsDouble() ?? 0.0;
                double doorHeightParam = doorType?.get_Parameter(BuiltInParameter.DOOR_HEIGHT)?.AsDouble() ?? 0.0;
                doorProperties["Width (mm)"] = Math.Round(doorWidthParam * 304.8);
                doorProperties["Height (mm)"] = Math.Round(doorHeightParam * 304.8);

                // Group / room info
                doorProperties["Group"] = result.Door.GroupId != ElementId.InvalidElementId
                    ? doc.GetElement(result.Door.GroupId)?.Name ?? ""
                    : "";
                doorProperties["Room From"] = result.Door.FromRoom?.Name ?? "";
                doorProperties["Room To"] = result.Door.ToRoom?.Name ?? "";

                // Adjacent walls count
                int dimensionedAdjacentWallsCount = 0;
                if (doorDistances.TryGetValue(result.Door.Id.IntegerValue, out List<DimensionInfo> dimsForThisDoorInGrid))
                {
                    dimensionedAdjacentWallsCount = dimsForThisDoorInGrid.Select(d => d.WallId).Distinct().Count();
                }
                doorProperties["Adjacent Walls Count"] = dimensionedAdjacentWallsCount;

                // Initialise distance columns
                foreach (string orientationLabel in orderedOrientationLabels)
                {
                    doorProperties[$"{orientationLabel} (mm)"] = "-";
                }

                // Fill actual distances
                if (doorDistances.TryGetValue(result.Door.Id.IntegerValue, out var distances))
                {
                    foreach (var dim in distances)
                    {
                        double distanceInMm = Math.Round(dim.Value * 304.8);
                        doorProperties[$"{dim.OrientationLabel} (mm)"] = distanceInMm;
                    }
                }

                // Mark
                doorProperties["Mark"] = result.Door.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString() ?? "";

                doorProperties["Door Element Id"] = result.Door.Id.IntegerValue;
                doorData.Add(doorProperties);
            }

            if (doorData.Count > 0)
            {
                List<Dictionary<string, object>> selectedFromGrid = CustomGUIs.DataGrid(doorData, propertyNames, false);
                if (selectedFromGrid?.Any() == true)
                {
                    var finalSelection = selectedDoors
                        .Where(d => selectedFromGrid.Any(s =>
                            s.ContainsKey("Door Element Id") &&
                            (int)s["Door Element Id"] == d.Id.IntegerValue))
                        .Select(d => d.Id)
                        .ToList();
                    if (finalSelection.Any())
                    {
                        uidoc.Selection.SetElementIds(finalSelection);
                    }
                }
            }
            else if (selectedDoors.Any(sd => sd.Host is Wall) && !dimensionsSkippedBySetting)
            {
                TaskDialog.Show("Result", "No doors with adjacent walls suitable for dimensioning were found, or no dimensions were created.");
            }

            // summary tracing unchanged …
            List<string> dimensionResultsSummary = new List<string>();
            if (doorDistances.Any())
            {
                string dimensionStatus = dimensionsSkippedBySetting ? "calculated but not drawn" : "created";
                foreach (var kvp in doorDistances)
                {
                    var orientations = kvp.Value.Select(d => d.OrientationLabel).ToList();
                    var uniqueWallIdsInvolved = kvp.Value.Select(d => d.WallId).Distinct().Count();
                    string summary = $"Door {kvp.Key}: {kvp.Value.Count} dimension(s) {dimensionStatus} to {uniqueWallIdsInvolved} unique wall(s). Orientations: ({string.Join(", ", orientations)})";

                    int requiresBothSidesCount = kvp.Value.Where(d => d.RequiresBothSides).Select(d => d.WallId).Distinct().Count();
                    if (requiresBothSidesCount > 0)
                    {
                        summary += $" - {requiresBothSidesCount} wall(s) spanned both sides (front/back relative to door).";
                    }
                    dimensionResultsSummary.Add(summary);
                }
            }
        }
    }
}
