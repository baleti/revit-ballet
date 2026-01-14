using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

public partial class CustomGUIs
{
    // Store the current UIDocument for edit operations
    private static UIDocument _currentUIDoc = null;

    /// <summary>
    /// Set the current UIDocument for edit operations
    /// Call this from your command before showing the DataGrid
    /// </summary>
    public static void SetCurrentUIDocument(UIDocument uidoc)
    {
        _currentUIDoc = uidoc;
    }

    /// <summary>
    /// Apply pending cell edits to actual Revit elements
    /// </summary>
    public static bool ApplyCellEditsToEntities()
    {

        if (_pendingCellEdits.Count == 0)
        {

            System.Windows.Forms.MessageBox.Show("No edits to apply.", "Apply Edits",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            return false;
        }


        if (_currentUIDoc == null)
        {

            System.Windows.Forms.MessageBox.Show(
                "Cannot apply edits: No active Revit document.\n\n" +
                "Make sure to call CustomGUIs.SetCurrentUIDocument(uidoc) before showing the DataGrid.",
                "Apply Edits Error",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
            return false;
        }


        int successCount = 0;
        int errorCount = 0;
        var errorMessages = new List<string>();

        // Group edits by entry (element) for efficient transaction handling
        var editsByEntry = new Dictionary<long, List<(string columnName, object newValue)>>();

        foreach (var kvp in _pendingCellEdits)
        {
            string editKey = kvp.Key; // format: "InternalID|ColumnName"
            object newValue = kvp.Value;

            // Parse edit key
            var parts = editKey.Split('|');
            if (parts.Length != 2) continue;

            long internalId = long.Parse(parts[0]);
            string columnName = parts[1];

            if (!editsByEntry.ContainsKey(internalId))
                editsByEntry[internalId] = new List<(string, object)>();

            editsByEntry[internalId].Add((columnName, newValue));
        }

        // Group entries by document for cross-document edit support
        var entriesByDocument = new Dictionary<Document, List<(long internalId, Dictionary<string, object> entry, List<(string, object)> edits)>>();


        foreach (var entryGroup in editsByEntry)
        {
            long internalId = entryGroup.Key;
            var edits = entryGroup.Value;


            // Find the entry with this internal ID
            var entry = _modifiedEntries.FirstOrDefault(e => GetInternalId(e) == internalId);
            if (entry == null)
            {
                errorMessages.Add($"Could not find entry with internal ID {internalId}");
                errorCount++;
                continue;
            }

            // Log entry info
            string entryName = entry.ContainsKey("Name") ? entry["Name"]?.ToString() : "(no name)";

            // Determine which document this entry belongs to
            Document entryDoc = null;

            // First, check if entry has __Document field (Session scope commands)
            if (entry.ContainsKey("__Document") && entry["__Document"] is Document doc)
            {
                entryDoc = doc;
            }
            else
            {
                // Fall back to active document
                entryDoc = _currentUIDoc.Document;
            }

            // Group by document
            if (!entriesByDocument.ContainsKey(entryDoc))
            {
                entriesByDocument[entryDoc] = new List<(long, Dictionary<string, object>, List<(string, object)>)>();
            }

            entriesByDocument[entryDoc].Add((internalId, entry, edits));
        }

        foreach (var docGroup in entriesByDocument)
        {
        }

        // Process edits for each document separately with its own transaction

        Document activeDoc = _currentUIDoc.Document;

        foreach (var docGroup in entriesByDocument)
        {
            Document doc = docGroup.Key;
            var entries = docGroup.Value;


            using (Transaction trans = new Transaction(doc, "Apply DataGrid Edits"))
            {
                trans.Start();

                // CRITICAL: Two-phase rename for columns with uniqueness constraints
                // Group edits by column name to detect if we need two-phase rename
                var editsByColumn = new Dictionary<string, List<(long internalId, Dictionary<string, object> entry, object newValue)>>();
                    foreach (var (internalId, entry, edits) in entries)
                    {
                        foreach (var (columnName, newValue) in edits)
                        {
                            if (!editsByColumn.ContainsKey(columnName))
                                editsByColumn[columnName] = new List<(long, Dictionary<string, object>, object)>();
                            editsByColumn[columnName].Add((internalId, entry, newValue));
                        }
                    }

                    // Check which columns require two-phase rename
                    var uniqueNameColumns = new HashSet<string>();
                    foreach (var columnName in editsByColumn.Keys)
                    {
                        var handler = ColumnHandlerRegistry.GetHandler(columnName);
                        if (handler != null && handler.RequiresUniqueName && editsByColumn[columnName].Count > 1)
                        {
                            uniqueNameColumns.Add(columnName);
                        }
                    }

                    // DUPLICATE DETECTION: Check for duplicate target names before starting two-phase rename
                    if (uniqueNameColumns.Any())
                    {
                        var duplicateInfo = DetectDuplicateNames(editsByColumn, uniqueNameColumns, doc);

                        if (duplicateInfo.HasDuplicates)
                        {

                            // Show dialog to user
                            var dialogResult = ShowDuplicateNameDialog(duplicateInfo);

                            if (dialogResult == System.Windows.Forms.DialogResult.Cancel)
                            {
                                trans.RollBack();

                                // Clear pending edits
                                _pendingCellEdits.Clear();
                                return false;
                            }
                            else if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                            {
                                AppendUniqueSuffixes(editsByColumn, duplicateInfo);
                            }
                            // DialogResult.No means "proceed with duplicates" - no action needed
                        }
                        else
                        {
                        }
                    }

                    // Phase 1: Rename all unique-name columns to temporary names
                    if (uniqueNameColumns.Any())
                    {
                        // Random number generator for numeric temp values
                        var random = new System.Random();

                        foreach (var columnName in uniqueNameColumns)
                        {
                            var columnEdits = editsByColumn[columnName];

                            // Determine if this column needs numeric temp values
                            bool useNumericTemp = columnName.Equals("Detail Number", StringComparison.OrdinalIgnoreCase);

                            foreach (var (internalId, entry, newValue) in columnEdits)
                            {
                                Element elem = GetElementFromEntry(doc, entry);
                                if (elem == null)
                                {
                                    continue;
                                }

                                // Generate temporary unique name/number
                                string tempName;
                                if (useNumericTemp)
                                {
                                    // Use random large number for numeric fields (Detail Number, etc.)
                                    tempName = random.Next(900000, 999999).ToString();
                                }
                                else
                                {
                                    // Use UUID-based string for text fields (View Name, Sheet Name, etc.)
                                    tempName = $"{elem.Id.AsLong()}-temp-{Guid.NewGuid().ToString().Substring(0, 8)}";
                                }

                                var handler = ColumnHandlerRegistry.GetHandler(columnName);
                                if (handler != null && handler.Setter != null)
                                {
                                    try
                                    {
                                        handler.Setter(elem, doc, tempName);
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                        }

                        // Phase 2: Rename from temporary names to final target names
                        foreach (var columnName in uniqueNameColumns)
                        {
                            var columnEdits = editsByColumn[columnName];

                            foreach (var (internalId, entry, newValue) in columnEdits)
                            {
                                Element elem = GetElementFromEntry(doc, entry);
                                if (elem == null)
                                {
                                    errorCount++;
                                    continue;
                                }

                                string finalName = newValue?.ToString() ?? "";

                                var handler = ColumnHandlerRegistry.GetHandler(columnName);
                                if (handler != null && handler.Setter != null)
                                {
                                    try
                                    {
                                        bool success = handler.Setter(elem, doc, finalName);
                                        if (success)
                                            successCount++;
                                        else
                                        {
                                            errorMessages.Add($"{GetEntryDisplayName(entry)}.{columnName}: Failed to apply");
                                            errorCount++;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        errorMessages.Add($"{GetEntryDisplayName(entry)}.{columnName}: {ex.Message}");
                                        errorCount++;
                                    }
                                }
                            }
                        }
                    }

                    // Now process remaining edits (non-unique-name columns)

                    foreach (var (internalId, entry, edits) in entries)
                    {
                        string entryName = entry.ContainsKey("Name") ? entry["Name"]?.ToString() : "(no name)";

                        // Check if this is a workset entry (from SelectByWorksetsIn* commands)
                        bool isWorksetEntry = entry.ContainsKey("WorksetId") && entry.ContainsKey("Workset");

                        if (isWorksetEntry)
                        {
                            // Apply workset-specific edits
                            foreach (var (columnName, newValue) in edits)
                            {
                                // Skip if already processed in two-phase rename
                                if (uniqueNameColumns.Contains(columnName))
                                {
                                    continue;
                                }

                                string worksetName = entry.ContainsKey("Workset") ? entry["Workset"]?.ToString() : "(unknown)";

                                try
                                {
                                    if (ApplyWorksetEdit(doc, entry, columnName, newValue))
                                    {
                                        successCount++;
                                    }
                                    else
                                    {
                                        errorMessages.Add($"{GetEntryDisplayName(entry)}.{columnName}: Failed to apply");
                                        errorCount++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errorMessages.Add($"{GetEntryDisplayName(entry)}.{columnName}: {ex.Message}");
                                    errorCount++;
                                }
                            }
                        }
                        else
                        {
                            // Get the Revit element
                            Element elem = GetElementFromEntry(doc, entry);
                            if (elem == null)
                            {
                                errorMessages.Add($"Could not find Revit element for entry: {GetEntryDisplayName(entry)}");
                                errorCount++;
                                continue;
                            }

                            // BATCH CROP REGION EDITS: Apply all crop region columns together to avoid invalid intermediate states
                            var cropRegionColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                            {
                                "Crop Region Top", "Crop Region Bottom", "Crop Region Left", "Crop Region Right"
                            };
                            var cropEdits = edits.Where(e => cropRegionColumns.Contains(e.Item1)).ToList();
                            var processedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                            if (cropEdits.Count > 1) // Only batch if 2+ crop edits (single edits work fine individually)
                            {

                                // Extract values
                                double? newTop = null, newBottom = null, newLeft = null, newRight = null;
                                foreach (var cropEdit in cropEdits)
                                {
                                    string columnName = cropEdit.Item1;
                                    object newValue = cropEdit.Item2;

                                    processedColumns.Add(columnName);

                                    if (columnName.Equals("Crop Region Top", StringComparison.OrdinalIgnoreCase) &&
                                        double.TryParse(newValue?.ToString(), out double v1))
                                        newTop = v1;
                                    else if (columnName.Equals("Crop Region Bottom", StringComparison.OrdinalIgnoreCase) &&
                                        double.TryParse(newValue?.ToString(), out double v2))
                                        newBottom = v2;
                                    else if (columnName.Equals("Crop Region Left", StringComparison.OrdinalIgnoreCase) &&
                                        double.TryParse(newValue?.ToString(), out double v3))
                                        newLeft = v3;
                                    else if (columnName.Equals("Crop Region Right", StringComparison.OrdinalIgnoreCase) &&
                                        double.TryParse(newValue?.ToString(), out double v4))
                                        newRight = v4;
                                }

                                try
                                {
                                    if (ColumnHandlerRegistry.ApplyCropRegionEdits(elem, doc, newTop, newBottom, newLeft, newRight))
                                    {
                                        successCount += cropEdits.Count;
                                    }
                                    else
                                    {
                                        foreach (var cropEdit in cropEdits)
                                        {
                                            errorMessages.Add($"{GetEntryDisplayName(entry)}.{cropEdit.Item1}: Failed to apply");
                                        }
                                        errorCount += cropEdits.Count;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    foreach (var cropEdit in cropEdits)
                                    {
                                        errorMessages.Add($"{GetEntryDisplayName(entry)}.{cropEdit.Item1}: {ex.Message}");
                                    }
                                    errorCount += cropEdits.Count;
                                }
                            }

                            // Apply each edit for this element
                            foreach (var (columnName, newValue) in edits)
                            {
                                // Skip if already processed in two-phase rename
                                if (uniqueNameColumns.Contains(columnName))
                                {
                                    continue;
                                }

                                // Skip if already processed in crop region batch
                                if (processedColumns.Contains(columnName))
                                {
                                    continue;
                                }


                                try
                                {
                                    if (ApplyPropertyEdit(elem, columnName, newValue, entry))
                                    {
                                        successCount++;
                                    }
                                    else
                                    {
                                        errorMessages.Add($"{GetEntryDisplayName(entry)}.{columnName}: Failed to apply");
                                        errorCount++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errorMessages.Add($"{GetEntryDisplayName(entry)}.{columnName}: {ex.Message}");
                                    errorCount++;
                                }
                            }
                        }
                    }


                    // RESILIENT COMMIT: Try to commit even if some individual edits failed
                    // The transaction contains all successful edits from try-catch blocks above
                    try
                    {
                        trans.Commit();
                        _editsWereApplied = true;
                    }
                    catch (Exception commitEx)
                    {
                        // CRITICAL: Only catch commit-specific failures here
                        // Individual edit failures are already caught and tracked above

                        try
                        {
                            trans.RollBack();
                        }
                        catch
                        {
                        }

                        // Write diagnostics immediately

                        // Add commit failure to error messages for summary
                        errorMessages.Add($"CRITICAL: Transaction commit failed - {commitEx.Message}. All edits rolled back.");
                        // Count all pending edits as failed
                        errorCount += editsByEntry.Count;
                    }
            }
            // using block has exited - transaction is now properly disposed
        }

        // Write final diagnostics


        // CRITICAL: Clear pending edits to prevent double-application
        // Some commands (like OpenSheetsInSession) manually check HasPendingEdits() after
        // DataGrid auto-applies, which would cause edits to be applied twice if not cleared
        _pendingCellEdits.Clear();

        // Show results only if there were errors
        if (errorCount > 0)
        {
            string resultMessage = $"Applied {successCount} edit(s) successfully.";
            resultMessage += $"\n\n{errorCount} edit(s) failed:";
            resultMessage += "\n" + string.Join("\n", errorMessages.Take(10));
            if (errorMessages.Count > 10)
                resultMessage += $"\n... and {errorMessages.Count - 10} more";

            System.Windows.Forms.MessageBox.Show(
                resultMessage,
                "Apply Edits - Some Failed",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
        }

        return true;
    }

    /// <summary>
    /// Get a Revit Element from a DataGrid entry
    /// Supports various entry formats: ElementId, ElementIdObject, Id (long)
    /// </summary>
    private static Element GetElementFromEntry(Document doc, Dictionary<string, object> entry)
    {
        // Try ElementIdObject first (most direct)
        if (entry.ContainsKey("ElementIdObject") && entry["ElementIdObject"] is ElementId elemId)
        {
            return doc.GetElement(elemId);
        }

        // Try ElementId (if stored directly)
        if (entry.ContainsKey("ElementId") && entry["ElementId"] is ElementId elemId2)
        {
            return doc.GetElement(elemId2);
        }

        // Try Id as long
        if (entry.ContainsKey("Id"))
        {
            try
            {
                long id = Convert.ToInt64(entry["Id"]);
                return doc.GetElement(id.ToElementId());
            }
            catch { }
        }

        // Try parsing various ID fields
        foreach (var key in new[] { "ElementId", "Id", "ElemId", "Element Id" })
        {
            if (entry.ContainsKey(key))
            {
                try
                {
                    var value = entry[key];
                    if (value is int intId)
                        return doc.GetElement(intId.ToElementId());
                    else if (value is long longId)
                        return doc.GetElement(longId.ToElementId());
                    else if (value is string strId && long.TryParse(strId, out long parsedId))
                        return doc.GetElement(parsedId.ToElementId());
                }
                catch { }
            }
        }

        return null;
    }

    /// <summary>
    /// Get display name for an entry (for error messages)
    /// </summary>
    private static string GetEntryDisplayName(Dictionary<string, object> entry)
    {
        if (entry.ContainsKey("Name") && entry["Name"] != null)
            return entry["Name"].ToString();
        if (entry.ContainsKey("DisplayName") && entry["DisplayName"] != null)
            return entry["DisplayName"].ToString();
        if (entry.ContainsKey("Id"))
            return $"ID {entry["Id"]}";
        return "Unknown";
    }

    /// <summary>
    /// Apply a property edit to a Revit element
    /// </summary>
    private static bool ApplyPropertyEdit(Element elem, string columnName, object newValue, Dictionary<string, object> entry)
    {
        // AUTOMATIC SYSTEM: Try handler registry first, with fallback to dynamic parameter detection
        ColumnHandlerRegistry.EnsureInitialized();
        var handler = ColumnHandlerRegistry.GetHandlerWithFallback(columnName, elem, _currentUIDoc?.Document);

        if (handler != null && handler.IsEditable && handler.Setter != null)
        {
            try
            {
                // Use handler to apply edit (includes both explicit and dynamic parameter handlers)
                bool success = handler.ApplyEdit(elem, _currentUIDoc.Document, newValue);
                return success;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Handler for '{columnName}' failed: {ex.Message}");
                return false;
            }
        }

        // FALLBACK: Legacy switch-based system for backward compatibility
        string lowerName = columnName.ToLowerInvariant();
        string strValue = newValue?.ToString() ?? "";

        // Built-in element properties
        switch (lowerName)
        {
            case "name":
            case "displayname":
            case "type":
            case "typename":
            case "type name":
                // SPECIAL CASE: ViewSheets use SHEET_NAME built-in parameter
                if (elem is ViewSheet viewSheet)
                {
                    try
                    {
                        Parameter sheetNameParam = viewSheet.get_Parameter(BuiltInParameter.SHEET_NAME);
                        if (sheetNameParam != null)
                        {
                            sheetNameParam.Set(strValue);
                            return true;
                        }
                    }
                    catch { }
                }

                // SPECIAL CASE: OST_Viewers (Callouts, Sections, Elevations)
                // These elements appear in category "Views" but may not be castable to View class
                // Some are wrapper/reference elements that point to actual views via ID_PARAM
                if (elem.Category?.Id.AsLong() == (int)BuiltInCategory.OST_Viewers)
                {
                    // First, try VIEW_NAME parameter (works for some view types like sections)
                    try
                    {
                        Parameter viewNameParam = elem.get_Parameter(BuiltInParameter.VIEW_NAME);
                        if (viewNameParam != null && !viewNameParam.IsReadOnly)
                        {
                            viewNameParam.Set(strValue);
                            return true;
                        }
                    }
                    catch { }

                    // If VIEW_NAME didn't work, check if this is a reference element
                    // that points to an actual view via ID_PARAM
                    try
                    {
                        Parameter idParam = elem.get_Parameter(BuiltInParameter.ID_PARAM);
                        if (idParam != null && idParam.HasValue)
                        {
                            ElementId referencedId = idParam.AsElementId();
                            if (referencedId != null && referencedId != ElementId.InvalidElementId)
                            {
                                Element referencedElem = elem.Document.GetElement(referencedId);
                                if (referencedElem is View referencedView)
                                {
                                    // Found the actual view - rename it
                                    referencedView.Name = strValue;
                                    return true;
                                }
                            }
                        }
                    }
                    catch { }
                }

                // SPECIAL CASE: Regular Views use the Name property directly
                if (elem is View regularView && !(elem is ViewSheet))
                {
                    try
                    {
                        regularView.Name = strValue;
                        return true;
                    }
                    catch { }
                }

                // If element IS a group itself, rename its group type
                if (elem is Group selectedGroup)
                {
                    try
                    {
                        selectedGroup.GroupType.Name = strValue;
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }

                // IMPORTANT: Check if element has a writable "Name" instance parameter first
                // Elements like Grids, Levels, etc. have Name instance parameters that should take precedence
                Parameter nameParam = elem.LookupParameter("Name");
                if (nameParam != null && !nameParam.IsReadOnly && nameParam.StorageType == StorageType.String)
                {
                    try
                    {
                        nameParam.Set(strValue);
                        return true;
                    }
                    catch
                    {
                        // If setting the parameter fails, fall through to type renaming
                    }
                }

                // For all other elements (including group members), rename the element's TYPE
                // Note: Even if an element is in a group, we can still rename its family type
                // The element's Name property refers to its type name, not the group name
                ElementId typeId = elem.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    Element elemType = elem.Document.GetElement(typeId);
                    if (elemType != null)
                    {
                        try
                        {
                            elemType.Name = strValue;
                            return true;
                        }
                        catch
                        {
                            return false;
                        }
                    }
                }

                // Fallback: try to rename the element directly (this may not work for most elements)
                try
                {
                    elem.Name = strValue;
                    return true;
                }
                catch
                {
                    return false;
                }

            case "comments":
                return SetBuiltInParameter(elem, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, strValue);

            case "mark":
                return SetBuiltInParameter(elem, BuiltInParameter.ALL_MODEL_MARK, strValue);

            case "description":
                return SetBuiltInParameter(elem, BuiltInParameter.ALL_MODEL_DESCRIPTION, strValue);

            // Workset
            case "workset":
            case "worksetname":
                return SetWorkset(elem, strValue);

            // Level
            case "level":
            case "levelname":
                return SetLevel(elem, strValue);

            // Constraint parameters (wall-specific)
            case "baseconstraint":
            case "base constraint":
                return SetLevelParameter(elem, BuiltInParameter.WALL_BASE_CONSTRAINT, strValue);

            case "topconstraint":
            case "top constraint":
                return SetLevelParameter(elem, BuiltInParameter.WALL_HEIGHT_TYPE, strValue);

            case "baseoffset":
            case "base offset":
                return SetBuiltInParameter(elem, BuiltInParameter.WALL_BASE_OFFSET, strValue);

            case "topoffset":
            case "top offset":
                return SetBuiltInParameter(elem, BuiltInParameter.WALL_TOP_OFFSET, strValue);

            case "unconnectedheight":
            case "unconnected height":
                return SetBuiltInParameter(elem, BuiltInParameter.WALL_USER_HEIGHT_PARAM, strValue);

            // Phase
            case "phasecreated":
                return SetPhase(elem, BuiltInParameter.PHASE_CREATED, strValue);

            case "phasedemolished":
                return SetPhase(elem, BuiltInParameter.PHASE_DEMOLISHED, strValue);

            // Sheet properties
            case "sheetnumber":
                if (elem is ViewSheet sheet)
                {
                    try
                    {
                        sheet.SheetNumber = strValue;
                        return true;
                    }
                    catch { return false; }
                }
                return false;

            case "sheetname":
                return SetBuiltInParameter(elem, BuiltInParameter.SHEET_NAME, strValue);

            case "drawnby":
                return SetBuiltInParameter(elem, BuiltInParameter.SHEET_DRAWN_BY, strValue);

            case "checkedby":
                return SetBuiltInParameter(elem, BuiltInParameter.SHEET_CHECKED_BY, strValue);

            case "approvedby":
                return SetBuiltInParameter(elem, BuiltInParameter.SHEET_APPROVED_BY, strValue);

            case "issuedby":
                return SetParameterValue(elem, "Issued By", strValue);

            case "issuedto":
                return SetParameterValue(elem, "Issued To", strValue);

            case "sheetissuedate":
            case "issuedate":
                return SetBuiltInParameter(elem, BuiltInParameter.SHEET_ISSUE_DATE, strValue);

            case "currentrevision":
            case "revisionnumber":
                return SetBuiltInParameter(elem, BuiltInParameter.SHEET_CURRENT_REVISION, strValue);

            case "currentrevisiondate":
            case "revisiondate":
                return SetBuiltInParameter(elem, BuiltInParameter.SHEET_CURRENT_REVISION_DATE, strValue);

            case "currentrevisiondescription":
            case "revisiondescription":
                return SetBuiltInParameter(elem, BuiltInParameter.SHEET_CURRENT_REVISION_DESCRIPTION, strValue);

            // View properties
            case "viewname":
                if (elem is View view)
                {
                    try
                    {
                        view.Name = strValue;
                        return true;
                    }
                    catch { return false; }
                }
                return false;

            case "scale":
            case "viewscale":
                if (elem is View v && v.CanBePrinted) // Only views that can be printed have scale
                {
                    if (int.TryParse(strValue, out int scaleValue))
                    {
                        try
                        {
                            v.Scale = scaleValue;
                            return true;
                        }
                        catch { return false; }
                    }
                }
                return false;

            case "detaillevel":
                if (elem is View vDetail)
                {
                    if (Enum.TryParse(strValue, true, out ViewDetailLevel detailLevel))
                    {
                        try
                        {
                            vDetail.DetailLevel = detailLevel;
                            return true;
                        }
                        catch { return false; }
                    }
                }
                return false;

            case "viewtemplate":
            case "template":
                if (elem is View vTemplate)
                {
                    // Find view template by name
                    var templates = new FilteredElementCollector(elem.Document)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => v.IsTemplate)
                        .ToList();

                    View targetTemplate = templates.FirstOrDefault(t =>
                        string.Equals(t.Name, strValue, StringComparison.OrdinalIgnoreCase));

                    if (targetTemplate != null)
                    {
                        try
                        {
                            vTemplate.ViewTemplateId = targetTemplate.Id;
                            return true;
                        }
                        catch { return false; }
                    }
                    else if (string.IsNullOrEmpty(strValue) || strValue.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove template
                        try
                        {
                            vTemplate.ViewTemplateId = ElementId.InvalidElementId;
                            return true;
                        }
                        catch { return false; }
                    }
                }
                return false;

            case "discipline":
                if (elem is View vDisc)
                {
                    if (Enum.TryParse(strValue, true, out ViewDiscipline discipline))
                    {
                        try
                        {
                            vDisc.Discipline = discipline;
                            return true;
                        }
                        catch { return false; }
                    }
                }
                return false;

            case "title":
            case "titleonsheet":
                if (elem is View vTitle)
                {
                    try
                    {
                        vTitle.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).Set(strValue);
                        return true;
                    }
                    catch { return false; }
                }
                return false;

            case "ownerview":
                // When editing OwnerView, we rename the view that owns this element
                // This is especially useful for Viewports - renaming the view placed on a sheet
                if (elem.OwnerViewId != null && elem.OwnerViewId != ElementId.InvalidElementId)
                {
                    Element ownerViewElem = elem.Document.GetElement(elem.OwnerViewId);
                    if (ownerViewElem is View ownerView)
                    {
                        try
                        {
                            ownerView.Name = strValue;
                            return true;
                        }
                        catch { return false; }
                    }
                }
                return false;

            case "scope box":
            case "scopebox":
                if (elem is View vScopeBox)
                {
                    // Log diagnostics to runtime directory


                    Parameter scopeBoxParam = vScopeBox.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP);
                    if (scopeBoxParam == null)
                    {
                        return false;
                    }


                    if (scopeBoxParam.HasValue)
                    {
                        ElementId currentScopeId = scopeBoxParam.AsElementId();
                        if (currentScopeId != null && currentScopeId != ElementId.InvalidElementId)
                        {
                            Element currentScope = elem.Document.GetElement(currentScopeId);
                        }
                        else
                        {
                        }
                    }

                    if (scopeBoxParam.IsReadOnly)
                    {
                        return false;
                    }

                    // Handle empty/none case - remove scope box
                    if (string.IsNullOrEmpty(strValue) || strValue.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            scopeBoxParam.Set(ElementId.InvalidElementId);
                            return true;
                        }
                        catch
                        {
                            return false;
                        }
                    }

                    // Find scope box by name
                    var scopeBoxes = new FilteredElementCollector(elem.Document)
                        .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                        .WhereElementIsNotElementType()
                        .Cast<Element>()
                        .ToList();

                    foreach (var sb in scopeBoxes)
                    {
                    }

                    Element targetScopeBox = scopeBoxes.FirstOrDefault(sb =>
                        string.Equals(sb.Name, strValue, StringComparison.OrdinalIgnoreCase));

                    if (targetScopeBox != null)
                    {
                        try
                        {
                            scopeBoxParam.Set(targetScopeBox.Id);
                            return true;
                        }
                        catch
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                return false;

            // Family properties
            case "familyname":
            case "family name":
            case "family":
                ElementId famTypeId = elem.GetTypeId();
                if (famTypeId != null && famTypeId != ElementId.InvalidElementId)
                {
                    Element famElemType = elem.Document.GetElement(famTypeId);
                    if (famElemType is FamilySymbol familySymbol)
                    {
                        try
                        {
                            Family family = familySymbol.Family;
                            family.Name = strValue;
                            return true;
                        }
                        catch { return false; }
                    }
                }
                return false;

            // Design Option
            case "designoption":
                return SetDesignOption(elem, strValue);

            // Subcategory
            case "subcategory":
                return SetSubcategory(elem, strValue);

            // Centroid coordinates
            case "x centroid":
            case "xcentroid":
                return SetCentroidX(elem, strValue);

            case "y centroid":
            case "ycentroid":
                return SetCentroidY(elem, strValue);

            case "z centroid":
            case "zcentroid":
                return SetCentroidZ(elem, strValue);

            // Room/Space properties (handled via parameters)
            case "number":
                return SetParameterValue(elem, "Number", strValue);

            // IFC parameters
            case "exporttoifc":
            case "export to ifc":
            case "ifc export":
                return SetYesNoByTypeParameter(elem, "Export to IFC", strValue);

            default:
                // Handle parameter columns (param_*, sharedparam_*, typeparam_*)
                if (lowerName.StartsWith("param_"))
                {
                    string paramName = columnName.Substring(6); // Remove "param_" prefix
                    return SetParameterValue(elem, paramName, strValue);
                }
                else if (lowerName.StartsWith("sharedparam_"))
                {
                    string paramName = columnName.Substring(12); // Remove "sharedparam_" prefix
                    return SetParameterValue(elem, paramName, strValue);
                }
                else if (lowerName.StartsWith("typeparam_"))
                {
                    string paramName = columnName.Substring(10); // Remove "typeparam_" prefix
                    // For type parameters, get the element type and set its parameter
                    ElementType elemType = elem.Document.GetElement(elem.GetTypeId()) as ElementType;
                    if (elemType != null)
                        return SetParameterValue(elemType, paramName, strValue);
                    return false;
                }
                else
                {
                    // Try as a regular parameter lookup
                    return SetParameterValue(elem, columnName, strValue);
                }
        }
    }

    /// <summary>
    /// Set a built-in parameter value
    /// </summary>
    private static bool SetBuiltInParameter(Element elem, BuiltInParameter builtInParam, string value)
    {
        Parameter param = elem.get_Parameter(builtInParam);
        if (param == null || param.IsReadOnly)
            return false;

        return SetParameterValueInternal(param, value);
    }

    /// <summary>
    /// Set a parameter by name
    /// </summary>
    private static bool SetParameterValue(Element elem, string paramName, string value)
    {
        // Special handling for built-in properties that aren't regular parameters
        string lowerParamName = paramName.ToLowerInvariant();

        // Workset
        if (lowerParamName == "workset" || lowerParamName == "worksetname")
            return SetWorkset(elem, value);

        // Level
        if (lowerParamName == "level" || lowerParamName == "levelname")
            return SetLevel(elem, value);

        // Try exact match first
        Parameter param = elem.LookupParameter(paramName);

        // If not found, try case-insensitive match
        if (param == null)
        {
            foreach (Parameter p in elem.Parameters)
            {
                if (string.Equals(p.Definition.Name, paramName, StringComparison.OrdinalIgnoreCase))
                {
                    param = p;
                    break;
                }
            }
        }

        // Try with underscore/space conversion
        if (param == null)
        {
            string paramNameWithSpaces = paramName.Replace("_", " ");
            param = elem.LookupParameter(paramNameWithSpaces);
        }

        if (param == null)
        {
            string paramNameWithUnderscores = paramName.Replace(" ", "_");
            param = elem.LookupParameter(paramNameWithUnderscores);
        }

        if (param == null || param.IsReadOnly)
            return false;

        return SetParameterValueInternal(param, value);
    }

    /// <summary>
    /// Set parameter value based on storage type
    /// </summary>
    private static bool SetParameterValueInternal(Parameter param, string value)
    {
        try
        {
            switch (param.StorageType)
            {
                case StorageType.String:
                    param.Set(value);
                    return true;

                case StorageType.Integer:
                    if (int.TryParse(value, out int intValue))
                    {
                        param.Set(intValue);
                        return true;
                    }
                    // Also try as ElementId for integer parameters
                    if (int.TryParse(value, out int elemIdInt))
                    {
                        param.Set(elemIdInt.ToElementId());
                        return true;
                    }
                    return false;

                case StorageType.Double:
                    // CRITICAL: Use SetValueString FIRST to handle unit conversion properly
                    // SetValueString respects the project's display units (metric vs imperial)
                    // and converts the user's input to internal units automatically.
                    // Direct Set() assumes the value is already in internal units (feet),
                    // which causes incorrect conversions in metric projects.
                    try
                    {
                        param.SetValueString(value);
                        return true;
                    }
                    catch
                    {
                        // SetValueString failed - try direct parsing as fallback
                        // This handles cases where the parameter doesn't support SetValueString
                        // or the value is already in internal units
                        if (double.TryParse(value, out double doubleValue))
                        {
                            try
                            {
                                param.Set(doubleValue);
                                return true;
                            }
                            catch
                            {
                                return false;
                            }
                        }
                        return false;
                    }

                case StorageType.ElementId:
                    // Try parsing as integer first
                    if (int.TryParse(value, out int idValue))
                    {
                        param.Set(idValue.ToElementId());
                        return true;
                    }
                    // Try parsing as long
                    if (long.TryParse(value, out long longIdValue))
                    {
                        param.Set(longIdValue.ToElementId());
                        return true;
                    }
                    // Try finding element by name (for named references like levels, worksets, etc.)
                    // This is a fallback for user-friendly input
                    return false;

                default:
                    return false;
            }
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Set element's workset by name
    /// </summary>
    private static bool SetWorkset(Element elem, string worksetName)
    {
        Document doc = elem.Document;

        if (!doc.IsWorkshared)
            return false;

        Parameter worksetParam = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
        if (worksetParam == null || worksetParam.IsReadOnly)
            return false;

        // Trim whitespace from input
        worksetName = worksetName?.Trim();
        if (string.IsNullOrEmpty(worksetName))
            return false;

        // Find workset by name
        var worksets = new FilteredWorksetCollector(doc)
            .OfKind(WorksetKind.UserWorkset)
            .ToList();

        Workset targetWorkset = worksets.FirstOrDefault(w =>
            string.Equals(w.Name, worksetName, StringComparison.OrdinalIgnoreCase));

        if (targetWorkset == null)
            return false;

        try
        {
            worksetParam.Set(targetWorkset.Id.IntegerValue);
            return true;
        }
        catch (Exception ex)
        {
            // Log the actual exception for debugging
            System.Diagnostics.Debug.WriteLine($"Failed to set workset: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Set element's level by name
    /// </summary>
    private static bool SetLevel(Element elem, string levelName)
    {
        Document doc = elem.Document;

        Parameter levelParam = elem.get_Parameter(BuiltInParameter.LEVEL_PARAM);
        if (levelParam == null || levelParam.IsReadOnly)
            return false;

        // Find level by name
        var levels = new FilteredElementCollector(doc)
            .OfClass(typeof(Level))
            .Cast<Level>()
            .ToList();

        Level targetLevel = levels.FirstOrDefault(l =>
            string.Equals(l.Name, levelName, StringComparison.OrdinalIgnoreCase));

        if (targetLevel == null)
            return false;

        try
        {
            levelParam.Set(targetLevel.Id);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Set a Yes/No/By Type parameter (like Export to IFC)
    /// </summary>
    private static bool SetYesNoByTypeParameter(Element elem, string paramName, string value)
    {
        Parameter param = elem.LookupParameter(paramName);
        if (param == null || param.IsReadOnly)
            return false;

        // Parse the value
        string lowerValue = value.Trim().ToLowerInvariant();
        int intValue;

        switch (lowerValue)
        {
            case "yes":
            case "true":
            case "1":
                intValue = 1;
                break;

            case "no":
            case "false":
            case "0":
                intValue = 0;
                break;

            case "by type":
            case "bytype":
            case "-1":
            case "default":
                intValue = -1;
                break;

            default:
                // Try parsing as integer directly
                if (!int.TryParse(value, out intValue))
                    return false;
                break;
        }

        try
        {
            param.Set(intValue);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Set a level parameter (like Base Constraint, Top Constraint) by level name
    /// </summary>
    private static bool SetLevelParameter(Element elem, BuiltInParameter builtInParam, string levelName)
    {
        Document doc = elem.Document;

        Parameter param = elem.get_Parameter(builtInParam);
        if (param == null || param.IsReadOnly)
            return false;

        // Find level by name
        var levels = new FilteredElementCollector(doc)
            .OfClass(typeof(Level))
            .Cast<Level>()
            .ToList();

        Level targetLevel = levels.FirstOrDefault(l =>
            string.Equals(l.Name, levelName, StringComparison.OrdinalIgnoreCase));

        if (targetLevel == null)
            return false;

        try
        {
            param.Set(targetLevel.Id);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Set element's phase by name
    /// </summary>
    private static bool SetPhase(Element elem, BuiltInParameter phaseParam, string phaseName)
    {
        Document doc = elem.Document;

        Parameter param = elem.get_Parameter(phaseParam);
        if (param == null || param.IsReadOnly)
            return false;

        // Find phase by name
        PhaseArray phases = doc.Phases;
        Phase targetPhase = null;

        foreach (Phase phase in phases)
        {
            if (string.Equals(phase.Name, phaseName, StringComparison.OrdinalIgnoreCase))
            {
                targetPhase = phase;
                break;
            }
        }

        if (targetPhase == null)
            return false;

        try
        {
            param.Set(targetPhase.Id);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Set element's design option by name
    /// </summary>
    private static bool SetDesignOption(Element elem, string designOptionName)
    {
        Document doc = elem.Document;

        Parameter designOptionParam = elem.get_Parameter(BuiltInParameter.DESIGN_OPTION_ID);
        if (designOptionParam == null || designOptionParam.IsReadOnly)
            return false;

        // Handle "Main Model" or "None" to remove from design option
        if (string.IsNullOrEmpty(designOptionName) ||
            designOptionName.Equals("Main Model", StringComparison.OrdinalIgnoreCase) ||
            designOptionName.Equals("None", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                designOptionParam.Set(ElementId.InvalidElementId);
                return true;
            }
            catch { return false; }
        }

        // Find design option by name
        var designOptions = new FilteredElementCollector(doc)
            .OfClass(typeof(DesignOption))
            .Cast<DesignOption>()
            .ToList();

        DesignOption targetOption = designOptions.FirstOrDefault(opt =>
            string.Equals(opt.Name, designOptionName, StringComparison.OrdinalIgnoreCase));

        if (targetOption == null)
            return false;

        try
        {
            designOptionParam.Set(targetOption.Id);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Set element's subcategory by name
    /// </summary>
    private static bool SetSubcategory(Element elem, string subcategoryName)
    {
        Document doc = elem.Document;

        // Get the element's category
        Category category = elem.Category;
        if (category == null)
            return false;

        // Get subcategory parameter
        Parameter subcatParam = elem.get_Parameter(BuiltInParameter.FAMILY_ELEM_SUBCATEGORY);
        if (subcatParam == null || subcatParam.IsReadOnly)
            return false;

        // Handle "None" or empty to use parent category
        if (string.IsNullOrEmpty(subcategoryName) ||
            subcategoryName.Equals("None", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                subcatParam.Set(category.Id);
                return true;
            }
            catch { return false; }
        }

        // Find subcategory by name within this category
        Category targetSubcategory = null;
        foreach (Category subcat in category.SubCategories)
        {
            if (string.Equals(subcat.Name, subcategoryName, StringComparison.OrdinalIgnoreCase))
            {
                targetSubcategory = subcat;
                break;
            }
        }

        if (targetSubcategory == null)
            return false;

        try
        {
            subcatParam.Set(targetSubcategory.Id);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Set element's X centroid coordinate
    /// </summary>
    private static bool SetCentroidX(Element elem, string value)
    {
        if (!double.TryParse(value, out double newX))
            return false;

        // Convert from display units to internal units (feet)
        double newXInternal = ConvertToInternalUnits(elem.Document, newX);
        return MoveCentroid(elem, newXInternal, null, null);
    }

    /// <summary>
    /// Set element's Y centroid coordinate
    /// </summary>
    private static bool SetCentroidY(Element elem, string value)
    {
        if (!double.TryParse(value, out double newY))
            return false;

        // Convert from display units to internal units (feet)
        double newYInternal = ConvertToInternalUnits(elem.Document, newY);
        return MoveCentroid(elem, null, newYInternal, null);
    }

    /// <summary>
    /// Set element's Z centroid coordinate
    /// </summary>
    private static bool SetCentroidZ(Element elem, string value)
    {
        if (!double.TryParse(value, out double newZ))
            return false;

        // Convert from display units to internal units (feet)
        double newZInternal = ConvertToInternalUnits(elem.Document, newZ);
        return MoveCentroid(elem, null, null, newZInternal);
    }

    /// <summary>
    /// Convert from display units to internal units (feet)
    /// </summary>
    private static double ConvertToInternalUnits(Document doc, double displayValue)
    {
#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
        Units projectUnits = doc.GetUnits();
        FormatOptions lengthOpts = projectUnits.GetFormatOptions(SpecTypeId.Length);
        ForgeTypeId unitTypeId = lengthOpts.GetUnitTypeId();
        return UnitUtils.ConvertToInternalUnits(displayValue, unitTypeId);
#else
        // Revit 2017-2020: Use DisplayUnitType
        Units projectUnits = doc.GetUnits();
        FormatOptions lengthOpts = projectUnits.GetFormatOptions(UnitType.UT_Length);
        DisplayUnitType unitType = lengthOpts.DisplayUnits;
        return UnitUtils.ConvertToInternalUnits(displayValue, unitType);
#endif
    }

    /// <summary>
    /// Move element by updating its centroid coordinates
    /// </summary>
    private static bool MoveCentroid(Element elem, double? newX, double? newY, double? newZ)
    {
        try
        {
            // Special case: Viewports on sheets
            if (elem is Viewport viewport)
            {
                // Get current position
                XYZ currentCenter = viewport.GetBoxCenter();

                // Viewports are 2D - no Z coordinate
                if (newZ.HasValue)
                {
                    return false;
                }

                // Calculate target position
                double targetX = newX ?? currentCenter.X;
                double targetY = newY ?? currentCenter.Y;
                XYZ newCenter = new XYZ(targetX, targetY, currentCenter.Z);

                try
                {
                    viewport.SetBoxCenter(newCenter);
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            // Get location point (for point-based elements)
            LocationPoint locationPoint = elem.Location as LocationPoint;
            if (locationPoint != null)
            {
                XYZ currentPoint = locationPoint.Point;
                double targetX = newX ?? currentPoint.X;
                double targetY = newY ?? currentPoint.Y;
                double targetZ = newZ ?? currentPoint.Z;

                // Check if Z movement is allowed (not level-constrained)
                if (newZ.HasValue && IsLevelConstrainedForEdit(elem))
                    return false; // Cannot change Z for level-constrained elements

                XYZ newPoint = new XYZ(targetX, targetY, targetZ);
                XYZ translation = newPoint - currentPoint;

                // Move the element
                ElementTransformUtils.MoveElement(elem.Document, elem.Id, translation);
                return true;
            }

            // Get location curve (for curve-based elements)
            LocationCurve locationCurve = elem.Location as LocationCurve;
            if (locationCurve != null)
            {
                Curve curve = locationCurve.Curve;
                XYZ currentMidpoint = (curve.GetEndPoint(0) + curve.GetEndPoint(1)) / 2.0;

                double targetX = newX ?? currentMidpoint.X;
                double targetY = newY ?? currentMidpoint.Y;
                double targetZ = newZ ?? currentMidpoint.Z;

                // Check if Z movement is allowed
                if (newZ.HasValue && IsLevelConstrainedForEdit(elem))
                    return false;

                XYZ newMidpoint = new XYZ(targetX, targetY, targetZ);
                XYZ translation = newMidpoint - currentMidpoint;

                // Move the element
                ElementTransformUtils.MoveElement(elem.Document, elem.Id, translation);
                return true;
            }

            // For other elements, try moving via bounding box
            BoundingBoxXYZ bb = elem.get_BoundingBox(null);
            if (bb == null)
            {
                var options = new Options();
                var geom = elem.get_Geometry(options);
                if (geom != null)
                {
                    bb = geom.GetBoundingBox();
                }
            }

            if (bb != null)
            {
                XYZ currentCentroid = (bb.Min + bb.Max) / 2.0;
                double targetX = newX ?? currentCentroid.X;
                double targetY = newY ?? currentCentroid.Y;
                double targetZ = newZ ?? currentCentroid.Z;

                XYZ newCentroid = new XYZ(targetX, targetY, targetZ);
                XYZ translation = newCentroid - currentCentroid;

                // Move the element
                ElementTransformUtils.MoveElement(elem.Document, elem.Id, translation);
                return true;
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Check if element is level-constrained (for editing purposes)
    /// </summary>
    private static bool IsLevelConstrainedForEdit(Element elem)
    {
        // Walls are always level-constrained
        if (elem is Wall)
            return true;

        // Check for base/top constraint parameters
        Parameter baseConstraint = elem.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
        Parameter topConstraint = elem.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);

        if (baseConstraint != null && baseConstraint.HasValue)
            return true;
        if (topConstraint != null && topConstraint.HasValue)
            return true;

        // Check for level parameter with host
        Parameter levelParam = elem.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
        if (levelParam != null && levelParam.HasValue)
        {
            if (elem is FamilyInstance famInst)
            {
                if (famInst.Host != null)
                    return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Apply edits to workset entries (from SelectByWorksetsIn* commands)
    /// </summary>
    private static bool ApplyWorksetEdit(Document doc, Dictionary<string, object> entry, string columnName, object newValue)
    {
        if (!doc.IsWorkshared)
            return false;

        // Get the WorksetId from the entry (stable identifier)
        if (!entry.TryGetValue("WorksetId", out var worksetIdObj) || worksetIdObj == null)
            return false;

        if (!(worksetIdObj is WorksetId worksetId))
            return false;

        string strValue = newValue?.ToString() ?? "";
        string lowerColumnName = columnName.ToLowerInvariant();

        // Get the workset by ID (not by name, since name may have been edited in the grid)
        WorksetTable worksetTable = doc.GetWorksetTable();
        Workset targetWorkset = worksetTable.GetWorkset(worksetId);

        if (targetWorkset == null)
            return false;

        // Apply the edit based on column name
        switch (lowerColumnName)
        {
            case "workset":
            case "worksetname":
                // Rename the workset
                try
                {
                    WorksetTable.RenameWorkset(doc, targetWorkset.Id, strValue);
                    return true;
                }
                catch
                {
                    return false;
                }

            case "editable":
                // Change editable state (Editable means "owned by current user")
                bool shouldBeEditable = strValue.Equals("Yes", StringComparison.OrdinalIgnoreCase) ||
                                       strValue.Equals("True", StringComparison.OrdinalIgnoreCase);
                try
                {
                    Workset ws = worksetTable.GetWorkset(targetWorkset.Id);

                    if (shouldBeEditable && !ws.IsEditable)
                    {
                        // Make editable (checkout/borrow)
                        TransactWithCentralOptions transOpts = new TransactWithCentralOptions();
                        ISet<WorksetId> worksetIds = new HashSet<WorksetId> { ws.Id };
                        WorksharingUtils.CheckoutWorksets(doc, worksetIds, transOpts);
                    }
                    else if (!shouldBeEditable && ws.IsEditable)
                    {
                        // Relinquish ownership (make non-editable)
                        RelinquishOptions relOpts = new RelinquishOptions(false);
                        relOpts.UserWorksets = true;
                        TransactWithCentralOptions transOpts = new TransactWithCentralOptions();

                        // Relinquish this specific workset
                        WorksharingUtils.RelinquishOwnership(doc, relOpts, transOpts);
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to change editable state: {ex.Message}");
                    return false;
                }

            case "opened":
                // NOTE: The Revit API does not provide a direct method to open/close worksets at runtime.
                // The Workset class and WorksetTable class do not have OpenWorkset/CloseWorkset methods.
                // The IsOpen property is read-only and can only be set when opening the document.
                //
                // TODO: Investigate if there's an undocumented way to change this via the API
                // that matches the functionality available in the Revit UI.
                System.Diagnostics.Debug.WriteLine("Changing workset opened state is not supported by the Revit API");
                return false;

            case "visibility":
                // Change workset visibility in active view
                try
                {
                    View activeView = doc.ActiveView;
                    if (activeView == null)
                        return false;

                    WorksetVisibility newVisibility;
                    if (strValue.Equals("Shown", StringComparison.OrdinalIgnoreCase) ||
                        strValue.Equals("Visible", StringComparison.OrdinalIgnoreCase))
                    {
                        newVisibility = WorksetVisibility.Visible;
                    }
                    else if (strValue.Equals("Hidden", StringComparison.OrdinalIgnoreCase))
                    {
                        newVisibility = WorksetVisibility.Hidden;
                    }
                    else if (strValue.IndexOf("Global", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            strValue.IndexOf("Using Global", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        newVisibility = WorksetVisibility.UseGlobalSetting;
                    }
                    else
                    {
                        return false;
                    }

                    activeView.SetWorksetVisibility(targetWorkset.Id, newVisibility);
                    return true;
                }
                catch
                {
                    return false;
                }

            default:
                return false;
        }
    }

    /// <summary>
    /// Data structure to hold duplicate name detection results
    /// </summary>
    private class DuplicateNameInfo
    {
        public bool HasDuplicates { get; set; }
        public int DuplicateCount { get; set; }
        public Dictionary<string, Dictionary<string, List<long>>> DuplicateColumns { get; set; }

        public DuplicateNameInfo()
        {
            DuplicateColumns = new Dictionary<string, Dictionary<string, List<long>>>();
        }
    }

    /// <summary>
    /// Detect duplicate target names in edit operations
    /// </summary>
    private static DuplicateNameInfo DetectDuplicateNames(
        Dictionary<string, List<(long internalId, Dictionary<string, object> entry, object newValue)>> editsByColumn,
        HashSet<string> uniqueNameColumns,
        Document doc,
        List<string> diagnosticLines = null)
    {
        var result = new DuplicateNameInfo();

        foreach (var columnName in uniqueNameColumns)
        {
            var columnEdits = editsByColumn[columnName];

            // SPECIAL CASE: Detail Number is only unique per sheet, not globally
            bool isDetailNumber = columnName.Equals("Detail Number", StringComparison.OrdinalIgnoreCase);

            // SPECIAL CASE: "Name" column can apply to both Views (unique) and Sheets (not unique)
            // Only check duplicates for Views, skip ViewSheets
            bool isNameColumn = columnName.Equals("Name", StringComparison.OrdinalIgnoreCase);

            if (isDetailNumber)
            {
                diagnosticLines?.Add($"    Column '{columnName}': Checking for duplicates per sheet (scoped uniqueness)");

                // Group by parent sheet first
                var editsBySheet = new Dictionary<ElementId, List<(long internalId, Dictionary<string, object> entry, object newValue)>>();

                foreach (var edit in columnEdits)
                {
                    // Get the element to find its parent sheet
                    Element elem = GetElementFromEntry(doc, edit.entry);
                    if (elem == null) continue;

                    // For viewports, the OwnerViewId is the sheet
                    ElementId sheetId = elem.OwnerViewId ?? ElementId.InvalidElementId;

                    if (!editsBySheet.ContainsKey(sheetId))
                        editsBySheet[sheetId] = new List<(long, Dictionary<string, object>, object)>();

                    editsBySheet[sheetId].Add(edit);
                }

                // Now check for duplicates within each sheet
                foreach (var sheetGroup in editsBySheet)
                {
                    ElementId sheetId = sheetGroup.Key;
                    var editsOnSheet = sheetGroup.Value;

                    string sheetName = sheetId != ElementId.InvalidElementId
                        ? doc.GetElement(sheetId)?.Name ?? $"Sheet ID {sheetId.AsLong()}"
                        : "No Sheet";

                    var nameGroups = editsOnSheet
                        .GroupBy(e => e.newValue?.ToString()?.Trim() ?? "")
                        .Where(g => g.Count() > 1) // Only duplicates
                        .ToList();

                    if (nameGroups.Any())
                    {
                        var duplicatesInColumn = result.DuplicateColumns.ContainsKey(columnName)
                            ? result.DuplicateColumns[columnName]
                            : new Dictionary<string, List<long>>();

                        foreach (var group in nameGroups)
                        {
                            string targetName = group.Key;
                            var elementIds = group.Select(e => e.internalId).ToList();

                            // Use a unique key that includes the sheet context
                            string duplicateKey = $"{targetName} (on {sheetName})";
                            duplicatesInColumn[duplicateKey] = elementIds;
                            result.DuplicateCount += group.Count() - 1; // Count extras as duplicates

                            diagnosticLines?.Add($"    Column '{columnName}': Value '{targetName}' appears {group.Count()} times on sheet '{sheetName}'");
                            diagnosticLines?.Add($"      Element IDs: {string.Join(", ", elementIds)}");
                        }

                        result.DuplicateColumns[columnName] = duplicatesInColumn;
                    }
                }
            }
            else if (isNameColumn)
            {
                // For "Name" column: UNIVERSAL duplicate detection for all elements with unique names
                // Only ViewSheets are excluded (they can have duplicate names)
                // Includes: Views, Scope Boxes, Levels, Grids, Selection Filters, OST_Viewers, etc.
                diagnosticLines?.Add($"    Column '{columnName}': Universal duplicate check (excluding ViewSheets)");

                // Include all elements except ViewSheets
                var uniqueNameEdits = new List<(long internalId, Dictionary<string, object> entry, object newValue)>();
                foreach (var edit in columnEdits)
                {
                    Element elem = GetElementFromEntry(doc, edit.entry);
                    if (elem == null) continue;

                    // Exclude ViewSheets (they can have duplicate names)
                    if (elem is ViewSheet)
                        continue;

                    // For OST_Viewers, check if they reference a ViewSheet (exclude those too)
                    if (elem.Category?.Id.AsLong() == (int)BuiltInCategory.OST_Viewers)
                    {
                        try
                        {
                            Parameter idParam = elem.get_Parameter(BuiltInParameter.ID_PARAM);
                            if (idParam != null && idParam.HasValue)
                            {
                                ElementId referencedId = idParam.AsElementId();
                                if (referencedId != null && referencedId != ElementId.InvalidElementId)
                                {
                                    Element referencedElem = doc.GetElement(referencedId);
                                    if (referencedElem is ViewSheet)
                                        continue; // Skip markers that reference sheets
                                }
                            }
                        }
                        catch { }
                    }

                    // Include all other elements
                    uniqueNameEdits.Add(edit);
                }

                // Check for duplicates among elements that require unique names
                var nameGroups = uniqueNameEdits
                    .GroupBy(e => e.newValue?.ToString()?.Trim() ?? "")
                    .Where(g => g.Count() > 1) // Only duplicates
                    .ToList();

                if (nameGroups.Any())
                {
                    var duplicatesInColumn = new Dictionary<string, List<long>>();

                    foreach (var group in nameGroups)
                    {
                        string targetName = group.Key;
                        var elementIds = group.Select(e => e.internalId).ToList();

                        duplicatesInColumn[targetName] = elementIds;
                        result.DuplicateCount += group.Count() - 1; // Count extras as duplicates

                        diagnosticLines?.Add($"    Column '{columnName}': Name '{targetName}' appears {group.Count()} times");
                        diagnosticLines?.Add($"      Element IDs: {string.Join(", ", elementIds)}");
                    }

                    result.DuplicateColumns[columnName] = duplicatesInColumn;
                }
            }
            else
            {
                // Global uniqueness check (original logic for View Name, Sheet Number, etc.)
                var nameGroups = columnEdits
                    .GroupBy(e => e.newValue?.ToString()?.Trim() ?? "")
                    .Where(g => g.Count() > 1) // Only duplicates
                    .ToList();

                if (nameGroups.Any())
                {
                    var duplicatesInColumn = new Dictionary<string, List<long>>();

                    foreach (var group in nameGroups)
                    {
                        string targetName = group.Key;
                        var elementIds = group.Select(e => e.internalId).ToList();

                        duplicatesInColumn[targetName] = elementIds;
                        result.DuplicateCount += group.Count() - 1; // Count extras as duplicates

                        diagnosticLines?.Add($"    Column '{columnName}': Name '{targetName}' appears {group.Count()} times");
                        diagnosticLines?.Add($"      Element IDs: {string.Join(", ", elementIds)}");
                    }

                    result.DuplicateColumns[columnName] = duplicatesInColumn;
                }
            }
        }

        result.HasDuplicates = result.DuplicateColumns.Any();
        return result;
    }

    /// <summary>
    /// Show dialog to user when duplicate names are detected
    /// </summary>
    private static System.Windows.Forms.DialogResult ShowDuplicateNameDialog(DuplicateNameInfo duplicateInfo)
    {
        // Build message
        var message = new System.Text.StringBuilder();
        message.AppendLine("Duplicate names detected!");
        message.AppendLine();
        message.AppendLine($"Found {duplicateInfo.DuplicateCount} duplicate name(s):");
        message.AppendLine();

        foreach (var column in duplicateInfo.DuplicateColumns)
        {
            message.AppendLine($"Column: {column.Key}");
            foreach (var dup in column.Value.Take(5)) // Show first 5
            {
                message.AppendLine($"  • '{dup.Key}' ({dup.Value.Count} elements)");
            }
            if (column.Value.Count > 5)
            {
                message.AppendLine($"  ... and {column.Value.Count - 5} more");
            }
            message.AppendLine();
        }

        message.AppendLine("What would you like to do?");
        message.AppendLine();
        message.AppendLine("YES: Append unique suffixes (e.g. ' - a3f8e2')");
        message.AppendLine("NO: Proceed anyway (some renames may fail)");
        message.AppendLine("CANCEL: Cancel all edits");

        // Show dialog
        var result = System.Windows.Forms.MessageBox.Show(
            message.ToString(),
            "Duplicate Names Detected",
            System.Windows.Forms.MessageBoxButtons.YesNoCancel,
            System.Windows.Forms.MessageBoxIcon.Warning,
            System.Windows.Forms.MessageBoxDefaultButton.Button1);

        return result;
    }

    /// <summary>
    /// Append unique suffixes to duplicate names
    /// </summary>
    private static void AppendUniqueSuffixes(
        Dictionary<string, List<(long internalId, Dictionary<string, object> entry, object newValue)>> editsByColumn,
        DuplicateNameInfo duplicateInfo,
        List<string> diagnosticLines = null)
    {
        var random = new System.Random();

        foreach (var columnInfo in duplicateInfo.DuplicateColumns)
        {
            string columnName = columnInfo.Key;
            var duplicates = columnInfo.Value;

            diagnosticLines?.Add($"  Appending suffixes to column '{columnName}'");

            foreach (var duplicateGroup in duplicates)
            {
                string baseName = duplicateGroup.Key;
                var elementIds = duplicateGroup.Value;

                // Generate unique suffix for this group
                string suffix = $" - {random.Next(100000, 999999):x6}";

                diagnosticLines?.Add($"    Group '{baseName}' ({elementIds.Count} elements) -> suffix '{suffix}'");

                // Find and update all edits for these elements
                var columnEdits = editsByColumn[columnName];
                for (int i = 0; i < columnEdits.Count; i++)
                {
                    var edit = columnEdits[i];
                    if (elementIds.Contains(edit.internalId))
                    {
                        string originalName = edit.newValue?.ToString() ?? "";
                        string newName = originalName + suffix;

                        // Replace the edit with updated name
                        columnEdits[i] = (edit.internalId, edit.entry, newName);

                        diagnosticLines?.Add($"      Element {edit.internalId}: '{originalName}' -> '{newName}'");
                    }
                }
            }
        }
    }
}
