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

        Document doc = _currentUIDoc.Document;
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

        // Find the entries and apply edits
        using (Transaction trans = new Transaction(doc, "Apply DataGrid Edits"))
        {
            trans.Start();

            try
            {
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

                    // Get the Revit element
                    Element elem = GetElementFromEntry(doc, entry);
                    if (elem == null)
                    {
                        errorMessages.Add($"Could not find Revit element for entry: {GetEntryDisplayName(entry)}");
                        errorCount++;
                        continue;
                    }

                    // Apply each edit for this element
                    foreach (var (columnName, newValue) in edits)
                    {
                        try
                        {
                            if (ApplyPropertyEdit(elem, columnName, newValue, entry))
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

                trans.Commit();
                _editsWereApplied = true;
            }
            catch (Exception ex)
            {
                trans.RollBack();
                System.Windows.Forms.MessageBox.Show(
                    $"Transaction failed: {ex.Message}\n\nAll edits have been rolled back.",
                    "Apply Edits Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

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
        string lowerName = columnName.ToLowerInvariant();
        string strValue = newValue?.ToString() ?? "";

        // Built-in element properties
        switch (lowerName)
        {
            case "name":
            case "type":
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

                // If element is in a group, rename the group type instead
                if (elem.GroupId != null && elem.GroupId != ElementId.InvalidElementId && elem.GroupId.AsLong() != -1)
                {
                    Element groupElem = elem.Document.GetElement(elem.GroupId);
                    if (groupElem is Group group)
                    {
                        try
                        {
                            group.GroupType.Name = strValue;
                            return true;
                        }
                        catch
                        {
                            return false;
                        }
                    }
                }

                // Not in a group, try to rename the element directly
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

            // Room/Space properties (handled via parameters)
            case "number":
                return SetParameterValue(elem, "Number", strValue);

            case "area":
                // Area is usually read-only, but try anyway
                return SetParameterValue(elem, "Area", strValue);

            case "volume":
                // Volume is usually read-only, but try anyway
                return SetParameterValue(elem, "Volume", strValue);

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
                    if (double.TryParse(value, out double doubleValue))
                    {
                        param.Set(doubleValue);
                        return true;
                    }
                    return false;

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
        catch
        {
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
}
