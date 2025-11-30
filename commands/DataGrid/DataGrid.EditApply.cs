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
            case "displayname":
            case "type":
            case "typename":
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
                // Elements like Grids, Levels, Views, etc. have Name instance parameters that should take precedence
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

            // Family properties
            case "familyname":
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
}
