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

                    // Check if this is a workset entry (from SelectByWorksetsIn* commands)
                    bool isWorksetEntry = entry.ContainsKey("WorksetId") && entry.ContainsKey("Workset");

                    if (isWorksetEntry)
                    {
                        // Apply workset-specific edits
                        foreach (var (columnName, newValue) in edits)
                        {
                            try
                            {
                                if (ApplyWorksetEdit(doc, entry, columnName, newValue))
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
        // AUTOMATIC SYSTEM: Try handler registry first
        ColumnHandlerRegistry.EnsureInitialized();
        var handler = ColumnHandlerRegistry.GetHandler(columnName);

        if (handler != null && handler.IsEditable && handler.Setter != null)
        {
            try
            {
                // Use handler to apply edit
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
                    string diagnosticPath = System.IO.Path.Combine(
                        RevitBallet.Commands.PathHelper.RuntimeDirectory,
                        "ScopeBoxEditDiagnostics.txt");

                    var diagnosticLines = new System.Collections.Generic.List<string>();
                    diagnosticLines.Add($"=== Scope Box Edit Attempt at {System.DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
                    diagnosticLines.Add($"View Name: {vScopeBox.Name}");
                    diagnosticLines.Add($"View Title: {vScopeBox.Title}");
                    diagnosticLines.Add($"View Type: {vScopeBox.ViewType}");
                    diagnosticLines.Add($"Is Template: {vScopeBox.IsTemplate}");
                    diagnosticLines.Add($"Requested Scope Box: '{strValue}'");

                    Parameter scopeBoxParam = vScopeBox.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP);
                    if (scopeBoxParam == null)
                    {
                        diagnosticLines.Add("ERROR: View does not have a scope box parameter");
                        System.IO.File.AppendAllLines(diagnosticPath, diagnosticLines);
                        return false;
                    }

                    diagnosticLines.Add($"Parameter Exists: Yes");
                    diagnosticLines.Add($"Parameter IsReadOnly: {scopeBoxParam.IsReadOnly}");
                    diagnosticLines.Add($"Parameter HasValue: {scopeBoxParam.HasValue}");

                    if (scopeBoxParam.HasValue)
                    {
                        ElementId currentScopeId = scopeBoxParam.AsElementId();
                        if (currentScopeId != null && currentScopeId != ElementId.InvalidElementId)
                        {
                            Element currentScope = elem.Document.GetElement(currentScopeId);
                            diagnosticLines.Add($"Current Scope Box: '{currentScope?.Name ?? "ERROR"}'");
                        }
                        else
                        {
                            diagnosticLines.Add($"Current Scope Box: None");
                        }
                    }

                    if (scopeBoxParam.IsReadOnly)
                    {
                        diagnosticLines.Add("ERROR: Parameter is read-only");
                        System.IO.File.AppendAllLines(diagnosticPath, diagnosticLines);
                        return false;
                    }

                    // Handle empty/none case - remove scope box
                    if (string.IsNullOrEmpty(strValue) || strValue.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            scopeBoxParam.Set(ElementId.InvalidElementId);
                            diagnosticLines.Add("SUCCESS: Cleared scope box");
                            System.IO.File.AppendAllLines(diagnosticPath, diagnosticLines);
                            return true;
                        }
                        catch (Exception ex)
                        {
                            diagnosticLines.Add($"ERROR: Failed to clear scope box: {ex.Message}");
                            System.IO.File.AppendAllLines(diagnosticPath, diagnosticLines);
                            return false;
                        }
                    }

                    // Find scope box by name
                    var scopeBoxes = new FilteredElementCollector(elem.Document)
                        .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                        .WhereElementIsNotElementType()
                        .Cast<Element>()
                        .ToList();

                    diagnosticLines.Add($"Available Scope Boxes in Document: {scopeBoxes.Count}");
                    foreach (var sb in scopeBoxes)
                    {
                        diagnosticLines.Add($"  - '{sb.Name}' (ID: {sb.Id.AsLong()})");
                    }

                    Element targetScopeBox = scopeBoxes.FirstOrDefault(sb =>
                        string.Equals(sb.Name, strValue, StringComparison.OrdinalIgnoreCase));

                    if (targetScopeBox != null)
                    {
                        diagnosticLines.Add($"Found matching scope box: '{targetScopeBox.Name}' (ID: {targetScopeBox.Id.AsLong()})");
                        try
                        {
                            scopeBoxParam.Set(targetScopeBox.Id);
                            diagnosticLines.Add($"SUCCESS: Set scope box to '{targetScopeBox.Name}'");
                            System.IO.File.AppendAllLines(diagnosticPath, diagnosticLines);
                            return true;
                        }
                        catch (Exception ex)
                        {
                            diagnosticLines.Add($"ERROR: Failed to set scope box: {ex.Message}");
                            diagnosticLines.Add($"ERROR: Stack trace: {ex.StackTrace}");
                            System.IO.File.AppendAllLines(diagnosticPath, diagnosticLines);
                            return false;
                        }
                    }
                    else
                    {
                        diagnosticLines.Add($"ERROR: No scope box named '{strValue}' found in document");
                        System.IO.File.AppendAllLines(diagnosticPath, diagnosticLines);
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
                // Log diagnostics to runtime/diagnostics directory
                string diagnosticPath = System.IO.Path.Combine(
                    RevitBallet.Commands.PathHelper.RuntimeDirectory,
                    "diagnostics",
                    $"ViewportCentroidEdit-{System.DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");

                var diagnosticLines = new System.Collections.Generic.List<string>();
                diagnosticLines.Add($"=== Viewport Centroid Edit at {System.DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ===");
                diagnosticLines.Add($"Viewport ID: {viewport.Id.AsLong()}");

                // Get owner view name
                try
                {
                    View ownerView = viewport.Document.GetElement(viewport.ViewId) as View;
                    diagnosticLines.Add($"Owner View: {ownerView?.Name ?? "Unknown"}");
                }
                catch
                {
                    diagnosticLines.Add($"Owner View: Error retrieving");
                }

                // Get current position using GetBoxCenter (what we use for editing)
                XYZ currentCenterBoxCenter = viewport.GetBoxCenter();
                diagnosticLines.Add($"Current Center (GetBoxCenter): X={currentCenterBoxCenter.X}, Y={currentCenterBoxCenter.Y}, Z={currentCenterBoxCenter.Z}");

                // Get current position using GetLabelOutline (what we use for reading)
                try
                {
                    Outline outline = viewport.GetLabelOutline();
                    XYZ minOutline = outline.MinimumPoint;
                    XYZ maxOutline = outline.MaximumPoint;
                    XYZ centerOutline = (minOutline + maxOutline) / 2.0;
                    diagnosticLines.Add($"Current Center (GetLabelOutline): X={centerOutline.X}, Y={centerOutline.Y}, Z={centerOutline.Z}");
                    diagnosticLines.Add($"Difference: dX={centerOutline.X - currentCenterBoxCenter.X}, dY={centerOutline.Y - currentCenterBoxCenter.Y}");
                }
                catch (Exception ex)
                {
                    diagnosticLines.Add($"GetLabelOutline Error: {ex.Message}");
                }

                // Get project units
                Document doc = viewport.Document;
#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                Units projectUnits = doc.GetUnits();
                FormatOptions lengthOpts = projectUnits.GetFormatOptions(SpecTypeId.Length);
                ForgeTypeId unitTypeId = lengthOpts.GetUnitTypeId();
                string unitName = unitTypeId.TypeId;

                // Convert current position to display units for logging
                double displayCurrentX = UnitUtils.ConvertFromInternalUnits(currentCenterBoxCenter.X, unitTypeId);
                double displayCurrentY = UnitUtils.ConvertFromInternalUnits(currentCenterBoxCenter.Y, unitTypeId);
#else
                // Revit 2017-2020: Use DisplayUnitType
                Units projectUnits = doc.GetUnits();
                FormatOptions lengthOpts = projectUnits.GetFormatOptions(UnitType.UT_Length);
                DisplayUnitType unitType = lengthOpts.DisplayUnits;
                string unitName = unitType.ToString();

                double displayCurrentX = UnitUtils.ConvertFromInternalUnits(currentCenterBoxCenter.X, unitType);
                double displayCurrentY = UnitUtils.ConvertFromInternalUnits(currentCenterBoxCenter.Y, unitType);
#endif
                diagnosticLines.Add($"Project Units: {unitName}");
                diagnosticLines.Add($"Current Position (Display Units): X={displayCurrentX}, Y={displayCurrentY}");

                // Log the requested changes
                diagnosticLines.Add($"Requested Changes (Internal Units): newX={newX?.ToString() ?? "null"}, newY={newY?.ToString() ?? "null"}, newZ={newZ?.ToString() ?? "null"}");

                // Calculate target position
                double targetX = newX ?? currentCenterBoxCenter.X;
                double targetY = newY ?? currentCenterBoxCenter.Y;
                diagnosticLines.Add($"Target Position (Internal Units): X={targetX}, Y={targetY}");

#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                double displayTargetX = UnitUtils.ConvertFromInternalUnits(targetX, unitTypeId);
                double displayTargetY = UnitUtils.ConvertFromInternalUnits(targetY, unitTypeId);
#else
                double displayTargetX = UnitUtils.ConvertFromInternalUnits(targetX, unitType);
                double displayTargetY = UnitUtils.ConvertFromInternalUnits(targetY, unitType);
#endif
                diagnosticLines.Add($"Target Position (Display Units): X={displayTargetX}, Y={displayTargetY}");

                // Viewports are 2D - no Z coordinate
                if (newZ.HasValue)
                {
                    diagnosticLines.Add("ERROR: Cannot set Z coordinate for viewports (2D elements on sheets)");
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
                    System.IO.File.WriteAllLines(diagnosticPath, diagnosticLines);
                    return false;
                }

                XYZ newCenter = new XYZ(targetX, targetY, currentCenterBoxCenter.Z);
                diagnosticLines.Add($"New Center to Apply: X={newCenter.X}, Y={newCenter.Y}, Z={newCenter.Z}");

                try
                {
                    viewport.SetBoxCenter(newCenter);
                    diagnosticLines.Add("SUCCESS: SetBoxCenter completed");

                    // Verify the new position
                    XYZ verifyCenter = viewport.GetBoxCenter();
                    diagnosticLines.Add($"Verified Position (GetBoxCenter): X={verifyCenter.X}, Y={verifyCenter.Y}, Z={verifyCenter.Z}");

#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                    double verifyDisplayX = UnitUtils.ConvertFromInternalUnits(verifyCenter.X, unitTypeId);
                    double verifyDisplayY = UnitUtils.ConvertFromInternalUnits(verifyCenter.Y, unitTypeId);
#else
                    double verifyDisplayX = UnitUtils.ConvertFromInternalUnits(verifyCenter.X, unitType);
                    double verifyDisplayY = UnitUtils.ConvertFromInternalUnits(verifyCenter.Y, unitType);
#endif
                    diagnosticLines.Add($"Verified Position (Display Units): X={verifyDisplayX}, Y={verifyDisplayY}");
                }
                catch (Exception ex)
                {
                    diagnosticLines.Add($"ERROR: SetBoxCenter failed: {ex.Message}");
                    diagnosticLines.Add($"Stack Trace: {ex.StackTrace}");
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
                    System.IO.File.WriteAllLines(diagnosticPath, diagnosticLines);
                    return false;
                }

                // Write diagnostics to file
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
                System.IO.File.WriteAllLines(diagnosticPath, diagnosticLines);

                return true;
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
}
