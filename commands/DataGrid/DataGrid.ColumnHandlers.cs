using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

/// <summary>
/// Column handler registry and infrastructure for automatic editing
/// </summary>
public partial class CustomGUIs
{
    /// <summary>
    /// Represents a column handler that knows how to read, write, and validate a column value
    /// FULLY AUTOMATIC: DataGrid uses handlers to automatically enable editing and validation
    /// </summary>
    public class ColumnHandler
    {
        public string ColumnName { get; set; }
        public bool IsEditable { get; set; }

        /// <summary>
        /// If true, this column requires unique values across all elements (e.g., View names, Sheet numbers)
        /// When true, ApplyCellEditsToEntities will use two-phase renaming to avoid naming conflicts:
        /// Phase 1: Rename all to temporary unique names (with UUID)
        /// Phase 2: Rename to final target names
        /// </summary>
        public bool RequiresUniqueName { get; set; }

        /// <summary>
        /// Function to get column value from element
        /// </summary>
        public Func<Element, Document, object> Getter { get; set; }

        /// <summary>
        /// Function to set column value on element
        /// Returns true if successful, false otherwise
        /// </summary>
        public Func<Element, Document, object, bool> Setter { get; set; }

        /// <summary>
        /// Function to validate new value before applying
        /// Parameters: (element, document, oldValue, newValue)
        /// Returns ValidationResult with IsValid and optional ErrorMessage
        /// </summary>
        public Func<Element, Document, object, object, ValidationResult> Validator { get; set; }

        public string Description { get; set; }

        /// <summary>
        /// Validate an edit before applying
        /// </summary>
        public ValidationResult Validate(Element elem, Document doc, object oldValue, object newValue)
        {
            if (Validator == null)
                return ValidationResult.Valid();

            try
            {
                return Validator(elem, doc, oldValue, newValue);
            }
            catch (Exception ex)
            {
                return ValidationResult.Invalid($"Validation error: {ex.Message}");
            }
        }

        /// <summary>
        /// Apply an edit to an element (after validation)
        /// </summary>
        public bool ApplyEdit(Element elem, Document doc, object newValue)
        {
            if (!IsEditable || Setter == null)
                return false;

            try
            {
                return Setter(elem, doc, newValue);
            }
            catch
            {
                return false;
            }
        }
    }

    /// <summary>
    /// Registry of column handlers for automatic editing
    /// </summary>
    public static class ColumnHandlerRegistry
    {
        private static Dictionary<string, ColumnHandler> _handlers = new Dictionary<string, ColumnHandler>(StringComparer.OrdinalIgnoreCase);
        private static bool _initialized = false;

        // Delegate for batched crop region setter (set during RegisterStandardHandlers)
        private static Func<Autodesk.Revit.DB.View, Document, double?, double?, double?, double?, bool> _setCropRegionFromProjectCoords;

        /// <summary>
        /// Register a column handler
        /// </summary>
        public static void Register(ColumnHandler handler)
        {
            _handlers[handler.ColumnName] = handler;
        }

        /// <summary>
        /// Apply crop region edits in batch (all at once to avoid invalid intermediate states)
        /// </summary>
        public static bool ApplyCropRegionEdits(Element elem, Document doc, double? newTop, double? newBottom, double? newLeft, double? newRight)
        {
            if (_setCropRegionFromProjectCoords == null)
                return false;

            // Get the view from the element (handles both View and Viewport)
            Autodesk.Revit.DB.View view = null;
            if (elem is Autodesk.Revit.DB.View viewElem)
                view = viewElem;
            else if (elem is Viewport viewport)
                view = doc.GetElement(viewport.ViewId) as Autodesk.Revit.DB.View;

            if (view == null)
                return false;

            // CRITICAL: Convert display units to internal units (feet)
#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
            Units projectUnits = doc.GetUnits();
            FormatOptions lengthOpts = projectUnits.GetFormatOptions(SpecTypeId.Length);
            ForgeTypeId unitTypeId = lengthOpts.GetUnitTypeId();

            double? internalTop = newTop.HasValue ? UnitUtils.ConvertToInternalUnits(newTop.Value, unitTypeId) : (double?)null;
            double? internalBottom = newBottom.HasValue ? UnitUtils.ConvertToInternalUnits(newBottom.Value, unitTypeId) : (double?)null;
            double? internalLeft = newLeft.HasValue ? UnitUtils.ConvertToInternalUnits(newLeft.Value, unitTypeId) : (double?)null;
            double? internalRight = newRight.HasValue ? UnitUtils.ConvertToInternalUnits(newRight.Value, unitTypeId) : (double?)null;
#else
            Units projectUnits = doc.GetUnits();
            FormatOptions lengthOpts = projectUnits.GetFormatOptions(UnitType.UT_Length);
            DisplayUnitType unitType = lengthOpts.DisplayUnits;

            double? internalTop = newTop.HasValue ? UnitUtils.ConvertToInternalUnits(newTop.Value, unitType) : (double?)null;
            double? internalBottom = newBottom.HasValue ? UnitUtils.ConvertToInternalUnits(newBottom.Value, unitType) : (double?)null;
            double? internalLeft = newLeft.HasValue ? UnitUtils.ConvertToInternalUnits(newLeft.Value, unitType) : (double?)null;
            double? internalRight = newRight.HasValue ? UnitUtils.ConvertToInternalUnits(newRight.Value, unitType) : (double?)null;
#endif

            return _setCropRegionFromProjectCoords(view, doc, internalTop, internalBottom, internalLeft, internalRight);
        }

        /// <summary>
        /// Get handler for a column name (case-insensitive)
        /// Returns null if no explicit handler registered
        /// </summary>
        public static ColumnHandler GetHandler(string columnName)
        {
            _handlers.TryGetValue(columnName, out var handler);
            return handler;
        }

        /// <summary>
        /// Get handler for a column name, with fallback to dynamic parameter detection
        /// This checks if the column is an actual element parameter even if not explicitly registered
        /// </summary>
        public static ColumnHandler GetHandlerWithFallback(string columnName, Element elem, Document doc)
        {
            // Try explicit handler first
            var handler = GetHandler(columnName);
            if (handler != null)
                return handler;

            // Fallback: Check if this is an actual parameter on the element
            if (elem != null && IsElementParameter(elem, columnName))
            {
                // Return a dynamic parameter handler
                return CreateDynamicParameterHandler(columnName);
            }

            return null;
        }

        /// <summary>
        /// Check if a column is editable via registered handler
        /// </summary>
        public static bool IsColumnEditable(string columnName)
        {
            var handler = GetHandler(columnName);
            return handler != null && handler.IsEditable;
        }

        /// <summary>
        /// Check if a column is editable, with fallback to dynamic parameter detection
        /// This allows arbitrary family parameters to be editable even if not explicitly registered
        /// </summary>
        public static bool IsColumnEditableWithFallback(string columnName, Element elem)
        {
            // Try explicit handler first
            var handler = GetHandler(columnName);
            if (handler != null)
                return handler.IsEditable;

            // Fallback: Check if this is an editable parameter
            if (elem != null && IsEditableParameter(elem, columnName))
                return true;

            return false;
        }

        /// <summary>
        /// Get all registered column names
        /// </summary>
        public static IEnumerable<string> GetAllColumnNames()
        {
            return _handlers.Keys;
        }

        /// <summary>
        /// Clear all registered handlers (useful for testing)
        /// </summary>
        public static void Clear()
        {
            _handlers.Clear();
            _initialized = false;
        }

        /// <summary>
        /// Check if a column name corresponds to an actual parameter on the element
        /// </summary>
        private static bool IsElementParameter(Element elem, string columnName)
        {
            if (elem == null) return false;

            // Try exact match
            Parameter param = elem.LookupParameter(columnName);
            if (param != null) return true;

            // Try case-insensitive match
            foreach (Parameter p in elem.Parameters)
            {
                if (string.Equals(p.Definition.Name, columnName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            // Try with space/underscore conversion
            string colWithSpaces = columnName.Replace("_", " ");
            param = elem.LookupParameter(colWithSpaces);
            if (param != null) return true;

            string colWithUnderscores = columnName.Replace(" ", "_");
            param = elem.LookupParameter(colWithUnderscores);
            if (param != null) return true;

            return false;
        }

        /// <summary>
        /// Check if a parameter on an element is editable (not read-only)
        /// </summary>
        private static bool IsEditableParameter(Element elem, string columnName)
        {
            if (elem == null) return false;

            Parameter param = FindParameter(elem, columnName);
            return param != null && !param.IsReadOnly;
        }

        /// <summary>
        /// Find a parameter by name with fuzzy matching
        /// </summary>
        private static Parameter FindParameter(Element elem, string columnName)
        {
            if (elem == null) return null;

            // Try exact match
            Parameter param = elem.LookupParameter(columnName);
            if (param != null) return param;

            // Try case-insensitive match
            foreach (Parameter p in elem.Parameters)
            {
                if (string.Equals(p.Definition.Name, columnName, StringComparison.OrdinalIgnoreCase))
                    return p;
            }

            // Try with space/underscore conversion
            string colWithSpaces = columnName.Replace("_", " ");
            param = elem.LookupParameter(colWithSpaces);
            if (param != null) return param;

            string colWithUnderscores = columnName.Replace(" ", "_");
            param = elem.LookupParameter(colWithUnderscores);
            if (param != null) return param;

            return null;
        }

        /// <summary>
        /// Create a dynamic parameter handler for arbitrary parameters
        /// </summary>
        private static ColumnHandler CreateDynamicParameterHandler(string columnName)
        {
            return new ColumnHandler
            {
                ColumnName = columnName,
                IsEditable = true,
                Description = $"Dynamic parameter: {columnName}",

                Getter = (elem, doc) =>
                {
                    Parameter param = FindParameter(elem, columnName);
                    if (param == null) return null;

                    // Return value as string (with units if applicable)
                    return param.AsValueString() ?? param.AsString() ?? "";
                },

                Setter = (elem, doc, newValue) =>
                {
                    Parameter param = FindParameter(elem, columnName);
                    if (param == null || param.IsReadOnly)
                        return false;

                    string strValue = newValue?.ToString() ?? "";

                    try
                    {
                        switch (param.StorageType)
                        {
                            case StorageType.String:
                                param.Set(strValue);
                                return true;

                            case StorageType.Integer:
                                // Check if this is a Yes/No parameter
                                bool isYesNoParameter = false;
#if REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                                try
                                {
                                    // Revit 2022+: Use GetDataType() and SpecTypeId.Boolean
                                    var dataType = param.Definition.GetDataType();
                                    isYesNoParameter = dataType != null && dataType == SpecTypeId.Boolean.YesNo;
                                }
                                catch { }
#else
                                // Revit 2017-2021: Use ParameterType enum
                                try
                                {
                                    isYesNoParameter = param.Definition.ParameterType == ParameterType.YesNo;
                                }
                                catch { }
#endif

                                if (isYesNoParameter)
                                {
                                    // Yes/No parameter - interpret text values
                                    string normalized = strValue.Trim().ToLowerInvariant();
                                    int yesNoValue;

                                    if (normalized == "yes" || normalized == "true" || normalized == "1")
                                        yesNoValue = 1;
                                    else if (normalized == "no" || normalized == "false" || normalized == "0")
                                        yesNoValue = 0;
                                    else
                                        return false; // Invalid value for Yes/No parameter

                                    param.Set(yesNoValue);
                                    return true;
                                }
                                else
                                {
                                    // Regular integer parameter
                                    if (int.TryParse(strValue, out int intValue))
                                    {
                                        param.Set(intValue);
                                        return true;
                                    }
                                    return false;
                                }

                            case StorageType.Double:
                                // Use SetValueString to handle unit conversion automatically
                                try
                                {
                                    param.SetValueString(strValue);
                                    return true;
                                }
                                catch
                                {
                                    // Fallback to direct parsing
                                    if (double.TryParse(strValue, out double doubleValue))
                                    {
                                        param.Set(doubleValue);
                                        return true;
                                    }
                                    return false;
                                }

                            case StorageType.ElementId:
                                if (int.TryParse(strValue, out int idValue))
                                {
                                    param.Set(idValue.ToElementId());
                                    return true;
                                }
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
            };
        }

        /// <summary>
        /// Register standard Revit column handlers (Family, Type Name, etc.)
        /// This is called automatically when first accessed
        /// </summary>
        public static void RegisterStandardHandlers()
        {
            if (_initialized)
                return;

            _initialized = true;

            // Family Name
            Register(new ColumnHandler
            {
                ColumnName = "Family",
                IsEditable = true,
                Description = "Family name for family instances or types",
                Getter = (elem, doc) =>
                {
                    // Handle case where elem IS the type (e.g., SelectByFamilyTypes commands)
                    if (elem is FamilySymbol directSymbol)
                        return directSymbol.FamilyName;

                    // Handle case where elem is an instance - get its type
                    ElementId typeId = elem.GetTypeId();
                    if (typeId == null || typeId == ElementId.InvalidElementId)
                        return null;

                    Element typeElement = doc.GetElement(typeId);
                    if (typeElement is FamilySymbol familySymbol)
                        return familySymbol.FamilyName;

                    Parameter familyParam = typeElement?.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                    return familyParam?.AsString() ?? "System Type";
                },
                Validator = ColumnValidators.All(
                    ColumnValidators.NotEmpty,
                    ColumnValidators.NoInvalidCharacters,
                    ColumnValidators.NoLeadingTrailingWhitespace
                ),
                Setter = (elem, doc, newValue) =>
                {
                    string strValue = newValue?.ToString() ?? "";

                    // Handle case where elem IS the type (e.g., SelectByFamilyTypes commands)
                    if (elem is FamilySymbol directSymbol)
                    {
                        try
                        {
                            Family family = directSymbol.Family;
                            family.Name = strValue;
                            return true;
                        }
                        catch { return false; }
                    }

                    // Handle case where elem is an instance - get its type
                    ElementId typeId = elem.GetTypeId();
                    if (typeId == null || typeId == ElementId.InvalidElementId)
                        return false;

                    Element typeElement = doc.GetElement(typeId);
                    if (typeElement is FamilySymbol familySymbol)
                    {
                        try
                        {
                            Family family = familySymbol.Family;
                            family.Name = strValue;
                            return true;
                        }
                        catch { return false; }
                    }
                    return false;
                }
            });

            // Type Name
            Register(new ColumnHandler
            {
                ColumnName = "Type Name",
                IsEditable = true,
                Description = "Type name for elements with types or types themselves",
                Validator = ColumnValidators.All(
                    ColumnValidators.NotEmpty,
                    ColumnValidators.NoInvalidCharacters,
                    ColumnValidators.NoLeadingTrailingWhitespace
                ),
                Getter = (elem, doc) =>
                {
                    // Handle case where elem IS the type (e.g., SelectByFamilyTypes commands)
                    if (elem is ElementType directType)
                        return directType.Name;

                    // Handle case where elem is an instance - get its type
                    ElementId typeId = elem.GetTypeId();
                    if (typeId == null || typeId == ElementId.InvalidElementId)
                        return null;

                    Element typeElement = doc.GetElement(typeId);
                    return typeElement?.Name;
                },
                Setter = (elem, doc, newValue) =>
                {
                    string strValue = newValue?.ToString() ?? "";

                    // Handle case where elem IS the type (e.g., SelectByFamilyTypes commands)
                    if (elem is ElementType directType)
                    {
                        try
                        {
                            directType.Name = strValue;
                            return true;
                        }
                        catch { return false; }
                    }

                    // Handle case where elem is an instance - get its type
                    ElementId typeId = elem.GetTypeId();
                    if (typeId == null || typeId == ElementId.InvalidElementId)
                        return false;

                    Element typeElement = doc.GetElement(typeId);
                    if (typeElement != null)
                    {
                        try
                        {
                            typeElement.Name = strValue;
                            return true;
                        }
                        catch { return false; }
                    }
                    return false;
                }
            });

            // Comments
            Register(new ColumnHandler
            {
                ColumnName = "Comments",
                IsEditable = true,
                Description = "Element comments parameter",
                Validator = ColumnValidators.MaxLength(1024),
                Getter = (elem, doc) =>
                {
                    Parameter param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                    return param?.AsString() ?? "";
                },
                Setter = (elem, doc, newValue) =>
                {
                    string strValue = newValue?.ToString() ?? "";
                    Parameter param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                    if (param != null && !param.IsReadOnly)
                    {
                        try
                        {
                            param.Set(strValue);
                            return true;
                        }
                        catch { return false; }
                    }
                    return false;
                }
            });

            // Mark
            Register(new ColumnHandler
            {
                ColumnName = "Mark",
                IsEditable = true,
                Description = "Element mark parameter",
                Validator = ColumnValidators.All(
                    ColumnValidators.NoInvalidCharacters,
                    ColumnValidators.MaxLength(256)
                ),
                Getter = (elem, doc) =>
                {
                    Parameter param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                    return param?.AsString() ?? "";
                },
                Setter = (elem, doc, newValue) =>
                {
                    string strValue = newValue?.ToString() ?? "";
                    Parameter param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                    if (param != null && !param.IsReadOnly)
                    {
                        try
                        {
                            param.Set(strValue);
                            return true;
                        }
                        catch { return false; }
                    }
                    return false;
                }
            });

            // Helper method to get view from element (view or viewport)
            Func<Element, Document, Autodesk.Revit.DB.View> GetViewFromElement = (elem, doc) =>
            {
                if (elem is Autodesk.Revit.DB.View view)
                    return view;
                if (elem is Viewport viewport)
                    return doc.GetElement(viewport.ViewId) as Autodesk.Revit.DB.View;
                return null;
            };

            // Helper method to check if crop region is rectangular
            Func<Autodesk.Revit.DB.View, bool> IsRectangularCropRegion = (view) =>
            {
                if (view == null || view.CropBox == null || !view.CropBoxActive)
                    return false;

                if (view is ViewPlan plan)
                {
                    try
                    {
                        var managerProperty = plan.GetType().GetProperty("CropRegionShapeManager");
                        if (managerProperty != null)
                        {
                            object manager = managerProperty.GetValue(plan, null);
                            if (manager != null)
                            {
                                var method = manager.GetType().GetMethod("GetCropRegionShape");
                                if (method != null)
                                {
                                    CurveLoop cropLoop = method.Invoke(manager, null) as CurveLoop;
                                    if (cropLoop != null)
                                    {
                                        return cropLoop.ToList().Count == 4;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        return false;
                    }
                }
                // For other view types (section, elevation), assume rectangular
                return true;
            };

            // Helper to convert from internal units to display units
            Func<double, Document, double> ConvertFromInternal = (internalValue, doc) =>
            {
#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                Units projectUnits = doc.GetUnits();
                FormatOptions lengthOpts = projectUnits.GetFormatOptions(SpecTypeId.Length);
                ForgeTypeId unitTypeId = lengthOpts.GetUnitTypeId();
                return UnitUtils.ConvertFromInternalUnits(internalValue, unitTypeId);
#else
                Units projectUnits = doc.GetUnits();
                FormatOptions lengthOpts = projectUnits.GetFormatOptions(UnitType.UT_Length);
                DisplayUnitType unitType = lengthOpts.DisplayUnits;
                return UnitUtils.ConvertFromInternalUnits(internalValue, unitType);
#endif
            };

            // Helper to convert from display units to internal units
            Func<double, Document, double> ConvertToInternal = (displayValue, doc) =>
            {
#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                Units projectUnits = doc.GetUnits();
                FormatOptions lengthOpts = projectUnits.GetFormatOptions(SpecTypeId.Length);
                ForgeTypeId unitTypeId = lengthOpts.GetUnitTypeId();
                return UnitUtils.ConvertToInternalUnits(displayValue, unitTypeId);
#else
                Units projectUnits = doc.GetUnits();
                FormatOptions lengthOpts = projectUnits.GetFormatOptions(UnitType.UT_Length);
                DisplayUnitType unitType = lengthOpts.DisplayUnits;
                return UnitUtils.ConvertToInternalUnits(displayValue, unitType);
#endif
            };

            // Helper to get survey point elevation offset
            Func<Document, double> GetSurveyPointElevation = (doc) =>
            {
                try
                {
                    // Get survey point
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    var surveyPoint = collector.OfCategory(BuiltInCategory.OST_SharedBasePoint)
                        .FirstOrDefault() as BasePoint;

                    if (surveyPoint != null)
                    {
                        // Get the survey point's elevation (Z coordinate in project coordinates)
                        return surveyPoint.get_BoundingBox(null).Min.Z;
                    }
                }
                catch { }

                return 0.0; // Fallback to project origin
            };

            // Helper to get crop region values in project coordinates relative to survey point
            Func<Autodesk.Revit.DB.View, Document, (double top, double bottom, double left, double right)?> GetCropRegionInProjectCoords = (view, doc) =>
            {
                if (view == null || !view.CropBoxActive || view.CropBox == null)
                    return null;

                BoundingBoxXYZ bbox = view.CropBox;
                Transform transform = bbox.Transform;
                double surveyElevation = GetSurveyPointElevation(doc);

                // Transform corner points to project coordinates
                XYZ minCorner = transform.OfPoint(bbox.Min);
                XYZ maxCorner = transform.OfPoint(bbox.Max);

                // For plan views: use X/Y from project coordinates
                // For elevation/section views: use Z (elevation) from project coordinates for top/bottom
                bool isElevationOrSection = view.ViewType == ViewType.Elevation || view.ViewType == ViewType.Section;

                if (isElevationOrSection)
                {
                    // Top/Bottom are elevations (Z in project coordinates), relative to survey point
                    double top = maxCorner.Z - surveyElevation;
                    double bottom = minCorner.Z - surveyElevation;
                    // Left/Right are horizontal positions in the view direction
                    double left = minCorner.X;
                    double right = maxCorner.X;
                    return (top, bottom, left, right);
                }
                else
                {
                    // Plan view: Top/Bottom are Y, Left/Right are X in project coordinates
                    double top = maxCorner.Y;
                    double bottom = minCorner.Y;
                    double left = minCorner.X;
                    double right = maxCorner.X;
                    return (top, bottom, left, right);
                }
            };

            // Helper to set crop region from project coordinates
            Func<Autodesk.Revit.DB.View, Document, double?, double?, double?, double?, bool> SetCropRegionFromProjectCoords = (view, doc, newTop, newBottom, newLeft, newRight) =>
            {
                if (view == null || !view.CropBoxActive || view.CropBox == null)
                    return false;

                try
                {
                    // Get current crop box and transform - used for depth/height preservation
                    BoundingBoxXYZ currentBbox = view.CropBox;
                    Transform transform = currentBbox.Transform;
                    double surveyElevation = GetSurveyPointElevation(doc);

                    // Get current values in project coordinates
                    var current = GetCropRegionInProjectCoords(view, doc);
                    if (!current.HasValue)
                        return false;

                    bool isElevationOrSection = view.ViewType == ViewType.Elevation || view.ViewType == ViewType.Section;

                    // Use provided values or keep current
                    double top = newTop ?? current.Value.top;
                    double bottom = newBottom ?? current.Value.bottom;
                    double left = newLeft ?? current.Value.left;
                    double right = newRight ?? current.Value.right;

                    // Transform current corners to project space to get depth dimension
                    XYZ currentMinProject = transform.OfPoint(currentBbox.Min);
                    XYZ currentMaxProject = transform.OfPoint(currentBbox.Max);

                    // Build new corners in project space
                    Transform inverseTransform = transform.Inverse;
                    XYZ newMinProject, newMaxProject;

                    if (isElevationOrSection)
                    {
                        // For elevation/section: modify X (left/right) and Z (elevation)
                        // Keep Y (depth) from current bbox
                        newMinProject = new XYZ(left, currentMinProject.Y, bottom + surveyElevation);
                        newMaxProject = new XYZ(right, currentMaxProject.Y, top + surveyElevation);
                    }
                    else
                    {
                        // For plan: modify X and Y, keep Z (height) from current
                        newMinProject = new XYZ(left, bottom, currentMinProject.Z);
                        newMaxProject = new XYZ(right, top, currentMaxProject.Z);
                    }

                    // Transform back to view space
                    XYZ newMinView = inverseTransform.OfPoint(newMinProject);
                    XYZ newMaxView = inverseTransform.OfPoint(newMaxProject);

                    // Create NEW BoundingBoxXYZ to avoid Revit API quirks with modified objects
                    BoundingBoxXYZ newBbox = new BoundingBoxXYZ();
                    newBbox.Transform = transform;
                    newBbox.Min = newMinView;
                    newBbox.Max = newMaxView;
                    view.CropBox = newBbox;

                    return true;
                }
                catch
                {
                    return false;
                }
            };

            // Save lambda to static field for batched crop region edits
            _setCropRegionFromProjectCoords = SetCropRegionFromProjectCoords;

            // Crop Region Top
            Register(new ColumnHandler
            {
                ColumnName = "Crop Region Top",
                IsEditable = true,
                Description = "Top position of crop region (elevation for sections/elevations, Y for plans) - relative to survey point",
                Getter = (elem, doc) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null || !IsRectangularCropRegion(view))
                        return null;

                    var coords = GetCropRegionInProjectCoords(view, doc);
                    if (!coords.HasValue)
                        return null;

                    return ConvertFromInternal(coords.Value.top, doc);
                },
                Setter = (elem, doc, newValue) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null || !IsRectangularCropRegion(view))
                        return false;

                    if (!double.TryParse(newValue?.ToString(), out double displayValue))
                        return false;

                    double internalValue = ConvertToInternal(displayValue, doc);
                    return SetCropRegionFromProjectCoords(view, doc, internalValue, null, null, null);
                }
            });

            // Crop Region Bottom
            Register(new ColumnHandler
            {
                ColumnName = "Crop Region Bottom",
                IsEditable = true,
                Description = "Bottom position of crop region (elevation for sections/elevations, Y for plans) - relative to survey point",
                Getter = (elem, doc) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null || !IsRectangularCropRegion(view))
                        return null;

                    var coords = GetCropRegionInProjectCoords(view, doc);
                    if (!coords.HasValue)
                        return null;

                    return ConvertFromInternal(coords.Value.bottom, doc);
                },
                Setter = (elem, doc, newValue) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null || !IsRectangularCropRegion(view))
                        return false;

                    if (!double.TryParse(newValue?.ToString(), out double displayValue))
                        return false;

                    double internalValue = ConvertToInternal(displayValue, doc);
                    return SetCropRegionFromProjectCoords(view, doc, null, internalValue, null, null);
                }
            });

            // Crop Region Left
            Register(new ColumnHandler
            {
                ColumnName = "Crop Region Left",
                IsEditable = true,
                Description = "Left position of crop region in project coordinates",
                Getter = (elem, doc) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null || !IsRectangularCropRegion(view))
                        return null;

                    var coords = GetCropRegionInProjectCoords(view, doc);
                    if (!coords.HasValue)
                        return null;

                    return ConvertFromInternal(coords.Value.left, doc);
                },
                Setter = (elem, doc, newValue) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null || !IsRectangularCropRegion(view))
                        return false;

                    if (!double.TryParse(newValue?.ToString(), out double displayValue))
                        return false;

                    double internalValue = ConvertToInternal(displayValue, doc);
                    return SetCropRegionFromProjectCoords(view, doc, null, null, internalValue, null);
                }
            });

            // Crop Region Right
            Register(new ColumnHandler
            {
                ColumnName = "Crop Region Right",
                IsEditable = true,
                Description = "Right position of crop region in project coordinates",
                Getter = (elem, doc) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null || !IsRectangularCropRegion(view))
                        return null;

                    var coords = GetCropRegionInProjectCoords(view, doc);
                    if (!coords.HasValue)
                        return null;

                    return ConvertFromInternal(coords.Value.right, doc);
                },
                Setter = (elem, doc, newValue) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null || !IsRectangularCropRegion(view))
                        return false;

                    if (!double.TryParse(newValue?.ToString(), out double displayValue))
                        return false;

                    double internalValue = ConvertToInternal(displayValue, doc);
                    return SetCropRegionFromProjectCoords(view, doc, null, null, null, internalValue);
                }
            });

            // Crop View (CropBoxActive)
            Register(new ColumnHandler
            {
                ColumnName = "Crop View",
                IsEditable = true,
                Description = "Whether crop region is active/enabled for the view",
                Getter = (elem, doc) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null) return null;
                    return view.CropBoxActive;
                },
                Setter = (elem, doc, newValue) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null) return false;

                    bool boolValue;
                    if (newValue is bool b)
                        boolValue = b;
                    else if (bool.TryParse(newValue?.ToString(), out bool parsed))
                        boolValue = parsed;
                    else
                        return false;

                    try
                    {
                        view.CropBoxActive = boolValue;
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }
            });

            // Crop Region Visible (CropBoxVisible)
            Register(new ColumnHandler
            {
                ColumnName = "Crop Region Visible",
                IsEditable = true,
                Description = "Whether crop region boundary is visible in the view",
                Getter = (elem, doc) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null) return null;
                    return view.CropBoxVisible;
                },
                Setter = (elem, doc, newValue) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null) return false;

                    bool boolValue;
                    if (newValue is bool b)
                        boolValue = b;
                    else if (bool.TryParse(newValue?.ToString(), out bool parsed))
                        boolValue = parsed;
                    else
                        return false;

                    try
                    {
                        view.CropBoxVisible = boolValue;
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }
            });

            // Annotation Crop
            Register(new ColumnHandler
            {
                ColumnName = "Annotation Crop",
                IsEditable = true,
                Description = "Whether annotation crop is active for the view",
                Getter = (elem, doc) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null) return null;

                    try
                    {
                        // Try to get via parameter (may not be available in all Revit versions)
                        Parameter param = view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE);
                        if (param != null)
                            return param.AsInteger() == 1;
                    }
                    catch { }

                    return null;
                },
                Setter = (elem, doc, newValue) =>
                {
                    var view = GetViewFromElement(elem, doc);
                    if (view == null) return false;

                    bool boolValue;
                    if (newValue is bool b)
                        boolValue = b;
                    else if (bool.TryParse(newValue?.ToString(), out bool parsed))
                        boolValue = parsed;
                    else
                        return false;

                    try
                    {
                        // Try to set via parameter (may not be available in all Revit versions)
                        Parameter param = view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE);
                        if (param != null && !param.IsReadOnly)
                        {
                            param.Set(boolValue ? 1 : 0);
                            return true;
                        }
                    }
                    catch { }

                    return false;
                }
            });

            // Name column - UNIVERSAL handler for any element with a Name property
            // CRITICAL: RequiresUniqueName=true for two-phase rename to avoid "element with that name already exists" errors
            // This handler works universally for: Views, Scope Boxes, Levels, Grids, Selection Sets, and any other named element
            // Special handling for: ViewSheets (use SHEET_NAME param), OST_Viewers (resolve to underlying view), Groups (rename GroupType)
            Register(new ColumnHandler
            {
                ColumnName = "Name",
                IsEditable = true,
                RequiresUniqueName = true,
                Description = "Element name (supports two-phase rename for swapping)",
                Validator = ColumnValidators.NotEmpty,
                Getter = (elem, doc) =>
                {
                    // Return Name for any element - this is universal
                    return elem.Name;
                },
                Setter = (elem, doc, newValue) =>
                {
                    string strValue = newValue?.ToString() ?? "";

                    // SPECIAL CASE: ViewSheets use SHEET_NAME parameter (not Name property)
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
                        catch { return false; }
                    }

                    // SPECIAL CASE: Groups - rename the GroupType, not the instance
                    if (elem is Group group)
                    {
                        try
                        {
                            group.GroupType.Name = strValue;
                            return true;
                        }
                        catch { return false; }
                    }

                    // SPECIAL CASE: OST_Viewers (Callouts, Sections, Elevations)
                    // These are marker elements that reference actual views
                    if (elem.Category?.Id.AsLong() == (int)BuiltInCategory.OST_Viewers)
                    {
                        // First, try VIEW_NAME parameter (works for some view types)
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

                        // Check if this marker references an actual view via ID_PARAM
                        try
                        {
                            Parameter idParam = elem.get_Parameter(BuiltInParameter.ID_PARAM);
                            if (idParam != null && idParam.HasValue)
                            {
                                ElementId referencedId = idParam.AsElementId();
                                if (referencedId != null && referencedId != ElementId.InvalidElementId)
                                {
                                    Element referencedElem = doc.GetElement(referencedId);
                                    if (referencedElem is Autodesk.Revit.DB.View referencedView)
                                    {
                                        referencedView.Name = strValue;
                                        return true;
                                    }
                                }
                            }
                        }
                        catch { }
                    }

                    // UNIVERSAL FALLBACK: Try to set Name property directly
                    // This works for Views, Scope Boxes, Levels, Grids, Selection Filters, and most other named elements
                    try
                    {
                        elem.Name = strValue;
                        return true;
                    }
                    catch { }

                    // LAST RESORT: Check for writable "Name" instance parameter
                    // Some elements have Name as a parameter rather than a property
                    try
                    {
                        Parameter nameParam = elem.LookupParameter("Name");
                        if (nameParam != null && !nameParam.IsReadOnly && nameParam.StorageType == StorageType.String)
                        {
                            nameParam.Set(strValue);
                            return true;
                        }
                    }
                    catch { }

                    return false;
                }
            });

            // View Name (for regular views and viewports, excluding sheets)
            // CRITICAL: RequiresUniqueName=true for two-phase rename to avoid "view with that name already exists" errors
            // Also handles OST_Viewers (elevation markers, section markers) which reference actual views
            Register(new ColumnHandler
            {
                ColumnName = "View Name",
                IsEditable = true,
                RequiresUniqueName = true,
                Description = "View name (unique per document) - works with Views, Viewports, and view markers",
                Validator = ColumnValidators.NotEmpty,
                Getter = (elem, doc) =>
                {
                    // Handle Viewports - get the underlying view
                    if (elem is Viewport viewport)
                    {
                        Autodesk.Revit.DB.View underlyingView = doc.GetElement(viewport.ViewId) as Autodesk.Revit.DB.View;
                        return underlyingView?.Name;
                    }

                    // Handle regular Views (non-sheets)
                    if (elem is Autodesk.Revit.DB.View view && !(view is ViewSheet))
                        return view.Name;

                    return null;
                },
                Setter = (elem, doc, newValue) =>
                {
                    string strValue = newValue?.ToString() ?? "";

                    // Handle Viewports - rename the underlying view
                    if (elem is Viewport viewport)
                    {
                        try
                        {
                            Autodesk.Revit.DB.View underlyingView = doc.GetElement(viewport.ViewId) as Autodesk.Revit.DB.View;
                            if (underlyingView != null && !(underlyingView is ViewSheet))
                            {
                                underlyingView.Name = strValue;
                                return true;
                            }
                        }
                        catch { return false; }
                    }

                    // Handle regular Views (non-sheets)
                    if (elem is Autodesk.Revit.DB.View view && !(view is ViewSheet))
                    {
                        try
                        {
                            view.Name = strValue;
                            return true;
                        }
                        catch { return false; }
                    }

                    // Handle OST_Viewers (Callouts, Sections, Elevations)
                    // These elements appear in category "Views" but may not be castable to View class
                    if (elem.Category?.Id.AsLong() == (int)BuiltInCategory.OST_Viewers)
                    {
                        // First, try VIEW_NAME parameter
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

                        // Check if this is a reference element that points to an actual view
                        try
                        {
                            Parameter idParam = elem.get_Parameter(BuiltInParameter.ID_PARAM);
                            if (idParam != null && idParam.HasValue)
                            {
                                ElementId referencedId = idParam.AsElementId();
                                if (referencedId != null && referencedId != ElementId.InvalidElementId)
                                {
                                    Element referencedElem = doc.GetElement(referencedId);
                                    if (referencedElem is Autodesk.Revit.DB.View referencedView && !(referencedView is ViewSheet))
                                    {
                                        referencedView.Name = strValue;
                                        return true;
                                    }
                                }
                            }
                        }
                        catch { return false; }
                    }

                    return false;
                }
            });

            // Sheet Name (for ViewSheets)
            // NOTE: Sheet names do NOT need to be unique (only sheet numbers do)
            Register(new ColumnHandler
            {
                ColumnName = "Sheet Name",
                IsEditable = true,
                RequiresUniqueName = false,
                Description = "Sheet name (can be duplicated)",
                Validator = ColumnValidators.NotEmpty,
                Getter = (elem, doc) =>
                {
                    if (elem is ViewSheet sheet)
                    {
                        Parameter param = sheet.get_Parameter(BuiltInParameter.SHEET_NAME);
                        return param?.AsString() ?? "";
                    }
                    return null;
                },
                Setter = (elem, doc, newValue) =>
                {
                    string strValue = newValue?.ToString() ?? "";
                    if (elem is ViewSheet sheet)
                    {
                        try
                        {
                            Parameter param = sheet.get_Parameter(BuiltInParameter.SHEET_NAME);
                            if (param != null && !param.IsReadOnly)
                            {
                                param.Set(strValue);
                                return true;
                            }
                        }
                        catch { return false; }
                    }
                    return false;
                }
            });

            // Sheet Number (for ViewSheets)
            // CRITICAL: RequiresUniqueName=true for two-phase rename
            Register(new ColumnHandler
            {
                ColumnName = "Sheet Number",
                IsEditable = true,
                RequiresUniqueName = true,
                Description = "Sheet number (unique per document)",
                Validator = ColumnValidators.NotEmpty,
                Getter = (elem, doc) =>
                {
                    if (elem is ViewSheet sheet)
                        return sheet.SheetNumber;
                    return null;
                },
                Setter = (elem, doc, newValue) =>
                {
                    string strValue = newValue?.ToString() ?? "";
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
                }
            });

            // Detail Number (for Viewports)
            // CRITICAL: RequiresUniqueName=true for two-phase rename
            // NOTE: Uses numeric temp values instead of UUID strings
            Register(new ColumnHandler
            {
                ColumnName = "Detail Number",
                IsEditable = true,
                RequiresUniqueName = true,
                Description = "Detail number for viewport (unique per sheet)",
                Validator = ColumnValidators.NotEmpty,
                Getter = (elem, doc) =>
                {
                    if (elem is Viewport viewport)
                    {
                        Parameter param = viewport.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER);
                        return param?.AsString() ?? "";
                    }
                    return null;
                },
                Setter = (elem, doc, newValue) =>
                {
                    string strValue = newValue?.ToString() ?? "";
                    if (elem is Viewport viewport)
                    {
                        try
                        {
                            Parameter param = viewport.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER);
                            if (param != null && !param.IsReadOnly)
                            {
                                param.Set(strValue);
                                return true;
                            }
                        }
                        catch { return false; }
                    }
                    return false;
                }
            });

            // X Centroid
            Register(new ColumnHandler
            {
                ColumnName = "X Centroid",
                IsEditable = true,
                Description = "X coordinate of element centroid in project coordinates",
                Getter = (elem, doc) =>
                {
                    // Getter is handled by ElementDataHelper - just return null here
                    // The actual value is computed during data collection
                    return null;
                },
                Setter = (elem, doc, newValue) =>
                {
                    if (!double.TryParse(newValue?.ToString(), out double displayValue))
                        return false;

                    // Convert from display units to internal units (feet)
                    double targetX = ConvertToInternal(displayValue, doc);

                    // Get current X coordinate
                    XYZ currentPos = GetElementPosition(elem, doc);
                    if (currentPos == null)
                        return false;

                    // Calculate offset
                    double deltaX = targetX - currentPos.X;
                    XYZ translation = new XYZ(deltaX, 0, 0);

                    // Move element
                    return MoveElement(elem, doc, translation);
                }
            });

            // Y Centroid
            Register(new ColumnHandler
            {
                ColumnName = "Y Centroid",
                IsEditable = true,
                Description = "Y coordinate of element centroid in project coordinates",
                Getter = (elem, doc) =>
                {
                    // Getter is handled by ElementDataHelper - just return null here
                    // The actual value is computed during data collection
                    return null;
                },
                Setter = (elem, doc, newValue) =>
                {
                    if (!double.TryParse(newValue?.ToString(), out double displayValue))
                        return false;

                    // Convert from display units to internal units (feet)
                    double targetY = ConvertToInternal(displayValue, doc);

                    // Get current Y coordinate
                    XYZ currentPos = GetElementPosition(elem, doc);
                    if (currentPos == null)
                        return false;

                    // Calculate offset
                    double deltaY = targetY - currentPos.Y;
                    XYZ translation = new XYZ(0, deltaY, 0);

                    // Move element
                    return MoveElement(elem, doc, translation);
                }
            });
        }

        /// <summary>
        /// Get the position (centroid) of an element
        /// Returns XYZ in internal units (feet)
        /// CRITICAL: Must match GetElementCentroid logic in FilterSelected.cs
        /// </summary>
        private static XYZ GetElementPosition(Element elem, Document doc)
        {
            // Viewport elements - use GetBoxOutline center
            if (elem is Viewport viewport)
            {
                try
                {
                    Outline outline = viewport.GetBoxOutline();
                    XYZ min = outline.MinimumPoint;
                    XYZ max = outline.MaximumPoint;
                    return (min + max) / 2.0;
                }
                catch { }
            }

            // LocationPoint elements
            if (elem.Location is LocationPoint locationPoint)
            {
                return locationPoint.Point;
            }

            // LocationCurve elements - use midpoint
            if (elem.Location is LocationCurve locationCurve)
            {
                Curve curve = locationCurve.Curve;
                return (curve.GetEndPoint(0) + curve.GetEndPoint(1)) / 2.0;
            }

            // For elements without location (including TextNote), use bounding box centroid
            // IMPORTANT: Do NOT use TextNote.Coord as it's the insertion point, not the visual center
            BoundingBoxXYZ bb = elem.get_BoundingBox(null);

            // For elements on sheets/views (including TextNote), try with owner view
            if (bb == null && elem.OwnerViewId != null && elem.OwnerViewId != ElementId.InvalidElementId)
            {
                try
                {
                    View ownerView = doc.GetElement(elem.OwnerViewId) as View;
                    if (ownerView != null)
                    {
                        bb = elem.get_BoundingBox(ownerView);
                    }
                }
                catch { }
            }

            if (bb != null)
            {
                return (bb.Min + bb.Max) / 2.0;
            }

            return null;
        }

        /// <summary>
        /// Move an element by a translation vector
        /// Returns true if successful
        /// CRITICAL: For TextNote, moves the Coord property which will translate the bounding box by the same amount
        /// </summary>
        private static bool MoveElement(Element elem, Document doc, XYZ translation)
        {
            try
            {
                // Viewport elements - use SetBoxCenter to reposition on sheet
                if (elem is Viewport viewport)
                {
                    XYZ currentCenter = GetElementPosition(viewport, doc);
                    if (currentCenter != null)
                    {
                        XYZ newCenter = currentCenter + translation;
                        viewport.SetBoxCenter(newCenter);
                        return true;
                    }
                    return false;
                }

                // LocationPoint elements
                if (elem.Location is LocationPoint locationPoint)
                {
                    XYZ newPoint = locationPoint.Point + translation;
                    locationPoint.Point = newPoint;
                    return true;
                }

                // LocationCurve elements
                if (elem.Location is LocationCurve locationCurve)
                {
                    locationCurve.Move(translation);
                    return true;
                }

                // TextNote elements and other elements without Location - use Coord property or ElementTransformUtils
                if (elem is TextNote textNote)
                {
                    var coordProp = textNote.GetType().GetProperty("Coord");
                    if (coordProp != null && coordProp.CanWrite)
                    {
                        XYZ currentCoord = coordProp.GetValue(textNote, null) as XYZ;
                        if (currentCoord != null)
                        {
                            XYZ newCoord = currentCoord + translation;
                            coordProp.SetValue(textNote, newCoord, null);
                            return true;
                        }
                    }
                }

                // Fallback: Use ElementTransformUtils for other elements
                ElementTransformUtils.MoveElement(doc, elem.Id, translation);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Ensure handlers are initialized (called automatically when registry is first accessed)
        /// </summary>
        public static void EnsureInitialized()
        {
            if (!_initialized)
                RegisterStandardHandlers();
        }
    }
}
