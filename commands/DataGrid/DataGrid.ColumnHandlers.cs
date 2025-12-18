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

        /// <summary>
        /// Register a column handler
        /// </summary>
        public static void Register(ColumnHandler handler)
        {
            _handlers[handler.ColumnName] = handler;
        }

        /// <summary>
        /// Get handler for a column name (case-insensitive)
        /// </summary>
        public static ColumnHandler GetHandler(string columnName)
        {
            _handlers.TryGetValue(columnName, out var handler);
            return handler;
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
