using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

/// <summary>
/// Validation infrastructure for DataGrid column editing
/// </summary>
public partial class CustomGUIs
{
    /// <summary>
    /// Result of a validation operation
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public string ErrorMessage { get; set; }

        public static ValidationResult Valid() => new ValidationResult { IsValid = true };

        public static ValidationResult Invalid(string errorMessage) =>
            new ValidationResult { IsValid = false, ErrorMessage = errorMessage };
    }

    /// <summary>
    /// Built-in validators for common validation scenarios
    /// </summary>
    public static class ColumnValidators
    {
        /// <summary>
        /// Validate that value is not empty
        /// </summary>
        public static ValidationResult NotEmpty(Element elem, Document doc, object oldValue, object newValue)
        {
            if (string.IsNullOrWhiteSpace(newValue?.ToString()))
                return ValidationResult.Invalid("Value cannot be empty");
            return ValidationResult.Valid();
        }

        /// <summary>
        /// Validate that value doesn't contain invalid filename characters
        /// </summary>
        public static ValidationResult NoInvalidCharacters(Element elem, Document doc, object oldValue, object newValue)
        {
            string strValue = newValue?.ToString() ?? "";
            char[] invalid = new[] { '/', '\\', ':', '*', '?', '"', '<', '>', '|' };

            if (strValue.IndexOfAny(invalid) >= 0)
                return ValidationResult.Invalid($"Value cannot contain: {string.Join(" ", invalid)}");

            return ValidationResult.Valid();
        }

        /// <summary>
        /// Validate that value doesn't start or end with whitespace
        /// </summary>
        public static ValidationResult NoLeadingTrailingWhitespace(Element elem, Document doc, object oldValue, object newValue)
        {
            string strValue = newValue?.ToString() ?? "";
            if (strValue != strValue.Trim())
                return ValidationResult.Invalid("Value cannot start or end with spaces");
            return ValidationResult.Valid();
        }

        /// <summary>
        /// Validate numeric range
        /// </summary>
        public static Func<Element, Document, object, object, ValidationResult> InRange(double min, double max)
        {
            return (elem, doc, oldValue, newValue) =>
            {
                if (double.TryParse(newValue?.ToString(), out double value))
                {
                    if (value < min || value > max)
                        return ValidationResult.Invalid($"Value must be between {min} and {max}");
                    return ValidationResult.Valid();
                }
                return ValidationResult.Invalid("Value must be a number");
            };
        }

        /// <summary>
        /// Validate maximum length
        /// </summary>
        public static Func<Element, Document, object, object, ValidationResult> MaxLength(int maxLength)
        {
            return (elem, doc, oldValue, newValue) =>
            {
                string strValue = newValue?.ToString() ?? "";
                if (strValue.Length > maxLength)
                    return ValidationResult.Invalid($"Value cannot exceed {maxLength} characters");
                return ValidationResult.Valid();
            };
        }

        /// <summary>
        /// Combine multiple validators - ALL must pass
        /// </summary>
        public static Func<Element, Document, object, object, ValidationResult> All(
            params Func<Element, Document, object, object, ValidationResult>[] validators)
        {
            return (elem, doc, oldValue, newValue) =>
            {
                foreach (var validator in validators)
                {
                    var result = validator(elem, doc, oldValue, newValue);
                    if (!result.IsValid)
                        return result;
                }
                return ValidationResult.Valid();
            };
        }
    }
}
