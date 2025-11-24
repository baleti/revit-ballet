using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

/// <summary>
/// Extension methods to provide API compatibility across different Revit versions.
/// Revit 2024+ changed ElementId from int-based to long-based, introducing breaking changes.
/// This compatibility layer ensures code works across Revit 2017-2026+.
/// </summary>
public static class RevitApiCompatibility
{
#if REVIT2011 || REVIT2012 || REVIT2013 || REVIT2014 || REVIT2015 || REVIT2016 || REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021 || REVIT2022 || REVIT2023
    /// <summary>
    /// Gets the numeric value of an ElementId as a long.
    /// In Revit 2024+, use ElementId.Value directly.
    /// In earlier versions (2017-2023), this returns IntegerValue cast to long.
    /// </summary>
    public static long AsLong(this ElementId elementId)
    {
        return elementId.IntegerValue;
    }

    /// <summary>
    /// Creates an ElementId from a long value in Revit versions prior to 2024.
    /// In Revit 2024+, the ElementId constructor accepts long directly.
    /// </summary>
    public static ElementId ToElementId(this long value)
    {
        return new ElementId((int)value);
    }

    /// <summary>
    /// Creates an ElementId from an int value (convenience method for compatibility).
    /// </summary>
    public static ElementId ToElementId(this int value)
    {
        return new ElementId(value);
    }

    /// <summary>
    /// Converts BuiltInParameter enum to long for compatibility.
    /// In Revit 2024+, BuiltInParameter values are long-based.
    /// In earlier versions, they are int-based.
    /// </summary>
    public static long ToLong(this BuiltInParameter parameter)
    {
        return (int)parameter;
    }

    /// <summary>
    /// Creates ElementId from BuiltInParameter for version compatibility.
    /// </summary>
    public static ElementId ToElementId(this BuiltInParameter parameter)
    {
        return new ElementId((int)parameter);
    }
#else
    /// <summary>
    /// Gets the numeric value of an ElementId as a long.
    /// In Revit 2024+, this is equivalent to ElementId.Value.
    /// </summary>
    public static long AsLong(this ElementId elementId)
    {
        return elementId.Value;
    }

    /// <summary>
    /// Creates an ElementId from a long value in Revit 2024+.
    /// This is a convenience extension to match the API in earlier versions.
    /// </summary>
    public static ElementId ToElementId(this long value)
    {
        return new ElementId(value);
    }

    /// <summary>
    /// Creates an ElementId from an int value (convenience method for compatibility).
    /// In Revit 2024+, this casts to long first.
    /// </summary>
    public static ElementId ToElementId(this int value)
    {
        return new ElementId((long)value);
    }

    /// <summary>
    /// Converts BuiltInParameter enum to long (no-op in Revit 2024+).
    /// </summary>
    public static long ToLong(this BuiltInParameter parameter)
    {
        return (long)parameter;
    }

    /// <summary>
    /// Creates ElementId from BuiltInParameter for version compatibility.
    /// </summary>
    public static ElementId ToElementId(this BuiltInParameter parameter)
    {
        return new ElementId((long)parameter);
    }
#endif

#if REVIT2011 || REVIT2012 || REVIT2013 || REVIT2014 || REVIT2015 || REVIT2016 || REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021 || REVIT2022
    /// <summary>
    /// Gets projection to sheet transform for a viewport (Revit 2022 and earlier).
    /// In Revit 2023+, Viewport.GetProjectionToSheetTransform() is built-in.
    /// </summary>
    public static Transform GetProjectionToSheetTransformCompat(this Viewport viewport)
    {
        // For pre-2023, we need to manually calculate this
        // Using the old method via GetBoxCenter and other properties
        var doc = viewport.Document;
        var view = doc.GetElement(viewport.ViewId) as View;
        if (view == null) return Transform.Identity;

        // Get the viewport's position on the sheet
        XYZ viewportCenter = viewport.GetBoxCenter();

        // Get view's crop box and scale
        double scale = viewport.Parameters.Cast<Parameter>()
            .FirstOrDefault(p => p.Definition.Name == "View Scale")?.AsInteger() ?? 1;

        if (scale == 0) scale = 1;
        double scaleFactor = 1.0 / scale;

        // Create transform
        Transform transform = Transform.CreateTranslation(viewportCenter);
        transform = transform.ScaleBasis(scaleFactor);

        return transform;
    }

    /// <summary>
    /// Helper class to provide compatibility for ModelToProjectionTransforms in pre-2023 versions.
    /// </summary>
    public class ModelToProjectionTransformAdapterCompat
    {
        private readonly Transform _transform;

        public ModelToProjectionTransformAdapterCompat(Transform transform)
        {
            _transform = transform;
        }

        public Transform GetModelToProjectionTransform()
        {
            return _transform;
        }
    }

    /// <summary>
    /// Gets model to projection transforms for a view (Revit 2022 and earlier).
    /// In Revit 2023+, View.GetModelToProjectionTransforms() is built-in.
    /// </summary>
    public static IList<ModelToProjectionTransformAdapterCompat> GetModelToProjectionTransforms(this View view)
    {
        // For pre-2023, create a simple transform based on the view's crop box
        // This is a simplified version - the actual transform calculation is complex
        Transform transform = Transform.Identity;

        try
        {
            // Try to get the view's crop box to build a transform
            if (view.CropBoxActive && view.CropBox != null)
            {
                BoundingBoxXYZ cropBox = view.CropBox;
                transform = cropBox.Transform;
            }
        }
        catch
        {
            // If we can't get crop box, use identity transform
        }

        return new List<ModelToProjectionTransformAdapterCompat>
        {
            new ModelToProjectionTransformAdapterCompat(transform)
        };
    }

    /// <summary>
    /// Gets references from Selection (Revit 2022 and earlier).
    /// In Revit 2023+, Selection.GetReferences() is built-in.
    /// Pre-2023 versions don't support Reference-based selection properly.
    /// This is a compatibility stub that returns an empty list.
    /// </summary>
    public static IList<Reference> GetReferences(this Autodesk.Revit.UI.Selection.Selection selection)
    {
        // Pre-2023 versions don't have native support for Reference-based selection
        // Return empty list as a compatibility stub
        return new List<Reference>();
    }

    /// <summary>
    /// Sets references in Selection (Revit 2022 and earlier).
    /// In Revit 2023+, Selection.SetReferences() is built-in.
    /// Pre-2023 versions don't support Reference-based selection properly.
    /// This is a compatibility stub that extracts ElementIds from references.
    /// </summary>
    public static void SetReferences(this Autodesk.Revit.UI.Selection.Selection selection, IList<Reference> references)
    {
        // Pre-2023 versions don't have native support for Reference-based selection
        // Extract ElementIds from references and set those instead
        var elementIds = new List<ElementId>();
        foreach (var reference in references)
        {
            if (reference != null && reference.ElementId != ElementId.InvalidElementId)
            {
                elementIds.Add(reference.ElementId);
            }
        }
        selection.SetElementIds(elementIds);
    }
#else
    /// <summary>
    /// Gets projection to sheet transform for a viewport (Revit 2023+).
    /// Wrapper for built-in Viewport.GetProjectionToSheetTransform().
    /// </summary>
    public static Transform GetProjectionToSheetTransformCompat(this Viewport viewport)
    {
        return viewport.GetProjectionToSheetTransform();
    }
#endif
}
