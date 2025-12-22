using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;

namespace RevitBallet.Commands
{
    /// <summary>
    /// Caches view grid data for OpenViewsInDocument and OpenViewsInSession commands.
    /// Uses document fingerprints to detect changes and invalidate cache when needed.
    /// </summary>
    public static class ViewDataCache
    {
        private class CachedViewData
        {
            public string Fingerprint { get; set; }
            public List<Dictionary<string, object>> GridData { get; set; }
            public List<string> Columns { get; set; }
        }

        private class SessionCacheData
        {
            public string SessionFingerprint { get; set; }
            public List<Dictionary<string, object>> GridData { get; set; }
            public List<string> Columns { get; set; }
        }

        // Cache storage: key is document PathName (or Title for unsaved docs)
        private static Dictionary<string, CachedViewData> documentCache =
            new Dictionary<string, CachedViewData>();

        // Session cache (for OpenViewsInSession)
        private static SessionCacheData sessionCache = null;

        /// <summary>
        /// Generates a fingerprint for a document's view state.
        /// Changes when views are added/deleted/renamed or viewports change.
        /// </summary>
        private static string GetDocumentFingerprint(Document doc)
        {
            // Collect view count and IDs
            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate &&
                           v.ViewType != ViewType.ProjectBrowser &&
                           v.ViewType != ViewType.SystemBrowser)
                .ToList();

            // Collect viewport count and IDs
            var viewports = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .ToList();

            // Build fingerprint from counts and max IDs
            var sb = new StringBuilder();
            sb.Append($"Views:{views.Count}|");
            sb.Append($"Viewports:{viewports.Count}|");

            // Include max view ID as a proxy for changes
            if (views.Any())
            {
#if REVIT2024 || REVIT2025 || REVIT2026
                long maxViewId = views.Max(v => v.Id.Value);
#else
                long maxViewId = views.Max(v => v.Id.IntegerValue);
#endif
                sb.Append($"MaxViewId:{maxViewId}|");
            }

            // Include max viewport ID as a proxy for changes
            if (viewports.Any())
            {
#if REVIT2024 || REVIT2025 || REVIT2026
                long maxViewportId = viewports.Max(vp => vp.Id.Value);
#else
                long maxViewportId = viewports.Max(vp => vp.Id.IntegerValue);
#endif
                sb.Append($"MaxViewportId:{maxViewportId}|");
            }

            // Include view names hash (detects renames)
            if (views.Any())
            {
#if REVIT2024 || REVIT2025 || REVIT2026
                string viewNamesHash = GetStringHash(
                    string.Join("|", views.OrderBy(v => v.Id.Value).Select(v => v.Name))
                );
#else
                string viewNamesHash = GetStringHash(
                    string.Join("|", views.OrderBy(v => v.Id.IntegerValue).Select(v => v.Name))
                );
#endif
                sb.Append($"ViewNames:{viewNamesHash}|");
            }

            return sb.ToString();
        }

        /// <summary>
        /// Generates a session fingerprint for all open documents.
        /// </summary>
        private static string GetSessionFingerprint(Application app)
        {
            var docFingerprints = new List<string>();

            foreach (Document doc in app.Documents)
            {
                if (doc.IsLinked || doc.IsFamilyDocument)
                    continue;

                string docKey = GetDocumentKey(doc);
                string fingerprint = GetDocumentFingerprint(doc);
                docFingerprints.Add($"{docKey}:{fingerprint}");
            }

            return string.Join("||", docFingerprints.OrderBy(s => s));
        }

        /// <summary>
        /// Gets a unique key for a document (PathName or Title).
        /// </summary>
        private static string GetDocumentKey(Document doc)
        {
            return string.IsNullOrEmpty(doc.PathName)
                ? $"UNSAVED:{doc.Title}"
                : doc.PathName;
        }

        /// <summary>
        /// Computes MD5 hash of a string (for view names, etc.).
        /// </summary>
        private static string GetStringHash(string input)
        {
            using (MD5 md5 = MD5.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(input);
                byte[] hashBytes = md5.ComputeHash(inputBytes);
                return BitConverter.ToString(hashBytes).Replace("-", "").Substring(0, 16);
            }
        }

        /// <summary>
        /// Tries to get cached grid data for a document.
        /// Returns true if valid cache exists, false if cache needs to be rebuilt.
        /// </summary>
        public static bool TryGetDocumentCache(
            Document doc,
            out List<Dictionary<string, object>> gridData,
            out List<string> columns)
        {
            gridData = null;
            columns = null;

            string docKey = GetDocumentKey(doc);
            string currentFingerprint = GetDocumentFingerprint(doc);

            // Check if we have cached data
            if (documentCache.TryGetValue(docKey, out var cached))
            {
                // Validate fingerprint only - no time-based expiration
                if (cached.Fingerprint == currentFingerprint)
                {
                    gridData = cached.GridData;
                    columns = cached.Columns;
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Saves grid data to document cache.
        /// </summary>
        public static void SaveDocumentCache(
            Document doc,
            List<Dictionary<string, object>> gridData,
            List<string> columns)
        {
            string docKey = GetDocumentKey(doc);
            string fingerprint = GetDocumentFingerprint(doc);

            documentCache[docKey] = new CachedViewData
            {
                Fingerprint = fingerprint,
                GridData = gridData,
                Columns = columns
            };
        }

        /// <summary>
        /// Tries to get cached grid data for session (all open documents).
        /// Returns true if valid cache exists, false if cache needs to be rebuilt.
        /// </summary>
        public static bool TryGetSessionCache(
            Application app,
            out List<Dictionary<string, object>> gridData,
            out List<string> columns)
        {
            gridData = null;
            columns = null;

            string currentFingerprint = GetSessionFingerprint(app);

            if (sessionCache != null)
            {
                // Validate fingerprint only - no time-based expiration
                if (sessionCache.SessionFingerprint == currentFingerprint)
                {
                    gridData = sessionCache.GridData;
                    columns = sessionCache.Columns;
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Saves grid data to session cache.
        /// </summary>
        public static void SaveSessionCache(
            Application app,
            List<Dictionary<string, object>> gridData,
            List<string> columns)
        {
            string fingerprint = GetSessionFingerprint(app);

            sessionCache = new SessionCacheData
            {
                SessionFingerprint = fingerprint,
                GridData = gridData,
                Columns = columns
            };
        }

        /// <summary>
        /// Invalidates cache for a specific document.
        /// Useful when you know a document has changed.
        /// </summary>
        public static void InvalidateDocument(Document doc)
        {
            string docKey = GetDocumentKey(doc);
            documentCache.Remove(docKey);

            // Also invalidate session cache since it includes this document
            sessionCache = null;
        }

        /// <summary>
        /// Invalidates all caches.
        /// </summary>
        public static void InvalidateAll()
        {
            documentCache.Clear();
            sessionCache = null;
        }

        /// <summary>
        /// Gets cache statistics for diagnostics.
        /// </summary>
        public static string GetCacheStats()
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Document Cache Entries: {documentCache.Count}");
            sb.AppendLine($"Session Cache: {(sessionCache != null ? "Active" : "Empty")}");

            if (documentCache.Any())
            {
                sb.AppendLine("\nDocument Cache Details:");
                foreach (var kvp in documentCache.OrderBy(k => k.Key))
                {
                    sb.AppendLine($"  {kvp.Key}: {kvp.Value.GridData.Count} rows (fingerprint: {kvp.Value.Fingerprint.Substring(0, 40)}...)");
                }
            }

            return sb.ToString();
        }
    }
}
