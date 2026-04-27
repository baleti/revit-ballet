using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace RevitBallet.Commands
{
    internal static class TomlHelper
    {
        public static Dictionary<string, string> Read(string filePath)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!File.Exists(filePath)) return result;

            foreach (var line in File.ReadAllLines(filePath, Encoding.UTF8))
            {
                var trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("#")) continue;

                int eq = trimmed.IndexOf('=');
                if (eq < 0) continue;

                string key = trimmed.Substring(0, eq).Trim();
                string val = trimmed.Substring(eq + 1).Trim();

                if (val.StartsWith("\"") && val.EndsWith("\"") && val.Length >= 2)
                    val = val.Substring(1, val.Length - 2)
                             .Replace("\\\\", "\x00BACKSLASH\x00")
                             .Replace("\\\"", "\"")
                             .Replace("\x00BACKSLASH\x00", "\\");

                result[key] = val;
            }
            return result;
        }

        public static void Write(string filePath, Dictionary<string, string> values)
        {
            string dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);

            var sb = new StringBuilder();
            foreach (var kv in values)
            {
                string val = kv.Value ?? "";
                bool isBool = val == "true" || val == "false";
                bool isInt = int.TryParse(val, out _);
                if (!isBool && !isInt)
                    val = "\"" + val.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
                sb.AppendLine($"{kv.Key} = {val}");
            }
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public static string GetString(Dictionary<string, string> d, string key, string def = "")
            => d.TryGetValue(key, out var v) ? v : def;

        public static bool GetBool(Dictionary<string, string> d, string key, bool def = false)
            => d.TryGetValue(key, out var v) ? v == "true" : def;

        public static int GetInt(Dictionary<string, string> d, string key, int def = 0)
            => d.TryGetValue(key, out var v) && int.TryParse(v, out int r) ? r : def;
    }
}
