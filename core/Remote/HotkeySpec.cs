using System;
using System.Collections.Generic;

namespace RevitBallet.Core.Remote
{
    /// <summary>A global hotkey as written in remote.json: modifiers joined with '+', then one key. "Win+`", "Alt+Shift+K", "Ctrl+F9".</summary>
    public sealed class HotkeySpec
    {
        // the RegisterHotKey modifier bits
        public const uint Alt = 1, Ctrl = 2, Shift = 4, Win = 8;

        public uint Modifiers;
        /// <summary>The key as written, normalised: a single character ("`", "K") or a name ("F9", "Space").</summary>
        public string Key;

        public bool IsBacktick { get { return Key == "`"; } }

        public static bool TryParse(string text, out HotkeySpec spec)
        {
            spec = null;
            if (string.IsNullOrWhiteSpace(text)) return false;
            // the key is last; a lone '+' key is written "Ctrl++"
            var parts = new List<string>();
            string rest = text.Trim();
            string key;
            if (rest.EndsWith("++")) { key = "+"; rest = rest.Substring(0, rest.Length - 2); }
            else
            {
                int i = rest.LastIndexOf('+');
                key = i < 0 ? rest : rest.Substring(i + 1);
                rest = i < 0 ? "" : rest.Substring(0, i);
            }
            key = key.Trim();
            if (key.Length == 0) return false;
            uint mods = 0;
            foreach (var m in rest.Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries))
            {
                switch (m.Trim().ToLowerInvariant())
                {
                    case "alt": mods |= Alt; break;
                    case "ctrl": case "control": mods |= Ctrl; break;
                    case "shift": mods |= Shift; break;
                    case "win": case "super": case "mod": case "meta": mods |= Win; break;
                    default: return false;
                }
            }
            if (mods == 0) return false;               // a bare key would swallow typing everywhere
            if (key.Length > 1) key = char.ToUpperInvariant(key[0]) + key.Substring(1).ToLowerInvariant();
            else key = key.ToUpperInvariant();
            if (key.Length > 1 && !IsNamedKey(key)) return false;
            spec = new HotkeySpec { Modifiers = mods, Key = key };
            return true;
        }

        private static bool IsNamedKey(string k)
        {
            if (k.Length >= 2 && k.Length <= 3 && k[0] == 'F')
            {
                int n; return int.TryParse(k.Substring(1), out n) && n >= 1 && n <= 24;
            }
            switch (k)
            {
                case "Space": case "Tab": case "Enter": case "Esc": case "Insert": case "Delete": case "Home": case "End": case "Pageup": case "Pagedown":
                case "Up": case "Down": case "Left": case "Right": case "Grave": return true;
            }
            return false;
        }

        public override string ToString()
        {
            var s = new List<string>();
            if ((Modifiers & Win) != 0) s.Add("Win");
            if ((Modifiers & Ctrl) != 0) s.Add("Ctrl");
            if ((Modifiers & Alt) != 0) s.Add("Alt");
            if ((Modifiers & Shift) != 0) s.Add("Shift");
            s.Add(Key);
            return string.Join("+", s);
        }
    }
}
