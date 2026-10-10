using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace RevitBalletInstaller
{
    /// <summary>What to do with keyboard shortcuts. Prompt = ask (interactive); Auto = the classic behaviour (quiet default).</summary>
    public enum ShortcutMode { Prompt, Auto, None, Ballet, All }

    /// <summary>One shortcut offered to the user, with what it would clash with.</summary>
    internal sealed class ShortcutRow
    {
        public string CommandId, Title, Keys;
        public bool IsBallet;
        /// <summary>Takes a key sequence the user (or Revit's defaults) already use for something else, or changes the user's own assignment.</summary>
        public bool Hard;
        public string Note = "";
        public bool Selected = true;
    }

    public partial class BalletInstaller
    {
        public static ShortcutMode ShortcutsMode = ShortcutMode.Prompt;
        public static bool ShortcutsReportOnly;

        /// <summary>null = classic behaviour; otherwise exactly the CommandIds to write.</summary>
        private HashSet<string> chosenShortcutIds;

        private sealed class YearShortcuts
        {
            public string Year;
            public Dictionary<string, string> Merged = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);     // bundled set: id -> keys
            public Dictionary<string, string> Existing = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);   // the user's KeyboardShortcuts.xml today
            public HashSet<string> CustomIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        private Dictionary<string, string> addinNames;   // AddInId -> command class name

        private string ResourceText(string suffix)
        {
            var name = Assembly.GetExecutingAssembly().GetManifestResourceNames().FirstOrDefault(r => r.EndsWith(suffix));
            return name == null ? null : GetResourceText(name);
        }

        private static Dictionary<string, string> ReadItems(XDocument doc)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var e in doc.Descendants("ShortcutItem"))
            {
                var id = (string)e.Attribute("CommandId");
                if (!string.IsNullOrEmpty(id)) d[id] = (string)e.Attribute("Shortcuts") ?? "";
            }
            return d;
        }

        private List<YearShortcuts> LoadShortcutData()
        {
            var list = new List<YearShortcuts>();
            var addin = ResourceText("revit-ballet.addin");
            addinNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (addin != null)
                foreach (var a in XDocument.Parse(addin).Descendants("AddIn"))
                {
                    var id = a.Element("AddInId")?.Value?.Trim(); var cls = a.Element("FullClassName")?.Value?.Trim();
                    if (id != null && cls != null) addinNames[id] = cls.Substring(cls.LastIndexOf('.') + 1);
                }
            var customIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var custom = ResourceText("KeyboardShortcuts-custom.xml");
            if (custom != null)
            {
                var body = string.Join("\n", custom.Split('\n').Where(l => !l.TrimStart().StartsWith("<?xml")));
                if (!body.Contains("<Shortcuts")) body = "<Shortcuts>" + body + "</Shortcuts>";
                foreach (var k in ReadItems(XDocument.Parse(body)).Keys) customIds.Add(k);
            }
            foreach (var inst in installations)
            {
                var merged = ResourceText("KeyboardShortcuts-" + inst.Year + ".xml");
                if (merged == null) continue;
                var ys = new YearShortcuts { Year = inst.Year, Merged = ReadItems(XDocument.Parse(merged)), CustomIds = customIds };
                try
                {
                    var p = Path.Combine(inst.RevitDataPath, "KeyboardShortcuts.xml");
                    if (File.Exists(p)) ys.Existing = ReadItems(XDocument.Load(p));
                }
                catch { }
                list.Add(ys);
            }
            return list;
        }

        private string ShortcutTitle(string id)
        {
            string n;
            return addinNames != null && addinNames.TryGetValue(id, out n) ? n : id;
        }

        private static IEnumerable<string> KeyList(string keys)
        {
            return (keys ?? "").Split('#').Select(k => k.Trim()).Where(k => k.Length > 0);
        }

        /// <summary>exact: the same key sequence; overlap: one is the start of the other (Revit then waits or fires the shorter one).</summary>
        private static string Collide(string a, string b)
        {
            foreach (var x in KeyList(a))
                foreach (var y in KeyList(b))
                {
                    if (string.Equals(x, y, StringComparison.OrdinalIgnoreCase)) return "exact:" + x;
                    if (x.IndexOf('+') < 0 && y.IndexOf('+') < 0 &&
                        (x.StartsWith(y, StringComparison.OrdinalIgnoreCase) || y.StartsWith(x, StringComparison.OrdinalIgnoreCase)))
                        return "overlap:" + x + "/" + y;
                }
            return null;
        }

        /// <summary>The rows offered for a choice (Ballet = Revit Ballet commands, All = also the other bundled shortcuts), with clash notes.</summary>
        private List<ShortcutRow> BuildShortcutRows(List<YearShortcuts> data, ShortcutMode option)
        {
            var rows = new List<ShortcutRow>();
            if (data.Count == 0) return rows;
            var newest = data.OrderByDescending(d => int.Parse(d.Year)).First();
            foreach (var kv in newest.Merged.OrderBy(k => addinNames.ContainsKey(k.Key) ? 0 : 1).ThenBy(k => ShortcutTitle(k.Key), StringComparer.OrdinalIgnoreCase))
            {
                bool ballet = addinNames.ContainsKey(kv.Key);
                if (!ballet && option != ShortcutMode.All) continue;
                rows.Add(new ShortcutRow { CommandId = kv.Key, Title = ShortcutTitle(kv.Key), Keys = kv.Value, IsBallet = ballet });
            }
            foreach (var row in rows)
            {
                var notes = new List<string>(); bool hard = false;
                foreach (var ys in data)
                {
                    string newKeys; if (!ys.Merged.TryGetValue(row.CommandId, out newKeys)) continue;
                    string old;
                    if (ys.Existing.TryGetValue(row.CommandId, out old) && !string.Equals(old, newKeys, StringComparison.OrdinalIgnoreCase))
                    { hard = true; notes.Add($"replaces your {(string.IsNullOrEmpty(old) ? "none" : old)} ({ys.Year})"); }
                    // other commands that already hold these keys: the user's current file, then (Ballet only) the bundled non-Ballet set standing in for Revit's defaults
                    var refs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    if (option != ShortcutMode.All)
                        foreach (var m in ys.Merged) if (!ys.CustomIds.Contains(m.Key)) refs[m.Key] = m.Value;
                    foreach (var u in ys.Existing) refs[u.Key] = u.Value;
                    foreach (var r in refs)
                    {
                        if (string.Equals(r.Key, row.CommandId, StringComparison.OrdinalIgnoreCase)) continue;
                        var c = Collide(newKeys, r.Value);
                        if (c == null) continue;
                        bool yours = ys.Existing.ContainsKey(r.Key);
                        if (c.StartsWith("exact:")) { hard = true; notes.Add($"clash: {c.Substring(6)} is {(yours ? "yours" : "a Revit default")} for {ShortcutTitle(r.Key)}"); }
                        else notes.Add($"overlaps {ShortcutTitle(r.Key)} {r.Value}");
                        if (notes.Count >= 4) break;
                    }
                }
                row.Hard = hard;
                row.Selected = !hard;
                row.Note = string.Join("; ", notes.Distinct().Take(3));
            }
            return rows;
        }

        private void PrintShortcutReport(List<ShortcutRow> rows)
        {
            Console.WriteLine($"shortcuts: {rows.Count} offered, {rows.Count(r => r.Hard)} clash with defaults or your own, {rows.Count(r => !r.Hard && r.Note.Length > 0)} overlap");
            foreach (var r in rows.Where(r => r.Note.Length > 0))
                Console.WriteLine($"  {(r.Hard ? "CLASH  " : "overlap")} {r.Title,-34} {r.Keys,-14} {r.Note}");
        }

        /// <summary>Decide what to do with shortcuts for this run (dialog, flag or classic). Returns false when the user cancelled the whole install.</summary>
        private bool PrepareShortcutChoice(bool quietMode)
        {
            var mode = ShortcutsMode;
            if (mode == ShortcutMode.Prompt && quietMode) mode = ShortcutMode.Auto;
            if (ShortcutsReportOnly)
            {
                var d = LoadShortcutData();
                Console.WriteLine("Revit Ballet shortcuts only:");
                PrintShortcutReport(BuildShortcutRows(d, ShortcutMode.Ballet));
                Console.WriteLine("Revit Ballet and the other bundled shortcuts:");
                PrintShortcutReport(BuildShortcutRows(d, ShortcutMode.All));
                return false;
            }
            if (mode == ShortcutMode.Auto) { chosenShortcutIds = null; return true; }
            if (mode == ShortcutMode.None) { chosenShortcutIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase); return true; }
            var data = LoadShortcutData();
            if (mode == ShortcutMode.Ballet || mode == ShortcutMode.All)
            {
                var rows = BuildShortcutRows(data, mode);
                if (_verbosity >= 1 && rows.Any(r => r.Hard)) PrintShortcutReport(rows.Where(r => r.Note.Length > 0).ToList());
                chosenShortcutIds = new HashSet<string>(rows.Select(r => r.CommandId), StringComparer.OrdinalIgnoreCase);
                return true;
            }
            // Prompt
            var choice = AskShortcuts(data);
            if (choice == null) return false;
            chosenShortcutIds = choice;
            return true;
        }

        /// <summary>Write the chosen shortcuts into each Revit's KeyboardShortcuts.xml, leaving every other entry as the user has it.</summary>
        private void ApplyChosenShortcuts(RevitInstallation installation, bool quietMode)
        {
            if (chosenShortcutIds == null || chosenShortcutIds.Count == 0) return;
            try
            {
                var merged = ResourceText("KeyboardShortcuts-" + installation.Year + ".xml");
                if (merged == null) return;
                var mergedDoc = XDocument.Parse(merged);
                string kbPath = Path.Combine(installation.RevitDataPath, "KeyboardShortcuts.xml");
                XDocument doc = File.Exists(kbPath) ? XDocument.Load(kbPath) : new XDocument(new XElement("Shortcuts"));
                var root = doc.Element("Shortcuts");
                if (root == null) return;
                foreach (var item in mergedDoc.Descendants("ShortcutItem"))
                {
                    var id = (string)item.Attribute("CommandId");
                    if (id == null || !chosenShortcutIds.Contains(id)) continue;
                    var existing = root.Descendants("ShortcutItem").FirstOrDefault(e => string.Equals((string)e.Attribute("CommandId"), id, StringComparison.OrdinalIgnoreCase));
                    if (existing != null) foreach (var a in item.Attributes()) existing.SetAttributeValue(a.Name, a.Value);
                    else root.Add(new XElement(item));
                }
                Directory.CreateDirectory(installation.RevitDataPath);
                doc.Save(kbPath);
            }
            catch (Exception ex)
            {
                if (!quietMode) MessageBox.Show($"Failed to deploy keyboard shortcuts for Revit {installation.Year}: {ex.Message}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else Console.Error.WriteLine($"Warning: failed to deploy keyboard shortcuts for Revit {installation.Year}: {ex.Message}");
            }
        }

        // ------------------------------------------------------------------ dialog

        private HashSet<string> AskShortcuts(List<YearShortcuts> data)
        {
            if (data.Count == 0) return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ballet = BuildShortcutRows(data, ShortcutMode.Ballet);
            var all = BuildShortcutRows(data, ShortcutMode.All);
            var font = new Font("Segoe UI", 9.5f);
            var form = new Form
            {
                Text = "Revit Ballet - Keyboard shortcuts", StartPosition = FormStartPosition.CenterScreen, Font = font, BackColor = Color.White,
                ClientSize = new Size(900, 640), MinimumSize = new Size(760, 520), ShowIcon = false
            };
            var intro = new Label { Left = 16, Top = 12, Width = 868, Height = 44, Text = "Install keyboard shortcuts? Your existing shortcuts are left as they are unless you tick a row that replaces one. Rows in red take keys that you or Revit already use for something else; they start unticked." };
            var rbNone = new RadioButton { Left = 20, Top = 62, Width = 860, Height = 24, Text = "Don't touch my keyboard shortcuts" };
            var rbBallet = new RadioButton { Left = 20, Top = 88, Width = 860, Height = 24, Text = $"Revit Ballet shortcuts only ({ballet.Count} commands)", Checked = true };
            var rbAll = new RadioButton { Left = 20, Top = 114, Width = 860, Height = 24, Text = $"Revit Ballet shortcuts and the other shortcuts bundled with it ({all.Count} in total)" };
            var summary = new Label { Left = 16, Top = 144, Width = 868, Height = 22, ForeColor = Color.FromArgb(90, 90, 90) };
            var list = new ListView { Left = 16, Top = 170, Width = 868, Height = 400, View = View.Details, CheckBoxes = true, FullRowSelect = true, HideSelection = false, GridLines = false, Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };
            list.Columns.Add("Command", 250); list.Columns.Add("Keys", 110); list.Columns.Add("Note", 480);
            var bAll = new Button { Left = 16, Top = 580, Width = 110, Height = 30, Text = "Tick all", Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            var bNone = new Button { Left = 132, Top = 580, Width = 110, Height = 30, Text = "Untick all", Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            var bSafe = new Button { Left = 248, Top = 580, Width = 190, Height = 30, Text = "Tick only without clashes", Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            var ok = new Button { Left = 784, Top = 596, Width = 100, Height = 30, Text = "Install", DialogResult = DialogResult.OK, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            var cancel = new Button { Left = 676, Top = 596, Width = 100, Height = 30, Text = "Cancel", DialogResult = DialogResult.Cancel, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            form.Controls.AddRange(new Control[] { intro, rbNone, rbBallet, rbAll, summary, list, bAll, bNone, bSafe, ok, cancel });
            form.AcceptButton = ok; form.CancelButton = cancel;
            ok.Top = 580; cancel.Top = 580;

            List<ShortcutRow> current = ballet;
            bool filling = false;
            Action fill = () =>
            {
                filling = true; list.BeginUpdate(); list.Items.Clear();
                foreach (var r in current)
                {
                    var it = new ListViewItem(new[] { r.Title, r.Keys, r.Note }) { Checked = r.Selected, Tag = r };
                    if (r.Hard) it.ForeColor = Color.FromArgb(180, 30, 30);
                    else if (r.Note.Length > 0) it.ForeColor = Color.FromArgb(150, 100, 0);
                    list.Items.Add(it);
                }
                list.EndUpdate(); filling = false;
            };
            Action refresh = () =>
            {
                int hard = current.Count(r => r.Hard), soft = current.Count(r => !r.Hard && r.Note.Length > 0);
                summary.Text = rbNone.Checked ? "No shortcuts will be changed."
                    : $"{list.CheckedItems.Count} of {current.Count} ticked.  {hard} clash with your own or Revit's default shortcuts (red), {soft} overlap another shortcut (amber).";
                list.Enabled = !rbNone.Checked; bAll.Enabled = bNone.Enabled = bSafe.Enabled = !rbNone.Checked;
            };
            list.ItemChecked += (s, e) => { if (filling) return; ((ShortcutRow)e.Item.Tag).Selected = e.Item.Checked; refresh(); };
            EventHandler pick = (s, e) => { if (!((RadioButton)s).Checked) return; if (rbBallet.Checked) current = ballet; else if (rbAll.Checked) current = all; fill(); refresh(); };
            rbNone.CheckedChanged += pick; rbBallet.CheckedChanged += pick; rbAll.CheckedChanged += pick;
            bAll.Click += (s, e) => { foreach (ListViewItem it in list.Items) it.Checked = true; };
            bNone.Click += (s, e) => { foreach (ListViewItem it in list.Items) it.Checked = false; };
            bSafe.Click += (s, e) => { foreach (ListViewItem it in list.Items) it.Checked = !((ShortcutRow)it.Tag).Hard; };
            fill(); refresh();
            if (form.ShowDialog() != DialogResult.OK) return null;
            if (rbNone.Checked) return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            return new HashSet<string>(current.Where(r => r.Selected).Select(r => r.CommandId), StringComparer.OrdinalIgnoreCase);
        }
    }
}
