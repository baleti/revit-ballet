using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using RevitBallet.Core.Query;
using RevitBallet.Core.Remote;

namespace RevitBalletRemote
{
    /// <summary>
    /// The command picker: one row per (command x Revit instance), the same query language as the DataGrid in Revit and the Linux pickers
    /// (RevitBallet.Core.Query). Enter runs the selected command (or the first one) in that instance; Esc closes.
    /// </summary>
    internal sealed class PickerForm : Form
    {
        private readonly RemoteService service;
        private readonly TextBox query = new TextBox();
        private readonly DataGridView grid = new DataGridView();
        private readonly Label status = new Label();
        private readonly QuerySchema schema = RemoteRows.Schema();

        private List<RemoteRow> rows = new List<RemoteRow>();
        private List<RemoteRow> shown = new List<RemoteRow>();
        private string refreshNote = "";
        private string selectedId;
        private string columnState = "";
        private int refreshGeneration;

        public event Action<RemoteRow> Invoked;

        public PickerForm(RemoteService service)
        {
            this.service = service;
            Text = "Revit commands";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            ShowInTaskbar = false;
            KeyPreview = false;
            var screen = Screen.FromPoint(Cursor.Position).WorkingArea;
            Size = new Size((int)(screen.Width * 0.6), (int)(screen.Height * 0.55));
            BackColor = Color.FromArgb(30, 30, 46);
            ForeColor = Color.FromArgb(205, 214, 244);
            Font = new Font("Segoe UI", 10f);

            query.Dock = DockStyle.Top; query.Font = new Font("Segoe UI", 13f); query.BorderStyle = BorderStyle.FixedSingle;
            query.BackColor = Color.FromArgb(24, 24, 37); query.ForeColor = ForeColor; query.Name = "query"; query.AccessibleName = "query"; query.AccessibleDescription = "Search commands";
            query.TextChanged += (s, e) => UpdateView();
            query.KeyDown += QueryKeyDown;

            status.Dock = DockStyle.Bottom; status.Height = 24; status.Padding = new Padding(8, 4, 0, 0);
            status.ForeColor = Color.FromArgb(166, 173, 200); status.Name = "status";

            grid.Dock = DockStyle.Fill; grid.Name = "grid"; grid.AccessibleName = "commands";
            grid.VirtualMode = true; grid.ReadOnly = true; grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.AllowUserToResizeRows = false;
            grid.RowHeadersVisible = false; grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = false;
            grid.BackgroundColor = BackColor; grid.GridColor = Color.FromArgb(49, 50, 68); grid.BorderStyle = BorderStyle.None;
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(24, 24, 37); grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(166, 173, 200);
            grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(24, 24, 37);
            grid.DefaultCellStyle.BackColor = BackColor; grid.DefaultCellStyle.ForeColor = ForeColor;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(69, 71, 90); grid.DefaultCellStyle.SelectionForeColor = Color.White;
            grid.CellValueNeeded += (s, e) => { if (e.RowIndex >= 0 && e.RowIndex < shown.Count) e.Value = CellText(shown[e.RowIndex], grid.Columns[e.ColumnIndex].Name); };
            grid.SelectionChanged += (s, e) => { var r = Selected(); if (r != null) selectedId = r.Id; };
            grid.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) InvokeSelected(); };

            Controls.Add(grid); Controls.Add(status); Controls.Add(query);     // dock order: query on top, status at the bottom, grid fills
            grid.BringToFront();
            Deactivate += (s, e) => { /* stays open: the user may click the status or switch back */ };
            Shown += (s, e) => query.Focus();
        }

        private static string CellText(RemoteRow row, string column)
        {
            if (column == "command") return row.Preview;
            object v;
            return row.Fields.TryGetValue(column, out v) ? (string)v : "";
        }

        /// <summary>Opens (or re-opens) the picker: instant from the cache, then refreshed from the targets in the background.</summary>
        public void Open()
        {
            query.Text = "";
            selectedId = null;
            var cached = service.LoadCache();
            rows = cached.Rows;
            refreshNote = cached.Rows.Count > 0 ? "refreshing..." : "looking for Revit instances...";
            UpdateView();
            TopMost = true;
            Show();
            WindowState = FormWindowState.Normal;
            Activate(); BringToFront();
            Native.SetForegroundWindow(Handle);
            query.Focus();
            TopMost = false;

            int generation = ++refreshGeneration;
            Task.Run(() => service.Refresh()).ContinueWith(t =>
            {
                if (IsDisposed || generation != refreshGeneration) return;
                BeginInvoke((Action)(() =>
                {
                    if (generation != refreshGeneration) return;
                    if (t.Exception == null)
                    {
                        rows = t.Result.Rows;
                        refreshNote = t.Result.Errors.Count == 0 ? "" : t.Result.Errors.Count + " target(s) unreachable: " + t.Result.Errors[0];
                    }
                    else refreshNote = "refresh failed: " + t.Exception.GetBaseException().Message;
                    UpdateView();
                }));
            });
        }

        // ------------------------------------------------------------------ query -> grid

        private void UpdateView()
        {
            var fields = rows.Select(r => (IDictionary<string, object>)r.Fields).ToList();
            QueryResult result = QueryEngine.Run(fields, schema, query.Text);
            shown = result.RowIndices.Select(i => rows[i]).ToList();
            ApplyColumns(result.Columns);
            grid.RowCount = 0;
            grid.RowCount = shown.Count;
            int keep = selectedId == null ? -1 : shown.FindIndex(r => r.Id == selectedId);
            if (keep >= 0) Select(keep);
            else { selectedId = null; if (grid.RowCount > 0) grid.ClearSelection(); }
            status.Text = shown.Count + " of " + rows.Count + " commands" + (refreshNote.Length > 0 ? "   " + refreshNote : "");
        }

        /// <summary>The command's own name is always the first, flexible column; the query's columns follow.</summary>
        private void ApplyColumns(List<string> columns)
        {
            var names = new List<string> { "command" };
            names.AddRange(columns.Where(c => c != "command"));
            string state = string.Join("|", names);
            if (state == columnState) return;
            columnState = state;
            grid.Columns.Clear();
            foreach (string n in names)
            {
                var col = new DataGridViewTextBoxColumn { Name = n, HeaderText = n, SortMode = DataGridViewColumnSortMode.NotSortable };
                if (n == "command") { col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; col.FillWeight = 100; }
                else { col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None; col.Width = ColumnWidth(n); }
                grid.Columns.Add(col);
            }
        }

        private static int ColumnWidth(string name)
        {
            switch (name)
            {
                case "input": return 90; case "output": return 100; case "year": return 60; case "pid": return 70; case "state": return 170;
                case "target": return 160; case "instance": return 260; case "class": return 190; default: return 120;
            }
        }

        // ------------------------------------------------------------------ keys and selection

        private RemoteRow Selected()
        {
            if (grid.SelectedRows.Count == 0) return null;
            int i = grid.SelectedRows[0].Index;
            return i >= 0 && i < shown.Count ? shown[i] : null;
        }

        private void Select(int index)
        {
            if (index < 0 || index >= shown.Count) return;
            grid.ClearSelection();
            grid.Rows[index].Selected = true;
            grid.CurrentCell = grid.Rows[index].Cells[0];
            selectedId = shown[index].Id;
        }

        private void MoveSelection(int step)
        {
            if (shown.Count == 0) return;
            int cur = Selected() == null ? -1 : grid.SelectedRows[0].Index;
            Select(cur < 0 ? 0 : Math.Max(0, Math.Min(shown.Count - 1, cur + step)));
        }

        private void QueryKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Escape: Hide(); e.SuppressKeyPress = true; return;
                case Keys.Enter: InvokeSelected(); e.SuppressKeyPress = true; return;
                case Keys.Down: MoveSelection(1); e.SuppressKeyPress = true; return;
                case Keys.Up: MoveSelection(-1); e.SuppressKeyPress = true; return;
                case Keys.PageDown: MoveSelection(Math.Max(1, grid.DisplayedRowCount(false) - 1)); e.SuppressKeyPress = true; return;
                case Keys.PageUp: MoveSelection(-Math.Max(1, grid.DisplayedRowCount(false) - 1)); e.SuppressKeyPress = true; return;
                case Keys.J: if (e.Control) { MoveSelection(1); e.SuppressKeyPress = true; } return;
                case Keys.K: if (e.Control) { MoveSelection(-1); e.SuppressKeyPress = true; } return;
                case Keys.Home: if (e.Control) { Select(0); e.SuppressKeyPress = true; } return;
                case Keys.End: if (e.Control) { Select(shown.Count - 1); e.SuppressKeyPress = true; } return;
            }
        }

        private void InvokeSelected()
        {
            var row = Selected() ?? (shown.Count > 0 ? shown[0] : null);
            if (row == null) return;
            Hide();
            var h = Invoked; if (h != null) h(row);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); return; }      // keep the form; the tray app owns its lifetime
            base.OnFormClosing(e);
        }

        // ------------------------------------------------------------------ for tests

        /// <summary>Rows currently shown and the columns, as plain text (the VM test reads this through UI Automation-free reflection).</summary>
        public string Dump()
        {
            return string.Join("\n", shown.Select(r => r.Id)) + "\n--\n" + columnState;
        }
    }
}
