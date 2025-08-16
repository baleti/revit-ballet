// File: CopySelectedElementsAlongContainingGroupsByRooms_ProgressForm.cs

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text;

// Aliases to resolve conflicts
using WinForm = System.Windows.Forms.Form;
using WinPoint = System.Drawing.Point;
using WinSize = System.Drawing.Size;
using WinColor = System.Drawing.Color;
using WinFont = System.Drawing.Font;
using WinControl = System.Windows.Forms.Control;

public class CopyElementsProgressForm : WinForm
{
    private ProgressBar progressBar;
    private Label statusLabel;
    private Label currentOperationLabel;
    private ListBox detailsListBox;
    private Button cancelButton;
    private Label elementCountLabel;
    private Label timeLabel;
    private Label phaseLabel;
    private Stopwatch stopwatch;
    private System.Windows.Forms.Timer updateTimer;

    public bool IsCancelled { get; private set; }
    private int _totalElements = 0;
    private int _elementsInGroups = 0;
    private int _elementsCopied = 0;

    public CopyElementsProgressForm()
    {
        InitializeComponent();
        stopwatch = new Stopwatch();
        stopwatch.Start();

        updateTimer = new System.Windows.Forms.Timer();
        updateTimer.Interval = 100;
        updateTimer.Tick += (s, e) => UpdateElapsedTime();
        updateTimer.Start();
    }

    private void InitializeComponent()
    {
        this.Text = "Copy Elements Along Groups - Progress";
        this.Size = new WinSize(650, 550);
        this.MinimumSize = new WinSize(500, 400);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        progressBar = new ProgressBar
        {
            Location = new WinPoint(12, 12),
            Size = new WinSize(this.ClientSize.Width - 24, 23),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Style = ProgressBarStyle.Continuous
        };

        phaseLabel = new Label
        {
            Location = new WinPoint(12, 40),
            Size = new WinSize(this.ClientSize.Width - 24, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "Phase: Initializing...",
            Font = new WinFont(this.Font, FontStyle.Bold)
        };

        statusLabel = new Label
        {
            Location = new WinPoint(12, 65),
            Size = new WinSize(this.ClientSize.Width - 24, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "Status: Starting..."
        };

        currentOperationLabel = new Label
        {
            Location = new WinPoint(12, 90),
            Size = new WinSize(this.ClientSize.Width - 24, 40),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "Current: ",
            AutoSize = false
        };

        timeLabel = new Label
        {
            Location = new WinPoint(this.ClientSize.Width - 172 - 12, 135),
            Size = new WinSize(172, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Text = "Elapsed: 00:00",
            TextAlign = ContentAlignment.TopRight
        };

        elementCountLabel = new Label
        {
            Location = new WinPoint(12, 135),
            Size = new WinSize(400, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Left,
            Text = "Elements: 0 selected | 0 in groups | 0 copied",
            Font = new WinFont(this.Font.FontFamily, 9, FontStyle.Regular)
        };

        detailsListBox = new ListBox
        {
            Location = new WinPoint(12, 160),
            Size = new WinSize(this.ClientSize.Width - 24, this.ClientSize.Height - 160 - 35 - 12),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            HorizontalScrollbar = true,
            Font = new WinFont("Consolas", 9),
            SelectionMode = SelectionMode.MultiExtended
        };

        cancelButton = new Button
        {
            Location = new WinPoint(this.ClientSize.Width - 75 - 12, this.ClientSize.Height - 23 - 12),
            Size = new WinSize(75, 23),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Text = "Cancel",
            DialogResult = DialogResult.Cancel
        };
        cancelButton.Click += (sender, e) =>
        {
            if (MessageBox.Show("Are you sure you want to cancel the operation?", 
                "Cancel Operation", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question) == DialogResult.Yes)
            {
                IsCancelled = true;
                this.DialogResult = DialogResult.Cancel;
                AddDetail("Operation cancelled by user", DetailType.Warning);
            }
        };

        this.Controls.AddRange(new WinControl[]
        {
            progressBar,
            phaseLabel,
            statusLabel,
            currentOperationLabel,
            timeLabel,
            elementCountLabel,
            detailsListBox,
            cancelButton
        });

        this.KeyPreview = true;
        this.KeyDown += new KeyEventHandler(Form_KeyDown);
    }

    private void Form_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Control && e.KeyCode == Keys.C)
        {
            if (detailsListBox.SelectedItems.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                foreach (var item in detailsListBox.SelectedItems)
                {
                    sb.AppendLine(item.ToString());
                }
                Clipboard.SetText(sb.ToString());
            }
        }
        else if (e.KeyCode == Keys.Escape)
        {
            cancelButton.PerformClick();
        }
    }

    private void UpdateElapsedTime()
    {
        if (!this.IsDisposed)
        {
            var elapsed = stopwatch.Elapsed;
            timeLabel.Text = $"Elapsed: {elapsed:mm\\:ss\\.ff}";
        }
    }

    public void SetPhase(string phaseName)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => SetPhase(phaseName)));
            return;
        }

        phaseLabel.Text = $"Phase: {phaseName}";
        AddDetail($"Starting phase: {phaseName}", DetailType.Phase);
    }

    public void UpdateProgress(int current, int total, string status, string currentOperation = "")
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => UpdateProgress(current, total, status, currentOperation)));
            return;
        }

        if (total > 0)
        {
            progressBar.Maximum = total;
            progressBar.Value = Math.Min(current, total);
            
            int percentage = (current * 100 / total);
            this.Text = $"Copy Elements Along Groups - {percentage}%";
        }

        statusLabel.Text = $"Status: {status}";
        
        if (!string.IsNullOrEmpty(currentOperation))
        {
            currentOperationLabel.Text = $"Current: {currentOperation}";
        }

        // Update estimated time remaining
        if (current > 0 && total > 0 && stopwatch.ElapsedMilliseconds > 1000)
        {
            var avgTimePerItem = stopwatch.ElapsedMilliseconds / (double)current;
            var remainingItems = total - current;
            var estimatedMs = avgTimePerItem * remainingItems;
            var estimatedTime = TimeSpan.FromMilliseconds(estimatedMs);

            if (estimatedTime.TotalSeconds > 2)
            {
                statusLabel.Text += $" (Est. remaining: {estimatedTime:mm\\:ss})";
            }
        }
    }

    public void UpdateElementCounts(int totalSelected, int inGroups, int copied)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => UpdateElementCounts(totalSelected, inGroups, copied)));
            return;
        }

        _totalElements = totalSelected;
        _elementsInGroups = inGroups;
        _elementsCopied = copied;

        elementCountLabel.Text = $"Elements: {totalSelected} selected | {inGroups} in groups | {copied} copied";

        if (copied > 0)
        {
            elementCountLabel.ForeColor = System.Drawing.Color.Green;
        }
    }

    public enum DetailType
    {
        Info,
        Success,
        Warning,
        Error,
        Phase
    }

public void AddDetail(string detail, DetailType type = DetailType.Info)
{
    if (this.InvokeRequired)
    {
        this.Invoke(new Action(() => AddDetail(detail, type)));
        return;
    }

    string timestamp = DateTime.Now.ToString("HH:mm:ss.ff");
    
    // Use traditional switch statement instead of switch expression
    string prefix;
    switch (type)
    {
        case DetailType.Success:
            prefix = "[✓]";
            break;
        case DetailType.Warning:
            prefix = "[!]";
            break;
        case DetailType.Error:
            prefix = "[✗]";
            break;
        case DetailType.Phase:
            prefix = "[►]";
            break;
        default:
            prefix = "[·]";
            break;
    }

    string entry = $"[{timestamp}] {prefix} {detail}";
    detailsListBox.Items.Add(entry);
    
    // Auto-scroll to bottom
    detailsListBox.TopIndex = Math.Max(0, detailsListBox.Items.Count - 1);

    // Flash color for important events
    if (type == DetailType.Success || type == DetailType.Error)
    {
        var originalColor = detailsListBox.BackColor;
        var flashColor = type == DetailType.Success ? 
            System.Drawing.Color.FromArgb(230, 255, 230) : 
            System.Drawing.Color.FromArgb(255, 230, 230);
        
        detailsListBox.BackColor = flashColor;
        Application.DoEvents();
        
        System.Threading.Tasks.Task.Delay(200).ContinueWith(t =>
        {
            if (!this.IsDisposed)
            {
                this.Invoke(new Action(() => detailsListBox.BackColor = originalColor));
            }
        });
    }
}

    public void AddGroupTypeProcessed(string groupTypeName, int instanceCount)
    {
        AddDetail($"Processing group type: {groupTypeName} ({instanceCount} instances)", DetailType.Info);
    }

    public void AddCopyOperation(string sourceGroup, string targetGroup, int elementCount)
    {
        AddDetail($"Copying {elementCount} elements: {sourceGroup} → {targetGroup}", DetailType.Success);
    }

    public void SetComplete(int totalCopied, long elapsedMs, bool isCancelled = false)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => SetComplete(totalCopied, elapsedMs, isCancelled)));
            return;
        }

        stopwatch.Stop();
        updateTimer.Stop();

        progressBar.Value = progressBar.Maximum;
        
        if (isCancelled)
        {
            this.Text = "Copy Elements - Cancelled";
            phaseLabel.Text = "Phase: Cancelled";
            statusLabel.Text = $"Operation cancelled. Copied {totalCopied} elements before cancellation.";
            AddDetail($"Operation cancelled. Total time: {elapsedMs}ms", DetailType.Warning);
        }
        else
        {
            this.Text = "Copy Elements - Complete";
            phaseLabel.Text = "Phase: Complete";
            statusLabel.Text = $"Successfully copied {totalCopied} elements in {elapsedMs}ms";
            AddDetail($"Operation completed successfully. Total time: {elapsedMs}ms", DetailType.Success);
        }

        // Change button to Close
        cancelButton.Text = "Close";
        cancelButton.Click -= null; // Remove all event handlers
        cancelButton.Click += (sender, e) => { this.Close(); };
        
        IsCancelled = false;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            updateTimer?.Stop();
            updateTimer?.Dispose();
            stopwatch?.Stop();
        }
        base.Dispose(disposing);
    }
}
