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
    private bool isComplete = false;
    private bool userScrolling = false;
    private bool userHasScrolledUp = false;
    private Timer scrollCheckTimer;
    private int lastItemCount = 0;

    public bool IsCancelled { get; private set; }
    private int _totalElements = 0;
    private int _elementsInGroups = 0;
    private int _elementsCopied = 0;
    private bool _preserveElementCounts = false;

    public CopyElementsProgressForm()
    {
        InitializeComponent();
        stopwatch = new Stopwatch();
        stopwatch.Start();

        updateTimer = new System.Windows.Forms.Timer();
        updateTimer.Interval = 100;
        updateTimer.Tick += (s, e) => UpdateElapsedTime();
        updateTimer.Start();

        scrollCheckTimer = new System.Windows.Forms.Timer();
        scrollCheckTimer.Interval = 200;
        scrollCheckTimer.Tick += (s, e) => CheckAutoScroll();
        scrollCheckTimer.Start();
    }

    private void InitializeComponent()
    {
        this.Text = "Copy Elements Along Groups - Progress";
        this.Size = new WinSize(650, 550);
        this.MinimumSize = new WinSize(500, 400);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MaximizeBox = true; // Enable maximize button
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

        // Track all forms of user scrolling
        bool isDraggingScrollbar = false;
        detailsListBox.MouseWheel += (s, e) =>
        {
            userScrolling = true;
            userHasScrolledUp = true;
        };

        // Track scrollbar dragging
        detailsListBox.MouseDown += (s, e) =>
        {
            // Check if click is on the scrollbar area
            if (e.X >= detailsListBox.ClientSize.Width - SystemInformation.VerticalScrollBarWidth)
            {
                isDraggingScrollbar = true;
                userScrolling = true;
                userHasScrolledUp = true;
            }
        };

        detailsListBox.MouseUp += (s, e) =>
        {
            isDraggingScrollbar = false;
        };
        detailsListBox.MouseMove += (s, e) =>
        {
            if (isDraggingScrollbar)
            {
                 userScrolling = true;
                userHasScrolledUp = true;
            }
        };
        
        // Add KeyDown event handler directly to ListBox for better key capture
        detailsListBox.KeyDown += DetailsListBox_KeyDown;

        cancelButton = new Button
        {
            Location = new WinPoint(this.ClientSize.Width - 75 - 12, this.ClientSize.Height - 23 - 12),
            Size = new WinSize(75, 23),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Text = "Cancel",
            DialogResult = DialogResult.Cancel
        };
        cancelButton.Click += CancelButton_Click;

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

    private void CancelButton_Click(object sender, EventArgs e)
    {
        if (isComplete)
        {
            // Just close if operation is complete
            this.Close();
        }
        else
        {
            // Show confirmation only if operation is still running
            if (MessageBox.Show("Are you sure you want to cancel the operation?",
                "Cancel Operation",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes)
            {
                IsCancelled = true;
                this.DialogResult = DialogResult.Cancel;
                AddDetail("Operation cancelled by user", DetailType.Warning);
            }
        }
    }

    private void Form_KeyDown(object sender, KeyEventArgs e)
    {
        // Handle Ctrl+A for Select All
        if (e.Control && e.KeyCode == Keys.A)
        {
            SelectAllItems();
            e.Handled = true;
        }
        // Handle Ctrl+C for Copy
        else if (e.Control && e.KeyCode == Keys.C)
        {
            CopySelectedItemsToClipboard();
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Escape)
        {
            if (isComplete)
            {
                this.Close();
            }
            else
            {
                cancelButton.PerformClick();
            }
        }
    }

    private void DetailsListBox_KeyDown(object sender, KeyEventArgs e)
    {
        // Also handle keyboard shortcuts when ListBox has focus
        // This ensures shortcuts work even when the ListBox is focused
        if (e.Control && e.KeyCode == Keys.A)
        {
            SelectAllItems();
            e.Handled = true;
        }
        else if (e.Control && e.KeyCode == Keys.C)
        {
            CopySelectedItemsToClipboard();
            e.Handled = true;
        }
    }

    private void SelectAllItems()
    {
        if (detailsListBox.Items.Count > 0)
        {
            detailsListBox.BeginUpdate();
            for (int i = 0; i < detailsListBox.Items.Count; i++)
            {
                detailsListBox.SetSelected(i, true);
            }
            detailsListBox.EndUpdate();
        }
    }

    private void CopySelectedItemsToClipboard()
    {
        if (detailsListBox.SelectedItems.Count > 0)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var item in detailsListBox.SelectedItems)
            {
                sb.AppendLine(item.ToString());
            }
            try
            {
                Clipboard.SetText(sb.ToString());
                
                // Optional: Show brief confirmation in status if complete
                if (isComplete)
                {
                    string originalStatus = statusLabel.Text;
                    statusLabel.Text = $"{originalStatus} (Copied {detailsListBox.SelectedItems.Count} items to clipboard)";
                    
                    // Reset status after 2 seconds
                    System.Windows.Forms.Timer resetTimer = new System.Windows.Forms.Timer();
                    resetTimer.Interval = 2000;
                    resetTimer.Tick += (s, e) =>
                    {
                        statusLabel.Text = originalStatus;
                        resetTimer.Stop();
                        resetTimer.Dispose();
                    };
                    resetTimer.Start();
                }
            }
            catch (Exception ex)
            {
                // Clipboard operation might fail, show message
                MessageBox.Show($"Failed to copy to clipboard: {ex.Message}", "Copy Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    private void UpdateElapsedTime()
    {
        if (!this.IsDisposed && !isComplete)
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

        // The preserve flag is meant to prevent the counts from being reset to 0 during copying
        // But we should still update them if we receive actual non-zero values

        if (!_preserveElementCounts)
        {
            // Normal mode: always update
            _totalElements = totalSelected;
            _elementsInGroups = inGroups;
        }
        else
        {
            // Preserve mode: only update if we're getting real values (not zeros)
            // This prevents the counts from being reset during the copy phase
            if (totalSelected > 0)
                _totalElements = totalSelected;
            if (inGroups > 0)
                _elementsInGroups = inGroups;
        }

        _elementsCopied = copied;

        // Use the stored values for display
        elementCountLabel.Text = $"Elements: {_totalElements} selected | {_elementsInGroups} in groups | {copied} copied";

        // Color indication
        if (copied > 0 && _totalElements > 0)
        {
            elementCountLabel.ForeColor = System.Drawing.Color.DarkGreen;
        }
    }

    public enum DetailType
    {
        Info,
        Success,
        Warning,
        Error,
        Phase,
        Progress
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
            case DetailType.Progress:
                prefix = "[...]";
                break;
            default:
                prefix = "[·]";
                break;
        }

        string entry = $"[{timestamp}] {prefix} {detail}";
        detailsListBox.Items.Add(entry);

        // Check if we should auto-scroll
        CheckForAutoScroll();
    }

    private void CheckAutoScroll()
    {
        if (detailsListBox.Items.Count != lastItemCount)
        {
            lastItemCount = detailsListBox.Items.Count;
            CheckForAutoScroll();
        }
    }

    private void CheckForAutoScroll()
    {
        if (isComplete || detailsListBox.Items.Count == 0) return;

        bool atBottom = IsAtBottom();

        // If user scrolled up but is now at bottom again, resume auto-scrolling
        if (userHasScrolledUp && atBottom)
        {
            userScrolling = false;
            userHasScrolledUp = false;
        }

        // Auto-scroll if we haven't manually scrolled or if we're back at bottom
        if (!userScrolling || (!userHasScrolledUp && atBottom))
        {
            detailsListBox.TopIndex = Math.Max(0, detailsListBox.Items.Count - 1);
        }
    }

    private bool IsAtBottom()
    {
        if (detailsListBox.Items.Count == 0) return true;

        // Check if the last item is visible
        int visibleItemCount = detailsListBox.ClientSize.Height / detailsListBox.ItemHeight;
        int lastVisibleIndex = detailsListBox.TopIndex + visibleItemCount - 1;
        return lastVisibleIndex >= detailsListBox.Items.Count - 1;
    }

    public void AddGroupTypeProcessed(string groupTypeName, int instanceCount)
    {
        AddDetail($"Counting Group instances: {groupTypeName} ({instanceCount} instances)", DetailType.Info);
    }

    public void AddCopyOperation(string sourceGroup, string targetGroupInfo)
    {
        AddDetail($"Copying to: {targetGroupInfo}", DetailType.Success);
    }

    public void AddMappingProgress(string message)
    {
        AddDetail(message, DetailType.Info);
    }

    public void AddIntermediateProgress(string message)
    {
        AddDetail(message, DetailType.Progress);
    }

    public void PreserveElementCounts()
    {
        _preserveElementCounts = true;
    }

    public void StopTimer()
    {
        if (!isComplete)
        {
            stopwatch.Stop();
            updateTimer.Stop();
        }
    }

    public void SetComplete(int totalCopied, long elapsedMs, bool isCancelled = false)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => SetComplete(totalCopied, elapsedMs, isCancelled)));
            return;
        }

        isComplete = true;
        stopwatch.Stop();
        updateTimer.Stop();

        // Update elapsed time one final time
        var elapsed = stopwatch.Elapsed;
        timeLabel.Text = $"Elapsed: {elapsed:mm\\:ss\\.ff}";

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
            statusLabel.Text = $"Successfully copied {totalCopied} elements";
            AddDetail($"Operation completed successfully", DetailType.Success);
        }

        // Change button to Close
        cancelButton.Text = "Close";
        IsCancelled = false;

        // Ensure we can still copy to clipboard after completion
        // Give focus to the ListBox to ensure keyboard shortcuts work
        detailsListBox.Focus();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            updateTimer?.Stop();
            updateTimer?.Dispose();
            scrollCheckTimer?.Stop();
            scrollCheckTimer?.Dispose();
            stopwatch?.Stop();
        }
        base.Dispose(disposing);
    }
}
