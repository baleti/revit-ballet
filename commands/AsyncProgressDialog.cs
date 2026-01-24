using System;
using System.Windows.Forms;

/// <summary>
/// Async progress dialog that doesn't block the work thread.
/// Work thread just updates a counter, dialog polls it on a timer.
/// </summary>
public class AsyncProgressDialog : IDisposable
{
    private Form progressForm;
    private Label statusLabel;
    private ProgressBar progressBar;
    private Label progressLabel;
    private Button cancelButton;
    private System.Windows.Forms.Timer updateTimer;

    private volatile bool isCancelled = false;
    private volatile int currentProgress = 0;
    private volatile int totalItems = 0;
    private volatile bool isShown = false;

    private readonly int delayMilliseconds;
    private readonly string operationName;
    private readonly System.Diagnostics.Stopwatch stopwatch;

    public bool IsCancelled => isCancelled;

    public AsyncProgressDialog(string operationName, int delayMilliseconds = 1500)
    {
        this.operationName = operationName;
        this.delayMilliseconds = delayMilliseconds;
        this.stopwatch = new System.Diagnostics.Stopwatch();
    }

    public void Start()
    {
        stopwatch.Start();
    }

    /// <summary>
    /// Call this occasionally to check if dialog should show (doesn't block)
    /// </summary>
    public void CheckIfShouldShow()
    {
        if (!isShown && !isCancelled && stopwatch.ElapsedMilliseconds >= delayMilliseconds)
        {
            ShowDialog();
        }
    }

    public void SetTotal(int total)
    {
        totalItems = total;
    }

    /// <summary>
    /// Set current progress (thread-safe, non-blocking)
    /// </summary>
    public void SetProgress(int progress)
    {
        currentProgress = progress;
    }

    private void ShowDialog()
    {
        if (isShown || isCancelled)
            return;

        isShown = true;

        progressForm = new Form
        {
            Text = "Operation in Progress",
            Width = 400,
            Height = 200,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterScreen,
            MaximizeBox = false,
            MinimizeBox = false,
            TopMost = true,
            ShowInTaskbar = false
        };

        statusLabel = new Label
        {
            Text = $"Processing: {operationName}",
            AutoSize = false,
            Width = 360,
            Height = 25,
            Left = 20,
            Top = 20,
            TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        };

        progressBar = new ProgressBar
        {
            Width = 360,
            Height = 25,
            Left = 20,
            Top = 50,
            Minimum = 0,
            Maximum = 100,
            Value = 0,
            Style = ProgressBarStyle.Continuous
        };

        progressLabel = new Label
        {
            Text = "Starting...",
            AutoSize = false,
            Width = 360,
            Height = 20,
            Left = 20,
            Top = 80,
            TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        };

        cancelButton = new Button
        {
            Text = "Cancel",
            Width = 100,
            Height = 30,
            Left = 150,
            Top = 110
        };

        cancelButton.Click += (s, e) =>
        {
            isCancelled = true;
            cancelButton.Enabled = false;
            cancelButton.Text = "Cancelling...";
        };

        progressForm.Controls.Add(statusLabel);
        progressForm.Controls.Add(progressBar);
        progressForm.Controls.Add(progressLabel);
        progressForm.Controls.Add(cancelButton);

        progressForm.FormClosing += (s, e) =>
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                isCancelled = true;
                e.Cancel = false;
            }
        };

        // Timer updates the dialog independently
        updateTimer = new System.Windows.Forms.Timer();
        updateTimer.Interval = 100; // Update UI every 100ms
        updateTimer.Tick += UpdateTimerTick;
        updateTimer.Start();

        progressForm.Show();

        // Force initial paint by processing messages once
        Application.DoEvents();
        progressForm.Refresh();
    }

    private void UpdateTimerTick(object sender, EventArgs e)
    {
        if (progressForm == null || progressForm.IsDisposed)
            return;

        int current = currentProgress;
        int total = totalItems;

        if (total > 0)
        {
            int percentage = (int)((double)current / total * 100);
            progressBar.Value = Math.Min(percentage, 100);
            progressLabel.Text = $"{current:N0} / {total:N0} ({percentage}%)";
        }
        else
        {
            progressLabel.Text = $"{current:N0} items processed";
        }

        // Process messages to keep dialog responsive
        Application.DoEvents();
    }

    public void Dispose()
    {
        stopwatch.Stop();

        if (updateTimer != null)
        {
            updateTimer.Stop();
            updateTimer.Dispose();
        }

        if (isShown && progressForm != null && !progressForm.IsDisposed)
        {
            progressForm.Close();
            progressForm.Dispose();
        }
    }
}
