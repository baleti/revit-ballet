using System;
using System.Diagnostics;
using System.Windows.Forms;

/// <summary>
/// Provides a cancellable progress dialog that appears after a delay
/// Works with Revit's UI thread model by using a modeless form
/// </summary>
public class CancellableProgressDialog : IDisposable
{
    private Form progressForm;
    private Label statusLabel;
    private ProgressBar progressBar;
    private Label progressLabel;
    private Button cancelButton;
    private bool isCancelled = false;
    private bool isShown = false;
    private readonly int delayMilliseconds;
    private readonly string operationName;
    private readonly Stopwatch stopwatch;
    private int totalItems = 0;
    private int processedItems = 0;

    /// <summary>
    /// Gets whether the user has requested cancellation
    /// </summary>
    public bool IsCancelled => isCancelled;

    /// <summary>
    /// Creates a new cancellable progress dialog
    /// </summary>
    /// <param name="operationName">Name of the operation being performed</param>
    /// <param name="delayMilliseconds">Delay before showing dialog (default 3000ms)</param>
    public CancellableProgressDialog(string operationName, int delayMilliseconds = 3000)
    {
        this.operationName = operationName;
        this.delayMilliseconds = delayMilliseconds;
        this.stopwatch = new Stopwatch();
    }

    /// <summary>
    /// Starts the progress dialog timer
    /// </summary>
    public void Start()
    {
        stopwatch.Start();
    }

    /// <summary>
    /// Checks if dialog should be shown and shows it if needed
    /// Call this periodically from your processing loop
    /// </summary>
    public void CheckAndShow()
    {
        if (!isShown && !isCancelled && stopwatch.ElapsedMilliseconds >= delayMilliseconds)
        {
            ShowDialog();
        }

        // Process Windows messages to keep form responsive
        if (isShown && progressForm != null && !progressForm.IsDisposed)
        {
            Application.DoEvents();
        }
    }

    /// <summary>
    /// Sets the total number of items to process
    /// </summary>
    public void SetTotal(int total)
    {
        totalItems = total;
        processedItems = 0;
        UpdateProgressDisplay();
    }

    /// <summary>
    /// Increments the progress counter
    /// </summary>
    public void IncrementProgress()
    {
        processedItems++;
        UpdateProgressDisplay();
    }

    /// <summary>
    /// Updates the progress display (bar and label)
    /// </summary>
    private void UpdateProgressDisplay()
    {
        if (isShown && progressForm != null && !progressForm.IsDisposed)
        {
            if (totalItems > 0)
            {
                int percentage = (int)((double)processedItems / totalItems * 100);
                if (progressBar != null)
                {
                    progressBar.Value = Math.Min(percentage, 100);
                }
                if (progressLabel != null)
                {
                    progressLabel.Text = $"{processedItems:N0} / {totalItems:N0} ({percentage}%)";
                }
            }
            else
            {
                // Indeterminate progress
                if (progressLabel != null)
                {
                    progressLabel.Text = $"{processedItems:N0} items processed";
                }
            }
        }
    }

    /// <summary>
    /// Updates the status message
    /// </summary>
    public void UpdateStatus(string message)
    {
        if (isShown && progressForm != null && !progressForm.IsDisposed)
        {
            if (statusLabel != null)
            {
                statusLabel.Text = message;
            }
        }
    }

    private void ShowDialog()
    {
        if (isShown || isCancelled)
            return;

        isShown = true;

        // Create modeless form on the current (UI) thread
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

        // Show modeless (non-blocking)
        progressForm.Show();

        // Update progress display with current values
        UpdateProgressDisplay();
    }

    /// <summary>
    /// Closes the dialog
    /// </summary>
    public void Dispose()
    {
        stopwatch.Stop();

        if (isShown && progressForm != null && !progressForm.IsDisposed)
        {
            progressForm.Close();
            progressForm.Dispose();
        }
    }
}
