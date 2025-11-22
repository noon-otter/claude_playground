using ExcelDna.Integration.CustomUI;
using NLog;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Domino;

/// <summary>
/// Controller for the Domino custom ribbon UI.
/// Provides user interface for monitoring tracked changes.
/// </summary>
[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private IRibbonUI? _ribbon;

    public override string GetCustomUI(string RibbonID)
    {
        Logger.Debug($"GetCustomUI called with RibbonID: {RibbonID}");
        // Return empty string as we define UI in .dna file
        return string.Empty;
    }

    public void OnLoad(IRibbonUI ribbon)
    {
        _ribbon = ribbon;
        Logger.Info("Ribbon UI loaded successfully");

        // Start a timer to refresh the timestamp display periodically
        var timer = new System.Timers.Timer(5000); // Refresh every 5 seconds
        timer.Elapsed += (s, e) => RefreshRibbon();
        timer.Start();
    }

    /// <summary>
    /// Gets the timestamp of the last tracked change for display in the ribbon.
    /// </summary>
    public string GetLastChangeTimestamp(IRibbonControl control)
    {
        try
        {
            var timestamp = ChangeTracker.GetLastChangeTimestamp();
            if (timestamp.HasValue)
            {
                // Format for display
                var timeAgo = DateTime.Now - timestamp.Value;

                if (timeAgo.TotalSeconds < 60)
                    return $"{(int)timeAgo.TotalSeconds}s ago";
                else if (timeAgo.TotalMinutes < 60)
                    return $"{(int)timeAgo.TotalMinutes}m ago";
                else if (timeAgo.TotalHours < 24)
                    return $"{(int)timeAgo.TotalHours}h ago";
                else
                    return timestamp.Value.ToString("MM/dd HH:mm");
            }

            return "No changes yet";
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error getting last change timestamp");
            return "Error";
        }
    }

    /// <summary>
    /// Refreshes the ribbon UI to update the timestamp display.
    /// </summary>
    public void RefreshTimestamp(IRibbonControl control)
    {
        Logger.Debug("Manual ribbon refresh requested");
        RefreshRibbon();
    }

    /// <summary>
    /// Opens the log file in the default text editor.
    /// </summary>
    public void OpenLogFile(IRibbonControl control)
    {
        try
        {
            var logFilePath = GetLogFilePath();

            if (File.Exists(logFilePath))
            {
                Logger.Info($"Opening log file: {logFilePath}");
                Process.Start(new ProcessStartInfo
                {
                    FileName = logFilePath,
                    UseShellExecute = true
                });
            }
            else
            {
                var logDir = Path.GetDirectoryName(logFilePath) ?? "";
                MessageBox.Show(
                    $"Log file not found.\n\nExpected location:\n{logFilePath}\n\nNo changes have been tracked yet, or the log directory doesn't exist.",
                    "Domino - Log File Not Found",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // Open the logs directory if it exists
                if (Directory.Exists(logDir))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = logDir,
                        UseShellExecute = true
                    });
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error opening log file");
            MessageBox.Show(
                $"Error opening log file:\n{ex.Message}",
                "Domino - Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// Shows information about the Domino add-in.
    /// </summary>
    public void ShowAbout(IRibbonControl control)
    {
        var version = GetType().Assembly.GetName().Version;
        var logPath = GetLogFilePath();

        MessageBox.Show(
            $"Domino Excel Add-In\n" +
            $"Version: {version}\n\n" +
            $"Unintrusive cell change tracking for compliance and auditing.\n\n" +
            $"Tracked Range: A1:D4 (all sheets, all workbooks)\n" +
            $"Tracked Events: Cell changes, workbook open/close\n\n" +
            $"Log Location:\n{logPath}\n\n" +
            $"© {DateTime.Now.Year} Your Organization",
            "About Domino",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);

        Logger.Info("About dialog displayed");
    }

    private void RefreshRibbon()
    {
        try
        {
            _ribbon?.Invalidate();
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error refreshing ribbon");
        }
    }

    private static string GetLogFilePath()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(localAppData, "Domino", "Logs", $"domino-{DateTime.Now:yyyy-MM-dd}.log");
    }
}
