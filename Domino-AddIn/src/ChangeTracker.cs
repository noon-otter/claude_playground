using NLog;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Domino;

/// <summary>
/// Core service responsible for tracking all cell changes in the monitored range (A1:D4)
/// across all sheets in all open workbooks.
/// </summary>
public class ChangeTracker : IDisposable
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private static DateTime? _lastChangeTimestamp;

    private readonly Excel.Application _app;
    private readonly Dictionary<string, WorkbookTracker> _workbookTrackers = new();
    private readonly object _lockObject = new();

    // Monitored range configuration
    private const string MONITORED_RANGE = "A1:D4";

    public ChangeTracker(Excel.Application app)
    {
        _app = app ?? throw new ArgumentNullException(nameof(app));
    }

    /// <summary>
    /// Starts tracking cell changes in all open workbooks.
    /// </summary>
    public void StartTracking()
    {
        try
        {
            Logger.Info("Starting change tracking...");

            // Hook application-level events
            _app.WorkbookOpen += OnWorkbookOpen;
            _app.WorkbookBeforeClose += OnWorkbookBeforeClose;
            _app.WorkbookNewSheet += OnWorkbookNewSheet;
            _app.WorkbookActivate += OnWorkbookActivate;

            // Track all currently open workbooks
            foreach (Excel.Workbook workbook in _app.Workbooks)
            {
                TrackWorkbook(workbook);
            }

            Logger.Info($"Change tracking started. Monitoring {_workbookTrackers.Count} workbook(s)");
        }
        catch (Exception ex)
        {
            Logger.Fatal(ex, "Failed to start change tracking");
            throw;
        }
    }

    /// <summary>
    /// Stops tracking and releases all resources.
    /// </summary>
    public void StopTracking()
    {
        try
        {
            Logger.Info("Stopping change tracking...");

            // Unhook application events
            _app.WorkbookOpen -= OnWorkbookOpen;
            _app.WorkbookBeforeClose -= OnWorkbookBeforeClose;
            _app.WorkbookNewSheet -= OnWorkbookNewSheet;
            _app.WorkbookActivate -= OnWorkbookActivate;

            // Stop tracking all workbooks
            lock (_lockObject)
            {
                foreach (var tracker in _workbookTrackers.Values)
                {
                    tracker.Dispose();
                }
                _workbookTrackers.Clear();
            }

            Logger.Info("Change tracking stopped");
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error stopping change tracking");
        }
    }

    /// <summary>
    /// Gets the timestamp of the last tracked change.
    /// </summary>
    public static DateTime? GetLastChangeTimestamp()
    {
        return _lastChangeTimestamp;
    }

    private void TrackWorkbook(Excel.Workbook workbook)
    {
        try
        {
            var workbookName = workbook.Name;

            lock (_lockObject)
            {
                if (_workbookTrackers.ContainsKey(workbookName))
                {
                    Logger.Debug($"Workbook already being tracked: {workbookName}");
                    return;
                }

                var tracker = new WorkbookTracker(workbook, OnCellChange);
                _workbookTrackers[workbookName] = tracker;

                Logger.Info($"Started tracking workbook: {workbookName}");
                LogChange("WORKBOOK_OPEN", workbookName, "", "", "");
            }
        }
        catch (Exception ex)
        {
            Logger.Error(ex, $"Error tracking workbook: {workbook.Name}");
        }
    }

    private void UntrackWorkbook(Excel.Workbook workbook)
    {
        try
        {
            var workbookName = workbook.Name;

            lock (_lockObject)
            {
                if (_workbookTrackers.TryGetValue(workbookName, out var tracker))
                {
                    tracker.Dispose();
                    _workbookTrackers.Remove(workbookName);

                    Logger.Info($"Stopped tracking workbook: {workbookName}");
                    LogChange("WORKBOOK_CLOSE", workbookName, "", "", "");
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error(ex, $"Error untracking workbook: {workbook.Name}");
        }
    }

    // Event Handlers
    private void OnWorkbookOpen(Excel.Workbook workbook)
    {
        TrackWorkbook(workbook);
    }

    private void OnWorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
    {
        UntrackWorkbook(workbook);
    }

    private void OnWorkbookNewSheet(Excel.Workbook workbook, object sheet)
    {
        // When a new sheet is added, refresh tracking for the workbook
        if (sheet is Excel.Worksheet)
        {
            Logger.Info($"New sheet added to workbook: {workbook.Name}");
            UntrackWorkbook(workbook);
            TrackWorkbook(workbook);
        }
    }

    private void OnWorkbookActivate(Excel.Workbook workbook)
    {
        // Ensure workbook is being tracked when activated
        if (!_workbookTrackers.ContainsKey(workbook.Name))
        {
            TrackWorkbook(workbook);
        }
    }

    private void OnCellChange(string workbookName, string sheetName, string cellAddress, string changeType, string oldValue, string newValue)
    {
        LogChange(changeType, workbookName, sheetName, cellAddress, $"{oldValue} → {newValue}");
    }

    private static void LogChange(string eventType, string workbookName, string sheetName, string cellAddress, string details)
    {
        _lastChangeTimestamp = DateTime.Now;

        var logMessage = string.IsNullOrEmpty(sheetName)
            ? $"[{eventType}] Workbook: {workbookName}"
            : $"[{eventType}] Workbook: {workbookName} | Sheet: {sheetName} | Cell: {cellAddress} | {details}";

        Logger.Info(logMessage);
    }

    public void Dispose()
    {
        StopTracking();
        GC.SuppressFinalize(this);
    }
}

/// <summary>
/// Tracks changes for a specific workbook and all its worksheets.
/// </summary>
internal class WorkbookTracker : IDisposable
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

    private readonly Excel.Workbook _workbook;
    private readonly Action<string, string, string, string, string, string> _onCellChange;
    private readonly List<WorksheetTracker> _worksheetTrackers = new();

    public WorkbookTracker(Excel.Workbook workbook, Action<string, string, string, string, string, string> onCellChange)
    {
        _workbook = workbook;
        _onCellChange = onCellChange;

        // Track all existing worksheets
        foreach (Excel.Worksheet sheet in _workbook.Worksheets)
        {
            TrackWorksheet(sheet);
        }
    }

    private void TrackWorksheet(Excel.Worksheet worksheet)
    {
        try
        {
            var tracker = new WorksheetTracker(
                _workbook.Name,
                worksheet,
                _onCellChange);

            _worksheetTrackers.Add(tracker);

            Logger.Debug($"Tracking worksheet: {worksheet.Name} in {_workbook.Name}");
        }
        catch (Exception ex)
        {
            Logger.Error(ex, $"Error tracking worksheet: {worksheet.Name}");
        }
    }

    public void Dispose()
    {
        foreach (var tracker in _worksheetTrackers)
        {
            tracker.Dispose();
        }
        _worksheetTrackers.Clear();
    }
}

/// <summary>
/// Tracks changes for a specific worksheet, monitoring the A1:D4 range.
/// </summary>
internal class WorksheetTracker : IDisposable
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

    private readonly string _workbookName;
    private readonly Excel.Worksheet _worksheet;
    private readonly Action<string, string, string, string, string, string> _onCellChange;

    private const string MONITORED_RANGE = "A1:D4";

    public WorksheetTracker(
        string workbookName,
        Excel.Worksheet worksheet,
        Action<string, string, string, string, string, string> onCellChange)
    {
        _workbookName = workbookName;
        _worksheet = worksheet;
        _onCellChange = onCellChange;

        // Subscribe to worksheet events
        _worksheet.Change += OnWorksheetChange;
    }

    private void OnWorksheetChange(Excel.Range target)
    {
        try
        {
            // Check if the change intersects with our monitored range
            var monitoredRange = _worksheet.Range[MONITORED_RANGE];
            var intersection = Application.Intersect(target, monitoredRange);

            if (intersection != null)
            {
                // A cell in our monitored range has changed
                ProcessChange(intersection);
                Marshal.ReleaseComObject(intersection);
            }

            Marshal.ReleaseComObject(monitoredRange);
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error processing worksheet change");
        }
    }

    private void ProcessChange(Excel.Range changedRange)
    {
        try
        {
            foreach (Excel.Range cell in changedRange.Cells)
            {
                try
                {
                    var address = cell.Address[false, false];
                    var value = cell.Value?.ToString() ?? "";
                    var formula = cell.Formula?.ToString() ?? "";

                    // Determine change type
                    var changeType = !string.IsNullOrEmpty(formula) && formula.StartsWith("=")
                        ? "FORMULA_CHANGE"
                        : "VALUE_CHANGE";

                    var displayValue = !string.IsNullOrEmpty(formula) && formula.StartsWith("=")
                        ? formula
                        : value;

                    _onCellChange(
                        _workbookName,
                        _worksheet.Name,
                        address,
                        changeType,
                        "", // Old value not easily accessible in real-time
                        displayValue);
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, "Error processing individual cell change");
                }
                finally
                {
                    Marshal.ReleaseComObject(cell);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error processing changed range");
        }
    }

    public void Dispose()
    {
        try
        {
            _worksheet.Change -= OnWorksheetChange;
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error disposing worksheet tracker");
        }
    }
}

/// <summary>
/// Helper class for Excel interop operations.
/// </summary>
internal static class Application
{
    /// <summary>
    /// Gets the intersection of two ranges.
    /// </summary>
    public static Excel.Range? Intersect(Excel.Range range1, Excel.Range range2)
    {
        try
        {
            var app = range1.Application;
            return app.Intersect(range1, range2);
        }
        catch
        {
            return null;
        }
    }
}
