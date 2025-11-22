using ExcelDna.Integration;
using NLog;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Domino;

/// <summary>
/// Main Excel-DNA Add-In class that initializes the Domino cell tracking system.
/// This add-in is designed for financial services compliance and auditing.
/// </summary>
public class AddIn : IExcelAddIn
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private Excel.Application? _excelApp;
    private ChangeTracker? _changeTracker;

    public void AutoOpen()
    {
        try
        {
            Logger.Info("===============================================");
            Logger.Info("Domino Add-In Starting...");
            Logger.Info($"Version: {GetType().Assembly.GetName().Version}");
            Logger.Info($"Machine: {Environment.MachineName}");
            Logger.Info($"User: {Environment.UserName}");
            Logger.Info("===============================================");

            // Get the Excel Application object
            _excelApp = (Excel.Application)ExcelDnaUtil.Application;

            if (_excelApp == null)
            {
                Logger.Fatal("Failed to get Excel Application object");
                throw new InvalidOperationException("Excel Application is not available");
            }

            // Initialize the change tracker
            _changeTracker = new ChangeTracker(_excelApp);
            _changeTracker.StartTracking();

            Logger.Info("Domino Add-In successfully initialized");
            Logger.Info($"Tracking cells A1:D4 in all sheets of all workbooks");
            Logger.Info($"Log file location: {GetLogFilePath()}");
        }
        catch (Exception ex)
        {
            Logger.Fatal(ex, "Critical error during add-in initialization");
            throw;
        }
    }

    public void AutoClose()
    {
        try
        {
            Logger.Info("Domino Add-In shutting down...");

            // Stop tracking and cleanup
            _changeTracker?.StopTracking();
            _changeTracker?.Dispose();

            // Release COM objects
            if (_excelApp != null)
            {
                Marshal.ReleaseComObject(_excelApp);
                _excelApp = null;
            }

            Logger.Info("Domino Add-In shutdown complete");
            Logger.Info("===============================================");

            // Flush logs before exit
            LogManager.Shutdown();
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error during add-in shutdown");
        }
    }

    private static string GetLogFilePath()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(localAppData, "Domino", "Logs", $"domino-{DateTime.Now:yyyy-MM-dd}.log");
    }
}
