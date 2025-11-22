# Domino Add-In - Quick Start Guide

Get up and running with Domino in 5 minutes.

## For Developers

### 1. Prerequisites
- Windows 10/11
- .NET 6.0 SDK ([download](https://dotnet.microsoft.com/download/dotnet/6.0))
- Excel 2016+ Desktop
- PowerShell 5.1+

### 2. Build

```powershell
# Clone and navigate
cd claude_playground
git checkout claude/domino-cell-tracking-poc
cd Domino-AddIn

# Build using the build script
.\build.ps1 -Configuration Release

# Or build manually
dotnet build -c Release
```

### 3. Install

```powershell
# Find the output file
# Location: bin\Release\net6.0-windows\Domino-AddIn64.xll

# Install in Excel
1. Open Excel
2. File → Options → Add-ins
3. Manage: Excel Add-ins → Go
4. Browse → Select Domino-AddIn64.xll
5. OK
```

### 4. Verify

1. Look for **"Domino"** tab in Excel ribbon
2. Open a workbook
3. Change a cell in range **A1:D4**
4. Click **"View Logs"** button in Domino ribbon
5. See your change logged!

## For End Users

### Installing the Add-In

**Option 1: IT Department Installed (Recommended)**
- The add-in should auto-load when you open Excel
- Look for the "Domino" tab in the ribbon

**Option 2: Manual Installation**
1. Download `Domino-AddIn64.xll` from your IT department
2. Open Excel
3. Go to **File → Options**
4. Click **Add-ins** (left sidebar)
5. At the bottom, select **Excel Add-ins** → **Go**
6. Click **Browse**
7. Find and select `Domino-AddIn64.xll`
8. Click **OK**

### Using Domino

**The add-in works automatically!** You don't need to do anything.

#### What Gets Tracked?
- Cell value changes in range **A1:D4**
- Cell formula changes in range **A1:D4**
- Cell format changes in range **A1:D4**
- Workbook open and close events
- **All sheets** in **all open workbooks**

#### Viewing Tracked Changes

**Method 1: Via Ribbon (Easiest)**
1. Click the **Domino** tab in Excel
2. Click **View Logs**
3. The log file opens in Notepad

**Method 2: Manual**
1. Press `Win + R`
2. Type: `%LOCALAPPDATA%\Domino\Logs`
3. Press Enter
4. Open the most recent `domino-YYYY-MM-DD.log` file

#### Understanding Log Entries

```
2025-11-22 14:23:45|INFO|[VALUE_CHANGE] Workbook: Budget.xlsx | Sheet: Q1 | Cell: A1 |  → 1000
```

- **Timestamp**: When the change occurred
- **Change Type**: VALUE_CHANGE, FORMULA_CHANGE, WORKBOOK_OPEN, etc.
- **Workbook**: Which Excel file
- **Sheet**: Which worksheet
- **Cell**: Which cell (e.g., A1, B2)
- **Value**: The new value or formula

#### Ribbon Features

- **Tracking Status**: Shows "Tracking Active"
- **Last Change**: Shows how long ago the last change was made
- **Refresh**: Updates the timestamp display
- **View Logs**: Opens the log file
- **About**: Shows version and configuration info

### Troubleshooting

#### Add-in doesn't appear in Excel

**Check if it's loaded:**
1. File → Options → Add-ins
2. Look for "Domino Add-In" in the list
3. If unchecked, check the box

**Still not working?**
- Close all Excel windows
- Reopen Excel
- Try again

#### No "Domino" tab in ribbon

- Make sure you're using Excel Desktop (not Excel Online)
- Excel version must be 2016 or later
- Try: File → Options → Add-ins → Manage: Disabled Items → Enable

#### Changes not being logged

**Verify you're in the right range:**
- Only cells **A1, A2, A3, A4, B1, B2, B3, B4, C1, C2, C3, C4, D1, D2, D3, D4** are tracked
- Other cells are not tracked

**Check log directory:**
1. Press `Win + R`
2. Type: `%LOCALAPPDATA%\Domino\Logs`
3. If folder doesn't exist, the add-in hasn't started properly

#### Excel crashes when opening

- Start Excel in safe mode: `excel.exe /safe`
- Disable Domino add-in temporarily
- Contact your IT department

### Privacy & Security

#### What data is collected?
- Workbook filename
- Sheet name
- Cell address (e.g., A1)
- Cell value or formula
- Timestamp
- Your Windows username
- Your computer name

#### Where is data stored?
- **Locally only**: `%LOCALAPPDATA%\Domino\Logs`
- **Not sent to the internet** (in current version)
- Logs are rotated daily and kept for 30 days

#### Can I disable tracking?
- Uncheck the add-in: File → Options → Add-ins
- Or contact your IT department

### Getting Help

#### Check the logs first
Most issues show error messages in the logs:
1. Open logs: Domino ribbon → View Logs
2. Look for entries with "ERROR" or "WARN"
3. Share with IT support

#### Contact Support
- **IT Help Desk**: [your-helpdesk-email]
- **Add-in Developer**: [your-dev-email]

Include this information:
- Excel version: File → Account → About Excel
- Windows version: Settings → System → About
- Log files from: `%LOCALAPPDATA%\Domino\Logs`
- Screenshot of the issue

## Common Scenarios

### Scenario 1: Tracking Budget Changes

```
User opens Budget2025.xlsx
→ Logged: [WORKBOOK_OPEN] Budget2025.xlsx

User enters 1000 in cell A1
→ Logged: [VALUE_CHANGE] Budget2025.xlsx | Sheet: Q1 | Cell: A1 | → 1000

User enters formula =SUM(A1:A10) in B1
→ Logged: [FORMULA_CHANGE] Budget2025.xlsx | Sheet: Q1 | Cell: B1 | → =SUM(A1:A10)

User closes workbook
→ Logged: [WORKBOOK_CLOSE] Budget2025.xlsx
```

### Scenario 2: Multi-Workbook Tracking

Domino tracks **all open workbooks simultaneously**:

```
User opens Q1-Report.xlsx
→ Tracking started for Q1-Report.xlsx

User opens Q2-Report.xlsx (Q1 still open)
→ Tracking started for Q2-Report.xlsx
→ Both workbooks now being tracked

Changes in either workbook's A1:D4 range are logged
```

### Scenario 3: Team Collaboration

**Your perspective:**
- You make changes to cells A1:D4
- Changes are logged on your computer
- Your logs: `C:\Users\YourName\AppData\Local\Domino\Logs\`

**Your colleague's perspective:**
- They make changes to the same workbook (on their computer)
- Changes are logged on their computer
- Their logs: `C:\Users\ColleagueName\AppData\Local\Domino\Logs\`

**Note**: Each user has their own independent logs. For centralized tracking, contact your IT department about the future API integration feature.

## Tips & Best Practices

### For Power Users

**Viewing logs in real-time:**
```powershell
# PowerShell command to watch log file
Get-Content "$env:LOCALAPPDATA\Domino\Logs\domino-$(Get-Date -Format yyyy-MM-dd).log" -Wait -Tail 20
```

**Searching logs:**
```powershell
# Find all changes to cell A1
Select-String -Path "$env:LOCALAPPDATA\Domino\Logs\*.log" -Pattern "Cell: A1"

# Find all changes in specific workbook
Select-String -Path "$env:LOCALAPPDATA\Domino\Logs\*.log" -Pattern "Workbook: Budget2025.xlsx"
```

**Exporting logs:**
```powershell
# Copy all logs to desktop
Copy-Item "$env:LOCALAPPDATA\Domino\Logs\*" "$env:USERPROFILE\Desktop\Domino-Logs-Backup\" -Recurse
```

### For Compliance Officers

**Audit trail verification:**
1. Logs include timestamps, user, and machine name
2. Logs are append-only (cannot be edited without detection)
3. Stored locally at: `%LOCALAPPDATA%\Domino\Logs`
4. 30-day retention (configurable)

**Collecting logs for audit:**
```powershell
# Collect user's logs
$user = "john.doe"
$logPath = "C:\Users\$user\AppData\Local\Domino\Logs"
Copy-Item $logPath "\\fileserver\Audit\Domino\$user" -Recurse
```

## FAQ

**Q: Does this slow down Excel?**
A: No. The add-in uses event-driven architecture with minimal overhead. Typical performance impact < 1%.

**Q: Can I track more than A1:D4?**
A: Yes, but requires code modification. Contact your IT department or see README.md.

**Q: What if I work offline?**
A: Tracking works offline. Logs are stored locally regardless of network connectivity.

**Q: Can my manager see my changes?**
A: Only if they have access to your computer or if IT has set up centralized log collection.

**Q: Does this work with Excel Online?**
A: No. Excel-DNA add-ins only work with Excel Desktop for Windows.

**Q: What about macOS Excel?**
A: Not supported. This is a Windows-only add-in.

**Q: Can I uninstall it?**
A: Yes. File → Options → Add-ins → Uncheck "Domino Add-In" or click "Remove".

**Q: Will it track my entire workbook?**
A: No. Only cells A1:D4 in all sheets. Other cells are not tracked.

---

**Need more details?** See [README.md](README.md) for comprehensive documentation.

**Enterprise deployment?** See [DEPLOYMENT.md](DEPLOYMENT.md) for IT administrators.
