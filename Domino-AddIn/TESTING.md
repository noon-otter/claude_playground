# Domino Add-In - Testing Guide

Comprehensive testing procedures for the Domino Excel cell tracking add-in.

## Testing Environment Setup

### Minimum Test Environment
- Windows 10 (version 1809 or later)
- Excel 2016 Desktop
- .NET 6.0 Runtime
- 4GB RAM
- 100MB free disk space

### Recommended Test Environment
- Windows 11
- Excel 2021 or Microsoft 365 Desktop
- .NET 6.0 or later
- 8GB+ RAM
- 1GB free disk space for logs

## Pre-Test Checklist

- [ ] Build completed successfully (no errors)
- [ ] NLog.config file present in output directory
- [ ] .xll file is not blocked (Right-click → Properties → Unblock)
- [ ] Antivirus/EDR temporarily disabled or add-in whitelisted
- [ ] No other Excel instances running
- [ ] Administrator privileges available (if needed)

## Unit Testing

### Test 1: Build and Output Validation

**Objective**: Verify build process produces correct outputs

```powershell
# Build the project
.\build.ps1 -Configuration Release

# Verify outputs
Test-Path "bin\Release\net6.0-windows\Domino-AddIn64.xll"  # Should be True
Test-Path "bin\Release\net6.0-windows\NLog.config"         # Should be True
```

**Expected Results**:
- ✅ Build completes without errors
- ✅ .xll file is created
- ✅ .xll file size > 100KB
- ✅ NLog.config is copied to output directory

**Pass Criteria**: All files present and no build errors

---

### Test 2: Add-In Load

**Objective**: Verify add-in loads in Excel without errors

**Steps**:
1. Open Excel
2. File → Options → Add-ins
3. Manage: Excel Add-ins → Go
4. Browse → Select `Domino-AddIn64.xll`
5. Check "Domino Add-In" → OK

**Expected Results**:
- ✅ Add-in appears in list as "Domino Add-In"
- ✅ No error dialogs appear
- ✅ Excel remains responsive
- ✅ "Domino" ribbon tab appears

**Validation**:
```powershell
# Check log file created
Test-Path "$env:LOCALAPPDATA\Domino\Logs\domino-$(Get-Date -Format yyyy-MM-dd).log"
```

**Pass Criteria**: Add-in loads successfully and creates log file

---

## Functional Testing

### Test 3: Value Change Tracking

**Objective**: Verify value changes in A1:D4 are tracked

**Test Cases**:

| Cell | Action | Expected Log Entry |
|------|--------|-------------------|
| A1 | Enter "100" | `[VALUE_CHANGE] ... Cell: A1 ... → 100` |
| B2 | Enter "Test" | `[VALUE_CHANGE] ... Cell: B2 ... → Test` |
| C3 | Enter "5.5" | `[VALUE_CHANGE] ... Cell: C3 ... → 5.5` |
| D4 | Enter "=TODAY()" | `[FORMULA_CHANGE] ... Cell: D4 ... → =TODAY()` |

**Steps**:
1. Open Excel with add-in loaded
2. Create new workbook
3. Enter values as per test cases
4. Open log file: `%LOCALAPPDATA%\Domino\Logs\domino-YYYY-MM-DD.log`
5. Verify each change is logged

**Pass Criteria**: All changes logged correctly with timestamp and cell address

---

### Test 4: Formula Change Tracking

**Objective**: Verify formula changes are tracked

**Test Cases**:
1. Enter `=SUM(A1:A10)` in A1 → Should log FORMULA_CHANGE
2. Enter `=IF(B1>0,"Yes","No")` in B1 → Should log FORMULA_CHANGE
3. Enter `=VLOOKUP(C1,$A$1:$B$10,2,FALSE)` in C1 → Should log FORMULA_CHANGE

**Steps**:
1. Enter each formula in cells A1-C1
2. Check log file
3. Verify formula text is captured correctly

**Pass Criteria**: All formulas logged with full formula text

---

### Test 5: Out-of-Range Changes (Negative Test)

**Objective**: Verify cells outside A1:D4 are NOT tracked

**Test Cases**:
| Cell | Action | Expected |
|------|--------|----------|
| E5 | Enter "Should not log" | No log entry |
| Z10 | Enter "999" | No log entry |
| A5 | Enter "100" | No log entry (outside range) |

**Steps**:
1. Note current log file size
2. Enter values in cells outside A1:D4
3. Wait 5 seconds
4. Check log file hasn't changed

**Pass Criteria**: No log entries for cells outside monitored range

---

### Test 6: Workbook Open/Close Tracking

**Objective**: Verify workbook lifecycle events are tracked

**Steps**:
1. Close all workbooks
2. Open a new workbook (Test1.xlsx)
3. Check log for `[WORKBOOK_OPEN] Test1.xlsx`
4. Close Test1.xlsx
5. Check log for `[WORKBOOK_CLOSE] Test1.xlsx`

**Expected Log**:
```
[WORKBOOK_OPEN] Workbook: Test1.xlsx
[WORKBOOK_CLOSE] Workbook: Test1.xlsx
```

**Pass Criteria**: Both open and close events logged correctly

---

### Test 7: Multi-Sheet Tracking

**Objective**: Verify tracking works across multiple sheets

**Steps**:
1. Create workbook with 3 sheets (Sheet1, Sheet2, Sheet3)
2. In Sheet1, enter "A" in A1
3. In Sheet2, enter "B" in B2
4. In Sheet3, enter "C" in C3
5. Check log file

**Expected Results**:
```
[VALUE_CHANGE] Workbook: Book1 | Sheet: Sheet1 | Cell: A1 | → A
[VALUE_CHANGE] Workbook: Book1 | Sheet: Sheet2 | Cell: B2 | → B
[VALUE_CHANGE] Workbook: Book1 | Sheet: Sheet3 | Cell: C3 | → C
```

**Pass Criteria**: All sheets tracked independently with correct sheet names

---

### Test 8: Multi-Workbook Tracking

**Objective**: Verify simultaneous tracking of multiple workbooks

**Steps**:
1. Open Workbook1.xlsx
2. Open Workbook2.xlsx (keep Workbook1 open)
3. In Workbook1, change A1
4. In Workbook2, change B2
5. Check log file

**Expected Results**:
- Changes in both workbooks logged
- Correct workbook name in each log entry
- No cross-contamination of events

**Pass Criteria**: Both workbooks tracked correctly and independently

---

### Test 9: Ribbon UI Functionality

**Objective**: Verify ribbon controls work correctly

| Control | Test Action | Expected Result |
|---------|-------------|-----------------|
| Last Change Timestamp | Make change in A1 | Timestamp updates (may need refresh) |
| Refresh Button | Click Refresh | Timestamp updates immediately |
| View Logs Button | Click View Logs | Log file opens in default editor |
| About Button | Click About | About dialog displays version info |

**Pass Criteria**: All buttons functional, no errors

---

## Performance Testing

### Test 10: Rapid Change Performance

**Objective**: Verify add-in handles rapid changes without lag

**Steps**:
1. Prepare script to enter values in A1:D4 rapidly
2. Execute script (or manually enter values quickly)
3. Monitor Excel responsiveness
4. Check all changes logged

**VBA Test Script**:
```vba
Sub TestRapidChanges()
    Application.ScreenUpdating = False
    Dim i As Long
    For i = 1 To 100
        Range("A1").Value = i
        Range("B2").Value = i * 2
        Range("C3").Value = i * 3
        Range("D4").Value = i * 4
    Next i
    Application.ScreenUpdating = True
End Sub
```

**Expected Results**:
- Excel remains responsive
- All (or most) changes logged
- No Excel crashes

**Pass Criteria**: <2 second delay for 100 changes, no crashes

---

### Test 11: Long-Running Session

**Objective**: Verify stability over extended use

**Steps**:
1. Load add-in
2. Keep Excel open for 8+ hours
3. Periodically make changes in A1:D4
4. Check for memory leaks
5. Verify all changes logged

**Monitoring**:
```powershell
# Monitor Excel memory usage
while ($true) {
    $excel = Get-Process excel -ErrorAction SilentlyContinue
    if ($excel) {
        $memMB = [math]::Round($excel.WorkingSet64 / 1MB, 2)
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Excel Memory: $memMB MB"
    }
    Start-Sleep -Seconds 300  # Check every 5 minutes
}
```

**Pass Criteria**: Memory usage stable (<50MB growth over 8 hours)

---

## Integration Testing

### Test 12: NLog Integration

**Objective**: Verify logging system works correctly

**Tests**:

**A. Log File Creation**
```powershell
# Expected path
$logPath = "$env:LOCALAPPDATA\Domino\Logs\domino-$(Get-Date -Format yyyy-MM-dd).log"
Test-Path $logPath  # Should be True
```

**B. Log Format**
- Verify timestamp format: `YYYY-MM-DD HH:mm:ss.ffff`
- Verify log level: `INFO`
- Verify logger name: `Domino.ChangeTracker`
- Verify message format

**C. Log Rotation**
```powershell
# Change system date to tomorrow (admin required)
# Or wait 24 hours
# Verify new log file created for new date
```

**D. Console Output** (if Excel launched from terminal)
- Verify console shows colored output
- Verify console messages match file log

**Pass Criteria**: All log mechanisms functional

---

### Test 13: Excel Interop Compatibility

**Objective**: Test compatibility with different Excel versions

**Test Matrix**:

| Excel Version | Expected Support |
|---------------|------------------|
| Excel 2016 | ✅ Fully supported |
| Excel 2019 | ✅ Fully supported |
| Excel 2021 | ✅ Fully supported |
| Microsoft 365 | ✅ Fully supported |
| Excel 2013 | ⚠️  May work (not officially supported) |

**Steps for each version**:
1. Install add-in
2. Run Test 3-8
3. Document any issues

**Pass Criteria**: No issues on Excel 2016+

---

## Security Testing

### Test 14: Code Signing Verification

**Objective**: Verify digitally signed add-in loads without warnings

**Steps**:
1. Sign .xll file with certificate
2. Load in Excel with strict macro security
3. Verify no security warnings

```powershell
# Check signature
Get-AuthenticodeSignature "Domino-AddIn64.xll" | Format-List
```

**Expected**:
- Status: Valid
- SignerCertificate: [Your Certificate]
- No security warnings in Excel

**Pass Criteria**: Signed add-in loads without security prompts

---

### Test 15: Sandboxing and Permissions

**Objective**: Verify add-in respects user permissions

**Tests**:

**A. Read-Only Log Directory**
1. Make log directory read-only
2. Load add-in
3. Make changes
4. Verify graceful failure (no crash)

**B. No Local App Data Access**
1. Deny write access to %LOCALAPPDATA%
2. Load add-in
3. Verify error logged (if alternate location available)

**Pass Criteria**: Add-in doesn't crash on permission errors

---

## Regression Testing

### Test 16: Upgrade Compatibility

**Objective**: Verify new version doesn't break existing functionality

**Steps**:
1. Install version 1.0
2. Create test workbook
3. Make changes and verify logging
4. Upgrade to version 1.1
5. Open same workbook
6. Verify tracking still works
7. Verify old logs still readable

**Pass Criteria**: No functionality regression

---

## User Acceptance Testing (UAT)

### Test 17: End-User Scenarios

**Scenario 1: Budget Analyst**
- User: Finance team member
- Task: Track budget changes in quarterly reports
- Workbooks: 3-5 simultaneous workbooks
- Expected: All changes in A1:D4 tracked accurately

**Scenario 2: Compliance Officer**
- User: Audit team member
- Task: Review change logs for compliance
- Action: Open log files, search for specific changes
- Expected: Log files easy to read and search

**Scenario 3: Power User**
- User: Heavy Excel user with many add-ins
- Task: Use Excel normally with Domino running
- Expected: No performance degradation, no conflicts

**Pass Criteria**: All users can complete tasks without issues

---

## Test Reporting

### Test Summary Template

```markdown
# Test Run Summary

**Date**: YYYY-MM-DD
**Tester**: Name
**Environment**: Windows 11, Excel 2021, .NET 6.0
**Build Version**: 1.0.0

## Results

| Test ID | Test Name | Status | Notes |
|---------|-----------|--------|-------|
| Test 1 | Build Validation | ✅ Pass | |
| Test 2 | Add-In Load | ✅ Pass | |
| Test 3 | Value Change Tracking | ✅ Pass | |
| Test 4 | Formula Tracking | ⚠️  Partial | Long formulas truncated |
| Test 5 | Out-of-Range | ✅ Pass | |
| ... | ... | ... | ... |

## Issues Found

1. **Issue #1**: Long formulas (>255 chars) truncated in log
   - Severity: Low
   - Workaround: None needed

## Overall Result

- Total Tests: 17
- Passed: 16
- Failed: 0
- Partial: 1

**Recommendation**: Ready for deployment
```

---

## Automated Testing (Future)

### Automated Test Suite (Planned)

```csharp
// Example: NUnit test framework
[TestFixture]
public class ChangeTrackerTests
{
    [Test]
    public void TestValueChange_InMonitoredRange_ShouldLog()
    {
        // Arrange
        var tracker = new ChangeTracker(mockExcelApp);

        // Act
        SimulateCellChange("A1", "100");

        // Assert
        var logs = ReadLogFile();
        Assert.Contains("[VALUE_CHANGE]", logs);
        Assert.Contains("Cell: A1", logs);
        Assert.Contains("→ 100", logs);
    }
}
```

---

## Troubleshooting Test Failures

### Common Issues

**Add-in doesn't load**
- Check .NET runtime installed
- Check file not blocked
- Check Excel Trust Center settings

**Changes not logged**
- Verify NLog.config in output directory
- Check log directory permissions
- Verify changing cells in A1:D4 range

**Ribbon doesn't appear**
- Check Excel version (2007+)
- Reset Excel customizations
- Check .dna file ribbon XML

**Performance issues**
- Check other add-ins for conflicts
- Monitor system resources
- Check log file size

---

**Testing Checklist**: Run all tests before each release
**Sign-off Required**: QA Lead, Product Owner, Security Team
