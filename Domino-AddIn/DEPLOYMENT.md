# Domino Add-In - Enterprise Deployment Guide

This guide covers deployment of the Domino Excel Add-In in financial services and enterprise environments with strict security requirements.

## Deployment Overview

### Distribution Formats
1. **XLL File** (Recommended): Single-file deployment
2. **MSI Installer**: For managed deployments via Group Policy
3. **VSTO Deployment**: For Office 365 environments

## Pre-Deployment Checklist

### Security Review
- [ ] Complete source code security audit
- [ ] Obtain code signing certificate from approved CA
- [ ] Sign all binaries with company certificate
- [ ] Document all external dependencies
- [ ] Review and approve logging data collected
- [ ] Configure log file encryption (if required)

### IT Infrastructure
- [ ] Verify .NET 6.0 Runtime installed on target machines
- [ ] Confirm Excel 2016+ Desktop version (not Office 365 Online)
- [ ] Test on representative user workstations
- [ ] Verify local storage permissions (%LOCALAPPDATA%)
- [ ] Configure antivirus/EDR exclusions if needed

### Compliance
- [ ] Data residency requirements met (logs stored locally)
- [ ] Audit trail requirements reviewed
- [ ] Change management approval obtained
- [ ] User notification/training completed
- [ ] Privacy impact assessment (if applicable)

## Build Process for Production

### 1. Clean Build

```bash
# Navigate to project directory
cd Domino-AddIn

# Clean previous builds
dotnet clean -c Release

# Restore packages
dotnet restore

# Build for release
dotnet build -c Release -p:Platform=x64
```

### 2. Code Signing

#### Using SignTool (Windows SDK)

```powershell
# Sign the XLL file
signtool sign `
  /f "path\to\certificate.pfx" `
  /p "certificate-password" `
  /tr http://timestamp.digicert.com `
  /td sha256 `
  /fd sha256 `
  /d "Domino Excel Add-In" `
  /du "https://your-company.com" `
  bin\Release\net6.0-windows\Domino-AddIn64.xll

# Verify signature
signtool verify /pa bin\Release\net6.0-windows\Domino-AddIn64.xll
```

#### Using Azure Key Vault (Cloud HSM)

```powershell
# Install AzureSignTool
dotnet tool install --global AzureSignTool

# Sign with Azure Key Vault
azuresigntool sign `
  -kvu "https://your-vault.vault.azure.net/" `
  -kvi "client-id" `
  -kvs "client-secret" `
  -kvc "cert-name" `
  -tr http://timestamp.digicert.com `
  -td sha256 `
  bin\Release\net6.0-windows\Domino-AddIn64.xll
```

### 3. Verification

```powershell
# Verify the signature
Get-AuthenticodeSignature bin\Release\net6.0-windows\Domino-AddIn64.xll | Format-List

# Expected output:
# SignerCertificate : [Your Certificate]
# Status            : Valid
```

## Deployment Methods

### Method 1: Manual Installation (Pilot Testing)

Best for initial rollout to pilot users.

**Step 1: Distribute the file**
```
Share via:
- Network share: \\fileserver\Deployments\Domino\Domino-AddIn64.xll
- Email (if policy allows)
- USB drive
```

**Step 2: User installation**
1. Copy `Domino-AddIn64.xll` to a permanent location:
   ```
   C:\Program Files\Domino-AddIn\Domino-AddIn64.xll
   ```

2. Open Excel → File → Options → Add-ins

3. Select "Excel Add-ins" from dropdown → Go

4. Click Browse → Select the .xll file

5. Check "Domino Add-In" → OK

**Step 3: Verify installation**
- Check for "Domino" tab in Excel ribbon
- Open log directory: `%LOCALAPPDATA%\Domino\Logs`
- Make a test change in A1:D4 range
- Verify log entry created

### Method 2: Group Policy Deployment (Recommended for Enterprise)

Best for organization-wide deployment.

#### Option A: Excel-DNA Auto-Load

Create a Group Policy to copy the .xll file to Excel's startup directory:

**GPO Configuration**:
```
Computer Configuration → Preferences → Windows Settings → Files

Source: \\fileserver\Deployments\Domino\Domino-AddIn64.xll
Destination: C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Excel\XLSTART\Domino-AddIn64.xll
```

Excel automatically loads all .xll files in XLSTART on startup.

#### Option B: Registry-Based Auto-Load

**Registry Key**:
```
HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Options

Name: OPEN
Type: REG_SZ
Value: /R "C:\Program Files\Domino-AddIn\Domino-AddIn64.xll"
```

**Group Policy Procedure**:
1. Open Group Policy Management Console
2. Create new GPO: "Domino Excel Add-In Deployment"
3. Edit GPO → User Configuration → Preferences → Windows Settings → Registry
4. New → Registry Item:
   - Action: Update
   - Hive: HKEY_CURRENT_USER
   - Key Path: Software\Microsoft\Office\16.0\Excel\Options
   - Value name: OPEN
   - Value type: REG_SZ
   - Value data: /R "C:\Program Files\Domino-AddIn\Domino-AddIn64.xll"

5. Link GPO to target OU

### Method 3: MSI Installer

For organizations requiring formal installers.

#### Create MSI with WiX Toolset

**install.wxs**:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="*"
           Name="Domino Excel Add-In"
           Language="1033"
           Version="1.0.0"
           Manufacturer="Your Organization"
           UpgradeCode="PUT-GUID-HERE">

    <Package InstallerVersion="200" Compressed="yes" InstallScope="perMachine" Platform="x64" />

    <MajorUpgrade DowngradeErrorMessage="A newer version is already installed." />

    <MediaTemplate EmbedCab="yes" />

    <Feature Id="ProductFeature" Title="Domino Add-In" Level="1">
      <ComponentGroupRef Id="ProductComponents" />
    </Feature>

    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="ProgramFiles64Folder">
        <Directory Id="INSTALLFOLDER" Name="Domino-AddIn" />
      </Directory>
    </Directory>

    <ComponentGroup Id="ProductComponents" Directory="INSTALLFOLDER">
      <Component Id="DominoXLL" Guid="PUT-GUID-HERE" Win64="yes">
        <File Id="DominoXLL" Source="bin\Release\net6.0-windows\Domino-AddIn64.xll" KeyPath="yes" />

        <!-- Auto-load registry key -->
        <RegistryValue Root="HKCU"
                       Key="Software\Microsoft\Office\16.0\Excel\Options"
                       Name="OPEN"
                       Type="string"
                       Value='/R "[INSTALLFOLDER]Domino-AddIn64.xll"' />
      </Component>

      <Component Id="NLogConfig" Guid="PUT-GUID-HERE" Win64="yes">
        <File Id="NLogConfig" Source="NLog.config" KeyPath="yes" />
      </Component>
    </ComponentGroup>
  </Product>
</Wix>
```

**Build MSI**:
```bash
# Install WiX Toolset
# Download from: https://wixtoolset.org/

# Compile
candle install.wxs
light -out Domino-AddIn.msi install.wixobj

# Sign the MSI
signtool sign /f certificate.pfx /p password /t http://timestamp.digicert.com Domino-AddIn.msi
```

**Deploy MSI via Group Policy**:
1. Copy MSI to network share
2. Group Policy Management → Create new GPO
3. Computer Configuration → Policies → Software Settings → Software Installation
4. New → Package → Select Domino-AddIn.msi
5. Link to target OU

### Method 4: SCCM/MECM Deployment

For Microsoft Endpoint Configuration Manager environments.

**Package Configuration**:
```
Content Source: \\fileserver\Deployments\Domino\
Program: msiexec /i Domino-AddIn.msi /quiet /norestart
Detection Method: File exists - C:\Program Files\Domino-AddIn\Domino-AddIn64.xll
Deployment Type: Required
Target Collection: "Finance Department Workstations"
```

## Configuration Management

### Centralized Configuration (Optional)

For standardized settings across organization:

**Option 1: Group Policy Registry**
```
HKEY_LOCAL_MACHINE\SOFTWARE\Domino\AddIn

LogLevel (REG_SZ): Info|Debug|Warn|Error
MonitoredRange (REG_SZ): A1:D4
EnableRibbon (REG_DWORD): 1
```

**Option 2: Configuration File**
Place in: `%PROGRAMDATA%\Domino\config.json`
```json
{
  "monitoredRange": "A1:D4",
  "logLevel": "Info",
  "enableRibbon": true,
  "apiEndpoint": "https://api.your-company.com/tracking"
}
```

## Monitoring and Maintenance

### Health Checks

**PowerShell Script** (`Check-DominoHealth.ps1`):
```powershell
# Check if add-in is loaded
$excelProcess = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcess) {
    $modules = Get-Process excel | Select-Object -ExpandProperty Modules | Where-Object { $_.ModuleName -like "*Domino*" }
    if ($modules) {
        Write-Host "✓ Domino add-in loaded" -ForegroundColor Green
    } else {
        Write-Host "✗ Domino add-in NOT loaded" -ForegroundColor Red
    }
}

# Check log directory
$logPath = "$env:LOCALAPPDATA\Domino\Logs"
if (Test-Path $logPath) {
    $logFiles = Get-ChildItem $logPath -Filter "domino-*.log"
    $latestLog = $logFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($latestLog -and $latestLog.LastWriteTime -gt (Get-Date).AddHours(-24)) {
        Write-Host "✓ Recent log activity: $($latestLog.LastWriteTime)" -ForegroundColor Green
    } else {
        Write-Host "⚠ No recent log activity" -ForegroundColor Yellow
    }
} else {
    Write-Host "✗ Log directory not found" -ForegroundColor Red
}
```

### Log Collection for Support

**Collect Diagnostic Data**:
```powershell
# Script to collect logs for troubleshooting
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outputPath = "C:\Temp\Domino-Diagnostics-$timestamp"

New-Item -Path $outputPath -ItemType Directory

# Copy logs
Copy-Item "$env:LOCALAPPDATA\Domino\Logs\*" $outputPath -Recurse

# Export installed add-ins
Get-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" | Out-File "$outputPath\excel-options.txt"

# System info
systeminfo | Out-File "$outputPath\systeminfo.txt"

# Compress
Compress-Archive -Path $outputPath -DestinationPath "$outputPath.zip"
```

## Rollback Procedure

### Uninstall via Registry

```powershell
# Remove auto-load registry key
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" -Name "OPEN"
```

### Uninstall via Excel UI

1. File → Options → Add-ins
2. Excel Add-ins → Go
3. Uncheck "Domino Add-In"
4. Remove...

### Clean Uninstall Script

```powershell
# Stop Excel
Stop-Process -Name excel -Force -ErrorAction SilentlyContinue

# Remove add-in file
Remove-Item "C:\Program Files\Domino-AddIn\*" -Recurse -Force

# Remove registry
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" -Name "OPEN" -ErrorAction SilentlyContinue

# Optionally remove logs
# Remove-Item "$env:LOCALAPPDATA\Domino" -Recurse -Force
```

## Security Hardening

### File System Permissions

Restrict write access to installation directory:
```powershell
$acl = Get-Acl "C:\Program Files\Domino-AddIn"
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "Users", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow"
)
$acl.SetAccessRule($rule)
Set-Acl "C:\Program Files\Domino-AddIn" $acl
```

### Application Whitelisting

**AppLocker Policy**:
```xml
<FilePublisherRule Id="..." Name="Domino Excel Add-In"
                   UserOrGroupSid="S-1-1-0" Action="Allow">
  <Conditions>
    <FilePublisherCondition PublisherName="O=Your Organization"
                           ProductName="Domino Excel Add-In"
                           BinaryName="Domino-AddIn64.xll">
      <BinaryVersionRange LowSection="1.0.0.0" HighSection="*" />
    </FilePublisherCondition>
  </Conditions>
</FilePublisherRule>
```

## Troubleshooting Deployment Issues

### Issue: Add-in not loading

**Diagnostics**:
```powershell
# Check .NET Runtime
dotnet --list-runtimes | Select-String "WindowsDesktop.App 6"

# Check file permissions
Get-Acl "C:\Program Files\Domino-AddIn\Domino-AddIn64.xll" | Format-List

# Check signature
Get-AuthenticodeSignature "C:\Program Files\Domino-AddIn\Domino-AddIn64.xll"

# Check Excel version
(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name VersionToReport).VersionToReport
```

### Issue: Trust Center blocking add-in

**Resolution via Group Policy**:
```
User Configuration → Administrative Templates → Microsoft Excel 2016 →
Application Settings → Security → Trust Center

"Require that application add-ins are signed by Trusted Publisher" → Enabled
```

Then add your code signing certificate to Trusted Publishers.

## Performance Monitoring

### Key Metrics

Monitor these via Performance Monitor (perfmon):
- Process: excel.exe
  - Private Bytes (memory usage)
  - % Processor Time
- Object: .NET CLR Memory
  - # Bytes in all Heaps

### Expected Performance

| Metric | Typical Value | Alert Threshold |
|--------|---------------|-----------------|
| Add-in load time | < 2 seconds | > 5 seconds |
| Memory usage | 5-10 MB | > 50 MB |
| CPU usage (idle) | < 1% | > 5% |
| Log file size/day | 1-5 MB | > 100 MB |

## Support Escalation

### Level 1: User Support
- Check if tracking is active (ribbon display)
- Verify log files are being created
- Restart Excel

### Level 2: IT Support
- Run diagnostic script
- Check event logs
- Verify .NET runtime
- Collect logs

### Level 3: Development Team
- Source code review
- Memory dump analysis
- Performance profiling
- Security audit

---

**Document Version**: 1.0
**Last Updated**: 2025-11-22
**Review Cycle**: Quarterly
