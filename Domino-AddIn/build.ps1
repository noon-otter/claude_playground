# Domino Excel Add-In Build Script
# For Windows development and deployment

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',

    [Parameter(Mandatory=$false)]
    [switch]$Clean,

    [Parameter(Mandatory=$false)]
    [switch]$Sign,

    [Parameter(Mandatory=$false)]
    [string]$CertificatePath,

    [Parameter(Mandatory=$false)]
    [string]$CertificatePassword
)

$ErrorActionPreference = "Stop"

# Script configuration
$ProjectName = "Domino"
$ProjectFile = "Domino.csproj"
$Platform = "x64"

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Domino Excel Add-In Build Script" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Check prerequisites
Write-Host "Checking prerequisites..." -ForegroundColor Yellow

# Check if .NET SDK is installed
try {
    $dotnetVersion = dotnet --version
    Write-Host "✓ .NET SDK version: $dotnetVersion" -ForegroundColor Green
} catch {
    Write-Host "✗ .NET SDK not found. Please install .NET 6.0 SDK or later." -ForegroundColor Red
    exit 1
}

# Check if project file exists
if (-not (Test-Path $ProjectFile)) {
    Write-Host "✗ Project file not found: $ProjectFile" -ForegroundColor Red
    exit 1
}
Write-Host "✓ Project file found: $ProjectFile" -ForegroundColor Green

# Clean if requested
if ($Clean) {
    Write-Host ""
    Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
    dotnet clean -c $Configuration
    if ($LASTEXITCODE -ne 0) {
        Write-Host "✗ Clean failed" -ForegroundColor Red
        exit 1
    }
    Write-Host "✓ Clean completed" -ForegroundColor Green
}

# Restore NuGet packages
Write-Host ""
Write-Host "Restoring NuGet packages..." -ForegroundColor Yellow
dotnet restore
if ($LASTEXITCODE -ne 0) {
    Write-Host "✗ Package restore failed" -ForegroundColor Red
    exit 1
}
Write-Host "✓ Packages restored" -ForegroundColor Green

# Build the project
Write-Host ""
Write-Host "Building project ($Configuration configuration, $Platform platform)..." -ForegroundColor Yellow
dotnet build -c $Configuration -p:Platform=$Platform --no-restore
if ($LASTEXITCODE -ne 0) {
    Write-Host "✗ Build failed" -ForegroundColor Red
    exit 1
}
Write-Host "✓ Build succeeded" -ForegroundColor Green

# Locate output files
$OutputDir = "bin\$Configuration\net6.0-windows"
$XllFile = Join-Path $OutputDir "$ProjectName-AddIn64.xll"

if (-not (Test-Path $XllFile)) {
    Write-Host "✗ Output file not found: $XllFile" -ForegroundColor Red
    Write-Host "Looking for files in output directory:" -ForegroundColor Yellow
    Get-ChildItem $OutputDir -Recurse | ForEach-Object { Write-Host "  $($_.FullName)" }
    exit 1
}

Write-Host ""
Write-Host "Build Output:" -ForegroundColor Cyan
Write-Host "  XLL File: $XllFile" -ForegroundColor White
Write-Host "  Size: $((Get-Item $XllFile).Length / 1KB) KB" -ForegroundColor White

# Code signing (if requested)
if ($Sign) {
    Write-Host ""
    Write-Host "Code signing..." -ForegroundColor Yellow

    if (-not $CertificatePath -or -not (Test-Path $CertificatePath)) {
        Write-Host "✗ Certificate file not found: $CertificatePath" -ForegroundColor Red
        Write-Host "Usage: .\build.ps1 -Sign -CertificatePath path\to\cert.pfx -CertificatePassword 'password'" -ForegroundColor Yellow
        exit 1
    }

    # Check if signtool is available
    $signTool = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if (-not $signTool) {
        Write-Host "✗ signtool.exe not found. Install Windows SDK." -ForegroundColor Red
        exit 1
    }

    $signArgs = @(
        'sign',
        '/f', $CertificatePath,
        '/p', $CertificatePassword,
        '/tr', 'http://timestamp.digicert.com',
        '/td', 'sha256',
        '/fd', 'sha256',
        '/d', 'Domino Excel Add-In',
        $XllFile
    )

    & signtool.exe $signArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Host "✗ Code signing failed" -ForegroundColor Red
        exit 1
    }

    # Verify signature
    & signtool.exe verify /pa $XllFile
    if ($LASTEXITCODE -ne 0) {
        Write-Host "✗ Signature verification failed" -ForegroundColor Red
        exit 1
    }

    Write-Host "✓ Code signed successfully" -ForegroundColor Green
}

# Build summary
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Build Complete!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Output file: $XllFile" -ForegroundColor White
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Test the add-in:" -ForegroundColor White
Write-Host "     - Open Excel" -ForegroundColor Gray
Write-Host "     - File → Options → Add-ins → Browse" -ForegroundColor Gray
Write-Host "     - Select: $XllFile" -ForegroundColor Gray
Write-Host ""
Write-Host "  2. Check logs at:" -ForegroundColor White
Write-Host "     %LOCALAPPDATA%\Domino\Logs\" -ForegroundColor Gray
Write-Host ""
Write-Host "  3. Deploy to users (see DEPLOYMENT.md)" -ForegroundColor White
Write-Host ""

# Optional: Copy to a deployment folder
$DeployDir = ".\Deploy"
if (-not (Test-Path $DeployDir)) {
    New-Item -Path $DeployDir -ItemType Directory | Out-Null
}

Copy-Item $XllFile $DeployDir -Force
Write-Host "✓ Copied to deployment folder: $DeployDir" -ForegroundColor Green
Write-Host ""
