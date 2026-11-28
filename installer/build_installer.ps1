# Builds dist\installer\CMTS_Command_Center_Setup_<version>.exe
#   1. installs the Python build dependencies
#   2. runs PyInstaller (folder build in dist\CMTS_Command_Center)
#   3. runs Inno Setup's compiler on installer\CMTS_Command_Center.iss
#
# Run from anywhere:  powershell -ExecutionPolicy Bypass -File installer\build_installer.ps1
# Or double-click build_installer.bat in the repo root.

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$match = Select-String -Path 'CMTS_Command_Center.py' -Pattern '^VERSION = "([^"]+)"' | Select-Object -First 1
if (-not $match) { throw 'Could not find VERSION = "x.y.z" in CMTS_Command_Center.py' }
$Version = $match.Matches[0].Groups[1].Value
Write-Host "== CMTS Command Center $Version ==" -ForegroundColor Cyan

# Find the Inno Setup compiler before spending time on the build
$Iscc = (Get-Command 'ISCC.exe' -ErrorAction SilentlyContinue).Source
if (-not $Iscc) {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    )
    $Iscc = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}
if (-not $Iscc) {
    throw "Inno Setup 6 not found. Install it with:  winget install JRSoftware.InnoSetup   (or https://jrsoftware.org/isdl.php)"
}

Write-Host "`n[1/3] Installing build dependencies..." -ForegroundColor Cyan
python -m pip install --upgrade pyinstaller paramiko customtkinter matplotlib openpyxl
if ($LASTEXITCODE -ne 0) { throw 'pip install failed' }

Write-Host "`n[2/3] Building the app with PyInstaller..." -ForegroundColor Cyan
python -m PyInstaller --noconfirm --clean CMTS_Command_Center.spec
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller failed' }
if (-not (Test-Path 'dist\CMTS_Command_Center\CMTS_Command_Center.exe')) { throw 'PyInstaller output not found' }

Write-Host "`n[3/3] Building the installer with Inno Setup..." -ForegroundColor Cyan
& $Iscc "/DMyAppVersion=$Version" 'installer\CMTS_Command_Center.iss'
if ($LASTEXITCODE -ne 0) { throw 'Inno Setup failed' }

$Out = Join-Path $Root "dist\installer\CMTS_Command_Center_Setup_$Version.exe"
Write-Host "`nDone: $Out" -ForegroundColor Green
