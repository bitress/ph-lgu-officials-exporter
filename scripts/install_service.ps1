# install_service.ps1
# Sets up the PH LGU Officials web server to auto-start on Windows logon.
# Uses HKCU\Run registry key — NO Administrator rights required.
# Run once: powershell -ExecutionPolicy Bypass -File install_service.ps1
#
# To uninstall: powershell -ExecutionPolicy Bypass -File install_service.ps1 -Uninstall

param(
    [string]$Port = "8000",
    [switch]$Uninstall
)

$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$AppDir      = Split-Path -Parent $ScriptDir   # project root (this script lives in scripts\)
$VbsLauncher = Join-Path $ScriptDir "start_web_server.vbs"
$LogDir      = Join-Path $AppDir "data\logs"
$RegPath     = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$RegName     = "PHLGUOfficials"

# --- Uninstall ---
if ($Uninstall) {
    Remove-ItemProperty -Path $RegPath -Name $RegName -ErrorAction SilentlyContinue
    Write-Host "[OK] Auto-start entry removed." -ForegroundColor Green

    # Also stop any running pythonw processes for this app
    Get-Process -Name "pythonw", "python", "wscript" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*ph-lgu*" -or $_.CommandLine -like "*start_web_server*" } |
        Stop-Process -Force -ErrorAction SilentlyContinue

    Write-Host "[OK] Server processes stopped." -ForegroundColor Green
    exit 0
}

# --- Banner ---
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  PH LGU Officials Web Server - Auto-Start Installer" -ForegroundColor Cyan
Write-Host "  Port: $Port  |  No Administrator required" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# --- Pre-flight: Python check ---
try {
    $pyVer = & python --version 2>&1
    Write-Host "[OK] Python found: $pyVer" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Python not found in PATH." -ForegroundColor Red
    exit 1
}

# --- Pre-flight: openpyxl + docxtpl ---
$ok = & python -c "import openpyxl, docxtpl, docxcompose; print('ok')" 2>&1
if ($ok -ne "ok") {
    Write-Host "[INFO] Installing openpyxl + docxtpl..." -ForegroundColor Yellow
    & pip install -r (Join-Path $AppDir "requirements.txt") -q
}

# --- Ensure logs directory exists ---
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    Write-Host "[OK] Created logs directory." -ForegroundColor Green
}

# --- Patch VBS with correct port and pythonw path ---
$pythonExe  = (& python -c "import sys,os; print(os.path.join(os.path.dirname(sys.executable),'pythonw.exe'))").Trim()
if (-not (Test-Path $pythonExe)) { $pythonExe = "pythonw" }

$vbsContent = Get-Content $VbsLauncher -Raw
$vbsContent  = $vbsContent -replace '--web \d+', "--web $Port"
# Update the hardcoded pythonw path in VBS
$vbsContent  = $vbsContent -replace 'pythonwExe = "C:\\.*?pythonw\.exe"', "pythonwExe = `"$pythonExe`""
Set-Content -Path $VbsLauncher -Value $vbsContent -NoNewline
Write-Host "[OK] Configured web server port: $Port" -ForegroundColor Green
Write-Host "[OK] Using Python: $pythonExe" -ForegroundColor Green


# --- Set HKCU Run key (no Admin needed) ---
$cmdValue = "wscript.exe `"$VbsLauncher`""
Set-ItemProperty -Path $RegPath -Name $RegName -Value $cmdValue -Force
Write-Host "[OK] Auto-start registry key set:" -ForegroundColor Green
Write-Host "     $RegName = $cmdValue" -ForegroundColor Gray

# --- Kill any existing server process on port ---
$existing = netstat -ano 2>$null | Select-String ":$Port\s"
if ($existing) {
    $pids = $existing | ForEach-Object {
        ($_ -split '\s+')[-1]
    } | Sort-Object -Unique
    foreach ($procId in $pids) {
        if ($procId -match '^\d+$' -and $procId -ne '0') {
            Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
        }
    }
    Start-Sleep -Seconds 1
    Write-Host "[OK] Stopped existing server process(es) on port $Port." -ForegroundColor Yellow
}

# --- Launch immediately ---
Write-Host "[INFO] Starting web server now..." -ForegroundColor Yellow
Start-Process -FilePath "wscript.exe" -ArgumentList "`"$VbsLauncher`"" -WindowStyle Hidden
Start-Sleep -Seconds 3

# --- Verify port is listening ---
$listening = netstat -ano 2>$null | Select-String ":$Port\s.*LISTENING"
if ($listening) {
    Write-Host "[OK] Server is UP and listening on port $Port." -ForegroundColor Green
} else {
    Write-Host "[WARN] Server may still be starting - check http://localhost:$Port in a moment." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Web app URL:    http://localhost:$Port" -ForegroundColor White
Write-Host "  Log file:       $LogDir\web_server.log" -ForegroundColor White
Write-Host "  Auto-start:     Registry HKCU\Run -> PHLGUOfficials" -ForegroundColor White
Write-Host ""
Write-Host "  Use manage_server.bat to stop/start/restart the server." -ForegroundColor Gray
Write-Host "  To REMOVE auto-start:" -ForegroundColor Gray
Write-Host "    powershell -File install_service.ps1 -Uninstall" -ForegroundColor Gray
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$url = "http://localhost:" + $Port
Start-Process $url
