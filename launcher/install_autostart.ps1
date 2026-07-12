# Installs a Windows Scheduled Task that auto-starts the BM2 server at user
# login. After running this once, the server will be listening on port 5000
# every time you log in — Chrome just needs to navigate to
# http://127.0.0.1:5000/scores (we suggest bookmarking that).
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File install_autostart.ps1
#   powershell -ExecutionPolicy Bypass -File install_autostart.ps1 -Uninstall
#
# Re-running the script is safe: it removes any prior copy of the task and
# re-registers a fresh one.

param(
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
# ASCII-only task name avoids encoding mismatch when Windows PowerShell 5.1
# loads this .ps1 (it reads as ANSI / system codepage unless the file has a
# UTF-8 BOM). The task is for the user's own machine — readable enough.
$TaskName  = 'BM2_AutoStart'
$here      = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot  = Split-Path -Parent $here
$appPy     = Join-Path $repoRoot 'app.py'

if (-not (Test-Path $appPy)) {
    Write-Error "app.py not found at $appPy"; exit 1
}

# Always remove the existing task first — Register-ScheduledTask without -Force
# would error otherwise, and -Force still complains if Description differs.
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Output "Removed prior task '$TaskName'."
}

if ($Uninstall) {
    Write-Output "Uninstall mode — done."
    exit 0
}

# Locate pythonw.exe so we get a windowless server. Prefer the same Python
# the user already runs via `py -3`, since virtualenvs / installed packages
# live there.
$pyw = $null
$pythonExe = (py -3 -c "import sys; print(sys.executable)" 2>$null)
if ($pythonExe) {
    $candidate = Join-Path (Split-Path -Parent $pythonExe) 'pythonw.exe'
    if (Test-Path $candidate) { $pyw = $candidate }
}
if (-not $pyw) {
    $pyw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue | Select-Object -First 1).Path
}
if (-not $pyw) {
    $pyw = (Get-Command pyw.exe -ErrorAction SilentlyContinue | Select-Object -First 1).Path
}
if (-not $pyw) {
    Write-Error "Neither pythonw.exe nor pyw.exe was found. Install Python's windowless launcher, or edit this script to use py.exe (will show a console)."
    exit 1
}
Write-Output "Using interpreter: $pyw"

# The action runs ensure_running.ps1, which is idempotent: it bails if port
# 5000 is already listening, otherwise spawns pythonw app.py. That lets us
# use the same task for both first-time start (AtLogOn) and crash-recovery
# polling (every 5 min) without ever double-spawning.
$ensureScript = Join-Path $here 'ensure_running.ps1'
if (-not (Test-Path $ensureScript)) {
    Write-Error "ensure_running.ps1 not found next to this script at $ensureScript"; exit 1
}
$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ensureScript`"" `
    -WorkingDirectory $repoRoot

# Single AtLogOn trigger. The supervisor (ensure_running.ps1) stays alive
# for the whole login session and uses Wait-Process to react instantly to
# pythonw exits — no polling, ~0% CPU when idle, ~20 MB RAM.
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Days 0) `
    -Hidden
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask `
    -TaskName    $TaskName `
    -Action      $action `
    -Trigger     $trigger `
    -Settings    $settings `
    -Principal   $principal `
    -Description "Auto-starts BM2 local web server on user login. Server listens on http://127.0.0.1:5000" | Out-Null

Write-Output ""
Write-Output "✓ Scheduled task '$TaskName' installed."
Write-Output ""
Write-Output "Next steps:"
Write-Output "  1. Run the task once now to start the server without rebooting:"
Write-Output "     Start-ScheduledTask -TaskName '$TaskName'"
Write-Output "  2. Bookmark this URL in Chrome:"
Write-Output "     http://127.0.0.1:5000/scores"
Write-Output ""
Write-Output "To uninstall later: run this script with -Uninstall"
