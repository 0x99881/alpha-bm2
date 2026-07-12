# Creates/updates the Binance Alpha desktop shortcut that runs start.bat minimized.
# Keep this file ASCII-safe so Windows PowerShell 5.1 reads it correctly.

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot  = Split-Path -Parent $scriptDir
$startBat  = Join-Path $repoRoot 'start.bat'

if (-not (Test-Path $startBat)) {
    Write-Error "start.bat not found at $startBat"
    exit 1
}

$desktop      = [Environment]::GetFolderPath('Desktop')
$shortcutName = [string][char]0x5E01 + [string][char]0x5B89 + 'Alpha' + [string][char]0x591A + [string][char]0x53F7 + [string][char]0x7BA1 + [string][char]0x7406 + [string][char]0x7CFB + [string][char]0x7EDF + '.lnk'
$shortcutPath = Join-Path $desktop $shortcutName

$shell    = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath       = $startBat
$shortcut.WorkingDirectory = $repoRoot
# 7 = minimized window (cmd.exe still runs, just doesn't steal focus).
$shortcut.WindowStyle      = 7
$shortcut.Description      = 'Start the Binance Alpha multi-account manager and open the browser.'
# Use cmd.exe's icon so it's recognizable; fallback to default if not found.
$cmdIcon = Join-Path $env:WINDIR 'System32\cmd.exe'
if (Test-Path $cmdIcon) { $shortcut.IconLocation = "$cmdIcon,0" }
$shortcut.Save()

Write-Output ("Created: " + $shortcutPath)
Write-Output ("Target : " + $startBat)
Write-Output ""
Write-Output "Double-click the desktop shortcut to start the Binance Alpha multi-account manager."
