# BM2 supervisor — keeps pythonw app.py alive without polling.
#
# Design:
#   * Task Scheduler fires this once at user login.
#   * The script spawns pythonw and then BLOCKS on Wait-Process. While
#     blocked it uses ~zero CPU and ~20 MB RAM — no polling, no timers.
#   * The moment pythonw dies (crash, taskkill, anything), Wait-Process
#     returns and we respawn immediately.
#   * A 2-second debounce prevents tight respawn loops when app.py is
#     fundamentally broken (e.g., port permanently taken by another app).
#
# Single-instance guard: a named system mutex. If two supervisors race
# (e.g., a stale one is still alive from a previous fast logoff/logon),
# only one wins; the loser exits cleanly without disturbing pythonw.

$ErrorActionPreference = 'SilentlyContinue'

# --- single-instance guard ---------------------------------------------------
$mutexName = "Global\BM2_Supervisor_v1"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)
if (-not $mutex.WaitOne(0)) {
    # Another supervisor is already running. Bail.
    exit 0
}

try {
    # --- resolve interpreter + app path -------------------------------------
    $here     = Split-Path -Parent $MyInvocation.MyCommand.Path
    $repoRoot = Split-Path -Parent $here
    $appPy    = Join-Path $repoRoot 'app.py'
    if (-not (Test-Path $appPy)) { exit 1 }

    $pyw = $null
    $pythonExe = (py -3 -c "import sys; print(sys.executable)" 2>$null)
    if ($pythonExe) {
        $candidate = Join-Path (Split-Path -Parent $pythonExe) 'pythonw.exe'
        if (Test-Path $candidate) { $pyw = $candidate }
    }
    if (-not $pyw) { $pyw = (Get-Command pythonw.exe).Path }
    if (-not $pyw) { exit 2 }

    # --- supervise forever --------------------------------------------------
    while ($true) {
        # If port 5000 is already taken by *something* (could be a lingering
        # pythonw from a previous boot, or another app), don't double-spawn.
        # Wait until it's released, then start fresh.
        $occupied = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue
        if ($occupied) {
            $existingPid = $occupied[0].OwningProcess
            try {
                # Block on the existing process; when it dies we'll respawn.
                Wait-Process -Id $existingPid -ErrorAction Stop
            } catch {
                # Process vanished between check and wait — fine, loop.
            }
            Start-Sleep -Seconds 1
            continue
        }

        # Spawn our pythonw and block until it exits.
        $proc = Start-Process -FilePath $pyw `
                              -ArgumentList 'app.py' `
                              -WorkingDirectory $repoRoot `
                              -WindowStyle Hidden `
                              -PassThru
        if (-not $proc) {
            Start-Sleep -Seconds 10
            continue
        }
        Wait-Process -Id $proc.Id

        # Debounce: if pythonw exits in <2s repeatedly, slow the loop down so
        # we don't pin CPU when something's fundamentally wrong.
        Start-Sleep -Seconds 2
    }
} finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
