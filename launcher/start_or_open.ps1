param(
    [switch]$Check,
    # start.bat forwards %* wholesale; unknown args must be ignored (the old
    # batch script ignored them) instead of failing parameter binding.
    [Parameter(ValueFromRemainingArguments = $true)]
    $IgnoredArgs
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$appPy = Join-Path $repoRoot 'app.py'
$appUrl = 'http://127.0.0.1:5000/scores'

function Get-PythonLaunch {
    $py = Get-Command py.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($py) {
        return @{
            File = $py.Source
            Args = @('-3', $appPy)
            Label = 'py -3'
        }
    }

    $python = Get-Command python.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($python) {
        return @{
            File = $python.Source
            Args = @($appPy)
            Label = 'python'
        }
    }

    return $null
}

function Test-Bm2Ready {
    param([int]$Seconds = 5)

    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        try {
            $response = Invoke-WebRequest -UseBasicParsing -Uri $appUrl -TimeoutSec 2
            if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 500) {
                return $true
            }
        } catch {
            $response = $_.Exception.Response
            if ($response -and [int]$response.StatusCode -ge 200 -and [int]$response.StatusCode -lt 500) {
                return $true
            }
        }

        Start-Sleep -Milliseconds 200
    }

    return $false
}

function Stop-StaleBm2Process {
    $connections = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue
    foreach ($connection in @($connections)) {
        $process = Get-CimInstance Win32_Process -Filter "ProcessId=$($connection.OwningProcess)" -ErrorAction SilentlyContinue
        if (-not $process) {
            continue
        }

        $isPython = $process.Name -match '^(python|py)'
        $isThisApp = ([string]$process.CommandLine) -match 'app\.py'
        if ($isPython -and $isThisApp) {
            Stop-Process -Id $connection.OwningProcess -Force -ErrorAction SilentlyContinue
        }
    }

    Start-Sleep -Milliseconds 500
}

if (-not (Test-Path $appPy)) {
    Write-Error "app.py not found at $appPy"
    exit 1
}

$python = Get-PythonLaunch
if (-not $python) {
    Write-Error 'Python was not found. Please install Python first.'
    exit 1
}

if ($Check) {
    Write-Output 'Start check passed'
    Write-Output "Current folder: $repoRoot"
    Write-Output "Python command: $($python.Label)"
    Write-Output "URL: $appUrl"
    exit 0
}

if (Test-Bm2Ready -Seconds 5) {
    Start-Process $appUrl
    exit 0
}

Stop-StaleBm2Process

$waitAndOpen = @"
`$url = '$appUrl'
`$deadline = (Get-Date).AddSeconds(30)
while ((Get-Date) -lt `$deadline) {
    try {
        `$response = Invoke-WebRequest -UseBasicParsing -Uri `$url -TimeoutSec 2
        if (`$response.StatusCode -ge 200 -and `$response.StatusCode -lt 500) {
            Start-Process `$url
            exit 0
        }
    } catch {
        `$response = `$_.Exception.Response
        if (`$response -and [int]`$response.StatusCode -ge 200 -and [int]`$response.StatusCode -lt 500) {
            Start-Process `$url
            exit 0
        }
    }

    Start-Sleep -Milliseconds 200
}
exit 1
"@

Start-Process -FilePath 'powershell.exe' `
    -ArgumentList @('-NoProfile', '-WindowStyle', 'Hidden', '-Command', $waitAndOpen) `
    -WindowStyle Hidden | Out-Null

Push-Location $repoRoot
try {
    & $python.File @($python.Args)
    if ($null -ne $LASTEXITCODE) {
        exit $LASTEXITCODE
    }
    exit 0
} finally {
    Pop-Location
}
