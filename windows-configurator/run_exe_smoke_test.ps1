param(
    [string]$exe
)

if (-not $exe) {
    # default (back-compat)
    $exe = Join-Path $PSScriptRoot 'CCI_New_PC_Setup_v0.1.6.exe'
}

Write-Host "Starting: $exe"

try {
    # Start the exe and capture process info (do not hide window so GUI is visible)
    $p = Start-Process -FilePath $exe -PassThru -ErrorAction Stop
} catch {
    Write-Host "Failed to start exe: $_"
    exit 2
}

Start-Sleep -Seconds 4
if (Get-Process -Id $p.Id -ErrorAction SilentlyContinue) {
    Write-Host "RUNNING (PID $($p.Id))"
} else {
    Write-Host "NOT_RUNNING"
}

# stop the process if running
try {
    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    Write-Host "STOPPED"
} catch {
    Write-Host "Failed to stop process: $_"
}

exit 0
