param(
    [string]$BaseName = "CCI_New_PC_Setup",
    [string]$Script = "computer_namer_gui.py",
    [switch]$Console
)

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$verFile = Join-Path $root "VERSION"
if (-not (Test-Path $verFile)) {
    "0.1.0" | Out-File -FilePath $verFile -Encoding UTF8
}

# Bump patch
$ver = Get-Content $verFile -Raw
$parts = $ver.Trim().Split('.')
if ($parts.Count -ne 3) { $parts = @('0','0','0') }
[int]$parts[2] += 1
$new = "$($parts[0]).$($parts[1]).$($parts[2])"
$new | Out-File -FilePath $verFile -Encoding UTF8

# Build exe name
$safe = $BaseName -replace ' ', '_'
$outname = "$safe`_v$new.exe"
Write-Host "Building $outname"

# Run PyInstaller (must be on PATH - uses user's Python)
# Script runs with working directory set to the script folder, so reference local paths.
if ($Console) {
    py -3 -m PyInstaller --onefile --uac-admin --add-data "assets;assets" --add-data "configure-windows.ps1;." "$Script"
} else {
    py -3 -m PyInstaller --onefile --windowed --uac-admin --add-data "assets;assets" --add-data "configure-windows.ps1;." "$Script"
}

# Move produced exe
$dist = Join-Path $root 'dist' 
$built = Get-ChildItem -Path $dist -Filter '*.exe' | Where-Object { $_.Name -like "$($Script.Split('.')[0])*" } | Select-Object -First 1
if ($built) {
    Move-Item -Path $built.FullName -Destination (Join-Path $root $outname) -Force
    Write-Host "Built and moved: $outname"
} else {
    Write-Host "Build did not produce expected exe in dist\. Check PyInstaller output." 
}
