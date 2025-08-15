# Package release artifacts into a versioned zip
Set-Location $PSScriptRoot
$v = (Get-Content -Path .\VERSION -Raw).Trim()
$exe = Join-Path $PSScriptRoot "CCI_New_PC_Setup_v${v}.exe"
$zip = Join-Path $PSScriptRoot "CCI_New_PC_Setup_v${v}.zip"
if (-not (Test-Path $exe)) {
    Write-Error "Expected exe not found: $exe"
    exit 1
}
$items = @($exe, "README_computer_namer.md")
# include assets folder if it exists
if (Test-Path .\assets) { $items += (Get-ChildItem -Path .\assets -Recurse | ForEach-Object { $_.FullName }) }
Compress-Archive -Path $items -DestinationPath $zip -Force
Write-Output $zip
