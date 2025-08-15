<# build.ps1
Helper to compile configure-windows.ps1 into an .exe using ps2exe.
Requires internet access once to install the ps2exe module. Run in an elevated PowerShell session.
#>

param(
    [string]$ScriptPath = "configure-windows.ps1",
    [string]$OutputExe = "configure-windows.exe"
)

if (-not (Test-Path $ScriptPath)) {
    Write-Error "Script $ScriptPath not found in current folder."
    exit 1
}

# Ensure elevated
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Please run this build script as Administrator."; exit 1
}

if (-not (Get-Command Invoke-ps2exe -ErrorAction SilentlyContinue)) {
    Write-Host "Installing ps2exe module (Invoke-ps2exe)..."
    Install-Module -Name ps2exe -Scope CurrentUser -Force
}

Import-Module ps2exe -Force

Invoke-ps2exe -inputFile $ScriptPath -outputFile $OutputExe -noConsole -icon "$PSScriptRoot\app.ico" -verbose

Write-Host "Built: $OutputExe"
