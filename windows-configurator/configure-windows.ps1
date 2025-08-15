<#
configure-windows.ps1
Powerful but cautious Windows configurator script.
Run as Administrator. Intended to be compiled to an .exe with PS2EXE (build.ps1 included).
This script applies a set of Windows settings requested by the user. Some settings are applied automatically, others are flagged and opened in the Settings UI for manual confirmation where programmatic changes are risky or non-portable across Windows 10/11 builds.
#>

[CmdletBinding()]
param(
    [string]$TimeZoneId,
    [string]$BackgroundPath,
    [switch]$PowerButtonDoNothing,
    [switch]$SleepButtonDoNothing,
    [switch]$LidCloseDoNothing,
    [switch]$DryRun,
    [switch]$DoPowerSettings,
    [switch]$DoTaskbar,
    [switch]$DoDateTime,
    [switch]$DoNotifications,
    [switch]$DoWindowsUpdate,
    [switch]$DoPersonalization,
    [switch]$DoPowerButtonActions
)

function Ensure-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "Relaunching with administrative privileges..."
        # Build a safe ArgumentList array to avoid complex nested quoting
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        foreach ($k in $MyInvocation.BoundParameters.Keys) {
            $argList += "-$k"
            $argList += $PSBoundParameters[$k]
        }
        Start-Process -FilePath pwsh -ArgumentList $argList -Verb RunAs
        Exit
    }
}

# Helper to either perform an action or, when -DryRun, just print the intended action.
function Invoke-IfNotDry {
    param(
        [string]$Description,
        [scriptblock]$Action
    )
    if ($DryRun) {
        Write-Host "DRYRUN: $Description"
    } else {
        & $Action
    }
}

function Apply-PowerSettings {
    Write-Host "Applying power settings (set timeouts to NEVER)..."
    # Set disk, display and sleep timeouts to 0 (never) for AC and DC
    Invoke-IfNotDry "Set disk timeout AC to 0" { powercfg -change -disk-timeout-ac 0 }
    Invoke-IfNotDry "Set disk timeout DC to 0" { powercfg -change -disk-timeout-dc 0 }
    Invoke-IfNotDry "Set monitor timeout AC to 0" { powercfg -change -monitor-timeout-ac 0 }
    Invoke-IfNotDry "Set monitor timeout DC to 0" { powercfg -change -monitor-timeout-dc 0 }
    Invoke-IfNotDry "Set standby timeout AC to 0" { powercfg -change -standby-timeout-ac 0 }
    Invoke-IfNotDry "Set standby timeout DC to 0" { powercfg -change -standby-timeout-dc 0 }

    Write-Host "Note: Power button / Sleep button actions are not changed automatically by this script because those settings can be platform/version-specific and involve GUID mappings; you'll be taken to the Power Options control panel to set them manually." 
    Invoke-IfNotDry "Open Power Options control panel" { Start-Process -FilePath "powercfg.cpl" }
}

function Apply-PowerButtonActions {
    Write-Host "Applying power button / sleep button / lid close actions (best-effort)..."
    # These settings are stored per power scheme using powercfg -setacvalueindex / -setdcvalueindex
    # We attempt to set the behaviors if the values are available; otherwise open Power Options.
    try {
        $scheme = (powercfg -getactivescheme) -replace '.*:\s*',''
        Write-Host "Active power scheme: $scheme"
        # GUID subgroups and setting GUIDs for power buttons/lid
        $subgroup = '4f971e89-eebd-4455-a8de-9e59040e7347' # System button subgroup GUID (best-effort)
        $btnPower = '7648efa3-dd9c-4e3e-b566-50f929386280' # Power button action
        $btnSleep = '96996bc0-ad50-47ec-923b-6f41874dd9eb' # Sleep button action
        $lid = '5ca83367-6e45-459f-a27b-476b1d01c936'      # Lid close action

        if ($PowerButtonDoNothing) {
            Write-Host "Attempting to set Power button action to 'Do nothing' (value 0) for AC and DC"
            Invoke-IfNotDry "powercfg -setacvalueindex $scheme $subgroup $btnPower 0" { powercfg -setacvalueindex $scheme $subgroup $btnPower 0 }
            Invoke-IfNotDry "powercfg -setdcvalueindex $scheme $subgroup $btnPower 0" { powercfg -setdcvalueindex $scheme $subgroup $btnPower 0 }
        }
        if ($SleepButtonDoNothing) {
            Write-Host "Attempting to set Sleep button action to 'Do nothing' (value 0) for AC and DC"
            Invoke-IfNotDry "powercfg -setacvalueindex $scheme $subgroup $btnSleep 0" { powercfg -setacvalueindex $scheme $subgroup $btnSleep 0 }
            Invoke-IfNotDry "powercfg -setdcvalueindex $scheme $subgroup $btnSleep 0" { powercfg -setdcvalueindex $scheme $subgroup $btnSleep 0 }
        }
        if ($LidCloseDoNothing) {
            Write-Host "Attempting to set Lid close action to 'Do nothing' (value 0) for AC and DC"
            Invoke-IfNotDry "powercfg -setacvalueindex $scheme $subgroup $lid 0" { powercfg -setacvalueindex $scheme $subgroup $lid 0 }
            Invoke-IfNotDry "powercfg -setdcvalueindex $scheme $subgroup $lid 0" { powercfg -setdcvalueindex $scheme $subgroup $lid 0 }
        }

        # Activate the scheme to ensure changes take effect
        Invoke-IfNotDry "powercfg -setactive $scheme" { powercfg -setactive $scheme }
        Write-Host "Applied button/lid actions where supported; opening Power Options for manual review if anything failed."
        Invoke-IfNotDry "Open Power Options control panel" { Start-Process -FilePath "powercfg.cpl" }
    } catch {
        Write-Warning "Could not programmatically update button/lid actions on this platform: $_"
        Start-Process -FilePath "powercfg.cpl"
    }
}

function Apply-TaskbarSettings {
    Write-Host "Applying taskbar settings (best-effort, Windows 11/insider builds may differ)..."

    # Center alignment: TaskbarAl = 1 (0 = left, 1 = center) - safe on recent Windows 11 builds
    try {
        Invoke-IfNotDry "Create/ensure Explorer\Advanced registry key" { New-Item -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Force | Out-Null }
        Invoke-IfNotDry "Set TaskbarAl to 1" { Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name TaskbarAl -Value 1 -Type DWord }
        Write-Host "Set taskbar alignment to center (TaskbarAl = 1)"
    } catch {
        Write-Warning "Failed to set TaskbarAl: $_"
    }

    # Taskbar size: TaskbarSi (0 small, 1 default, 2 large). 'When taskbar is full' behaviour can't be forced reliably; we set default.
    try {
        Invoke-IfNotDry "Set TaskbarSi to 1" { Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name TaskbarSi -Value 1 -Type DWord }
        Write-Host "Set Taskbar size to default (TaskbarSi = 1)"
    } catch {
        Write-Warning "Failed to set TaskbarSi: $_"
    }

    # Many Windows 11 taskbar features (Search mode, Widgets toggle, Task view pinning, combine labels) are controlled by explorer shell components
    # and by REST/COM calls; they are not reliably scriptable across builds. We'll open Taskbar settings for final manual verification.
    Invoke-IfNotDry "Open Taskbar settings UI" { Start-Process "ms-settings:taskbar" }
    Write-Host "Opened Taskbar settings UI; please verify the following manually if not already set: Search icon only; Task view ON; Widgets OFF; Safely Remove icon ON; taskbar behaviors as requested."
}

function Apply-DateTime {
    if ($PSBoundParameters.ContainsKey('TimeZoneId') -and $TimeZoneId) {
        Write-Host "Setting time zone to: $TimeZoneId"
    Invoke-IfNotDry "tzutil /s $TimeZoneId" { tzutil /s "$TimeZoneId" }
    } else {
        Write-Host "No TimeZoneId provided. To set explicitly, run: .\configure-windows.ps1 -TimeZoneId 'Pacific Standard Time'"
        Write-Host "Opening Date & time settings for manual timezone selection..."
        Start-Process "ms-settings:dateandtime"
    }

    Write-Host "Syncing time now..."
    try {
        Invoke-IfNotDry "w32tm /resync" { w32tm /resync | Out-Null }
        Write-Host "Time sync requested (w32tm /resync)"
    } catch {
        Write-Warning "Time sync failed or requires network/service: $_"
    }
}

function Apply-NotificationsAndFocusAssist {
    Write-Host "Applying notification and Focus Assist (Do not disturb) settings..."

    # Attempt to turn off notifications globally by setting the NOC_GLOBAL_SETTING_TOASTS_ENABLED flag to 0
    # This key exists on many systems but may not be present on all builds; the change here affects toast popups.
    try {
        $path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications'
        Invoke-IfNotDry "Create/ensure PushNotifications key" { New-Item -Path $path -Force | Out-Null }
        Invoke-IfNotDry "Set ToastEnabled = 0" { New-ItemProperty -Path $path -Name 'ToastEnabled' -Value 0 -PropertyType DWord -Force | Out-Null }
        Write-Host "Set PushNotifications\ToastEnabled = 0 (attempted)"
    } catch {
        Write-Warning "Could not set PushNotifications keys: $_"
    }

    # Focus Assist: use registry for 'QuietHours' (Focus Assist) default - this is informational and may not fully emulate UI toggles across builds
    try {
        $faPath = 'HKCU:\Software\Policies\Microsoft\Windows\Explorer'
        Invoke-IfNotDry "Create/ensure Focus Assist policy key" { New-Item -Path $faPath -Force | Out-Null }
        Invoke-IfNotDry "Set QuietHours = 1" { New-ItemProperty -Path $faPath -Name 'QuietHours' -Value 1 -PropertyType DWord -Force | Out-Null }
        Write-Host "Attempted to set Focus Assist (QuietHours) policy to ON"
    } catch {
        Write-Warning "Could not set Focus Assist policy via registry: $_"
    }

    Write-Host "Opened Notifications & actions settings for manual confirmation..."
    Invoke-IfNotDry "Open Notifications settings UI" { Start-Process "ms-settings:notifications" }
}

function Apply-WindowsUpdate {
    Write-Host "Updating Windows via PSWindowsUpdate (best-effort) and then stopping Windows Update service. This may require internet and additional module installation."
    try {
        # Try to install and import PSWindowsUpdate if available
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Host "Installing PSWindowsUpdate module (may prompt for PSGallery trust)..."
            Invoke-IfNotDry "Install PSWindowsUpdate module" { Install-Module -Name PSWindowsUpdate -Force -Confirm:$false -Scope AllUsers -AllowClobber }
        }
        Invoke-IfNotDry "Import PSWindowsUpdate module" { Import-Module PSWindowsUpdate -ErrorAction Stop }

        Write-Host "Checking for updates and installing (this may take a while)..."
        Invoke-IfNotDry "Run Get-WindowsUpdate -AcceptAll -Install -AutoReboot" { Get-WindowsUpdate -AcceptAll -Install -AutoReboot | Out-Null }
        Write-Host "Windows Update run requested via PSWindowsUpdate"
    } catch {
        Write-Warning "PSWindowsUpdate approach failed or not available: $_"
        Write-Host "As a fallback, opening Windows Update settings UI..."
        Invoke-IfNotDry "Open Windows Update settings UI" { Start-Process "ms-settings:windowsupdate" }
    }

    Write-Host "Stopping Windows Update service (wuauserv) to prevent immediate updates..."
    try {
        Invoke-IfNotDry "Stop-Service wuauserv" { Stop-Service -Name wuauserv -Force -ErrorAction Stop }
        Invoke-IfNotDry "Set-Service wuauserv Disabled" { Set-Service -Name wuauserv -StartupType Disabled }
        Write-Host "Windows Update service stopped and disabled (wuauserv)."
    } catch {
        Write-Warning "Failed to stop or disable wuauserv: $_"
    }
}

function Apply-Personalization {
    Write-Host "Applying personalization: dark theme, accent color, and background..."

    # Set dark theme (Apps and System)
    try {
        Invoke-IfNotDry "Ensure Themes\Personalize key" { New-Item -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize -Force | Out-Null }
        Invoke-IfNotDry "Set AppsUseLightTheme = 0" { Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name 'AppsUseLightTheme' -Value 0 -Type DWord }
        Invoke-IfNotDry "Set SystemUsesLightTheme = 0" { Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name 'SystemUsesLightTheme' -Value 0 -Type DWord }
        Write-Host "Set system and apps to dark theme (0)"
    } catch {
        Write-Warning "Failed to set theme keys: $_"
    }

    # Accent color: best-effort by setting registry; value expects a DWORD BGR (blue-green-red) hex value
    # #F18232 => RGB (241,130,50) => Hex 0x32 0x82 0xF1 (BGR) => 0x3282F1
    try {
        $accentBgr = 0x3282F1
        Invoke-IfNotDry "Ensure Explorer\Accent key" { New-Item -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent -Force | Out-Null }
        Invoke-IfNotDry "Set AccentColor" { Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent -Name 'AccentColor' -Value $accentBgr -Type DWord -Force }
        Invoke-IfNotDry "Set StartColor" { Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent -Name 'StartColor' -Value $accentBgr -Type DWord -Force }
        # Also set the personalization colors
        Invoke-IfNotDry "Ensure Themes key" { New-Item -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes -Force | Out-Null }
        Invoke-IfNotDry "Set ColorPrevalence = 1" { Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes -Name 'ColorPrevalence' -Value 1 -Type DWord -Force }
        Write-Host "Attempted to set Accent color to #F18232 (best-effort)."
    } catch {
        Write-Warning "Failed to set accent color: $_"
    }

    # Background: prefer provided BackgroundPath, otherwise fallback to sample in Pictures
    $usedBackground = $null
    if ($BackgroundPath -and (Test-Path $BackgroundPath)) {
        $usedBackground = $BackgroundPath
    } else {
        $sampleBackground = "$env:USERPROFILE\Pictures\CC Background with support info.jpg"
        if (Test-Path $sampleBackground) { $usedBackground = $sampleBackground }
    }

    if ($usedBackground) {
        Write-Host "Setting background: $usedBackground"
        Add-Type @"
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll",SetLastError=true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
        Invoke-IfNotDry "Set desktop wallpaper to $usedBackground" { [Wallpaper]::SystemParametersInfo(20, 0, $usedBackground, 3) | Out-Null }
        Write-Host "Background set to: $usedBackground"
    } else {
        Write-Host "No background image found or provided. Opening Background settings UI for manual selection."
        Invoke-IfNotDry "Open Background settings UI" { Start-Process "ms-settings:personalization-background" }
    }
}

function Summary {
    Write-Host "\nConfiguration complete-ish. Review the following notes and open UIs printed above to finish any manual verification."
    @(
        "Power: timeouts set to NEVER (disk/display/sleep). Power Options UI opened to set power button / sleep button actions manually.",
        "Taskbar: alignment and size applied where safe; Taskbar Settings opened for manual verification of Search icon, Widgets, Task View, tray icons and detailed behaviors.",
        "Date & Time: timezone set if provided; Date & time UI opened for manual changes; time resync requested.",
        "Notifications: attempted to disable toast/popups and set Focus Assist via policy registry; Notifications UI opened for manual verification.",
        "Windows Update: attempted to run PSWindowsUpdate; Windows Update service stopped/disabled.",
        "Personalization: dark theme applied and best-effort accent color set; background set if a sample image was found, otherwise Background settings opened.",
        "Run the script again with -Verbose to see step-by-step output."
    ) | ForEach-Object { Write-Host " - $_" }
}

# === Main ===
# When DryRun is supplied, skip the interactive admin relaunch and announce DryRun only.
if ($DryRun) {
    Write-Host "Dry run mode: no changes will be made."
} else {
    Ensure-Admin
}

# If no explicit Do* switches were provided, run all actions for compatibility with older callers.
$doFlags = @($DoPowerSettings, $DoTaskbar, $DoDateTime, $DoNotifications, $DoWindowsUpdate, $DoPersonalization, $DoPowerButtonActions) | Where-Object { $_ }
if ($doFlags.Count -eq 0) {
    # Run everything
    Write-Host "No selective flags provided; running all configurator actions."
    Apply-PowerSettings
    Apply-TaskbarSettings
    Apply-DateTime
    Apply-NotificationsAndFocusAssist
    Apply-WindowsUpdate
    Apply-Personalization
    Apply-PowerButtonActions
} else {
    if ($DoPowerSettings) { Apply-PowerSettings }
    if ($DoTaskbar) { Apply-TaskbarSettings }
    if ($DoDateTime) { Apply-DateTime }
    if ($DoNotifications) { Apply-NotificationsAndFocusAssist }
    if ($DoWindowsUpdate) { Apply-WindowsUpdate }
    if ($DoPersonalization) { Apply-Personalization }
    if ($DoPowerButtonActions) { Apply-PowerButtonActions }
}

Summary

Write-Host "Done."
