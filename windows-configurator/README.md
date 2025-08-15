# Windows Configurator (PowerShell)

This small project provides a PowerShell-based configurator that applies a number of Windows settings requested by the user. It aims to be safe and reversible where possible. The script can be compiled to a single .exe using the included `build.ps1` which uses the `ps2exe` module.

Files:
- `configure-windows.ps1` - main script. Run elevated.
- `build.ps1` - helper to compile the script into an .exe using `ps2exe`.

Quick usage:
1. Open an elevated PowerShell (Run as Administrator).
2. Change directory to this folder.
3. (Optional) Test the script in dry run mode:

```powershell
.\configure-windows.ps1 -DryRun
```

4. Run the script to apply safe changes (you will be prompted/opened Settings for some items):

```powershell
.\configure-windows.ps1
```

5. To compile into an .exe (requires internet access once to install module):

```powershell
.\build.ps1 -ScriptPath .\configure-windows.ps1 -OutputExe .\configure-windows.exe
```

Which settings are applied automatically (best-effort):
- Power (disk/display/sleep timeouts set to NEVER for AC and DC) — applied via `powercfg`.
- Taskbar alignment (center) and default size — applied via registry (best-effort).
- Date & Time — timezone can be set via `-TimeZoneId` parameter; time resync attempted.
- Notifications/Focus Assist — best-effort registry changes and Settings UI opened.
- Windows Update — attempts to run PSWindowsUpdate to install updates; afterwards stops and disables the `wuauserv` service.
- Personalization — sets dark theme (Apps and System) and attempts to set accent color to `#F18232`; sets background if a sample image exists at `%USERPROFILE%\Pictures\CC Background with support info.jpg`.

What is NOT fully automated (requires manual confirmation due to system differences):
- Power button / Sleep button actions: the script opens the Power Options control panel for you to set these to "Do nothing".
- Some taskbar toggles (Search icon only, Widgets off, Share window from taskbar, combine buttons labels) vary across Windows versions and are opened in the Taskbar UI for you to verify.

Notes and safety:
- The script makes registry changes and stops services; please review the script before running and run in a test environment if possible.
- If you want additional flags applied automatically, tell me which specific behavior(s) to add and I can extend the script. Some items require undocumented GUIDs or COM calls and will need per-machine testing.

