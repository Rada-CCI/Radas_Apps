$exe = Join-Path $PSScriptRoot 'CCI_New_PC_Setup_v0.1.10.exe'
$log = Join-Path $PSScriptRoot 'exe_run_log.txt'
$si = New-Object System.Diagnostics.ProcessStartInfo($exe)
$si.RedirectStandardOutput = $true
$si.RedirectStandardError = $true
$si.UseShellExecute = $false
$p = [System.Diagnostics.Process]::Start($si)
$out = $p.StandardOutput.ReadToEnd()
$err = $p.StandardError.ReadToEnd()
$p.WaitForExit()
Set-Content -Path $log -Value ("STDOUT:`n" + $out + "`nSTDERR:`n" + $err) -Encoding UTF8
Write-Host "Wrote $log"
