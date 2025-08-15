Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$w=[System.Windows.Forms.SystemInformation]::VirtualScreen.Width
$h=[System.Windows.Forms.SystemInformation]::VirtualScreen.Height
$bmp=New-Object System.Drawing.Bitmap($w,$h)
$g=[System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen([System.Drawing.Point]::Empty,[System.Drawing.Point]::Empty,$bmp.Size)
$out=Join-Path $PSScriptRoot 'gui_screenshot.png'
$bmp.Save($out,[System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()
Write-Host "Saved $out"
