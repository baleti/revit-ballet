# Runs inside the interactive session (started by shot.ps1): saves the whole desktop to C:\ci\out\desktop.png
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$b = [System.Windows.Forms.SystemInformation]::VirtualScreen
$bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($b.Left, $b.Top, 0, 0, $bmp.Size)
New-Item -ItemType Directory -Force C:\ci\out | Out-Null
$bmp.Save("C:\ci\out\desktop.png", [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
