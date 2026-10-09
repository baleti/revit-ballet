# Desktop screenshot of the interactive session -> C:\ci\out\desktop.png (the display is on the passthrough GPU, so a console screendump is blank)
Remove-Item C:\ci\out\desktop.png -ErrorAction SilentlyContinue
& C:\ci\it.ps1 -Name ci-shot -Exe powershell.exe -Arguments "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\ci\capture.ps1"
for ($i = 0; $i -lt 20 -and -not (Test-Path C:\ci\out\desktop.png); $i++) { Start-Sleep 1 }
Start-Sleep 1
if (-not (Test-Path C:\ci\out\desktop.png)) { Write-Output "NO-SCREENSHOT (no interactive session?)"; exit 1 }
Write-Output "OK"
