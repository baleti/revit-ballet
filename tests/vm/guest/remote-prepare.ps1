# Makes sure the add-in is installed and a Revit 2026 agent answers. The picker test uninstalls Revit Ballet at its end, so a rerun starts from nothing:
# stop Revit (so the installer can write the real folder, not the .update one), install, start Revit again.
param([int]$Year = 2026)
$dll = Join-Path $env:APPDATA "Autodesk\Revit\Addins\$Year\revit-ballet\revit-ballet.dll"
function Agent-Up { [bool](Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Where-Object { $_.LocalPort -ge 23717 -and $_.LocalPort -le 23817 }) }
& C:\ci\revit-down.ps1 -Year $Year | Out-Null
Remove-Item (Join-Path $env:APPDATA "Autodesk\Revit\Addins\$Year\revit-ballet.update") -Recurse -Force -ErrorAction SilentlyContinue
$p = Start-Process "C:\ci\installer.exe" -ArgumentList "--quiet --remote-picker off" -Wait -PassThru -WindowStyle Hidden
"installer exit $($p.ExitCode); add-in present: $(Test-Path $dll)"
& C:\ci\trust-and-sign.ps1 -Years $Year | Out-Null      # without it Revit stops at the "unsigned add-in" prompt, exactly as ci.sh avoids
& C:\ci\revit-up.ps1 -Year $Year -TimeoutSec 240
"agent up: $(Agent-Up)"
