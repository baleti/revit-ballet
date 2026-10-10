# Stop Revit: only the given year's process with -Year, every Revit without.
param([int]$Year = 0)
$exe = if ($Year) { "C:\Program Files\Autodesk\Revit $Year\Revit.exe" } else { $null }
Get-Process Revit -ErrorAction SilentlyContinue | Where-Object { -not $exe -or $_.Path -eq $exe } | Stop-Process -Force
Start-Sleep 2
