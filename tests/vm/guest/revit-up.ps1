# Start Revit <Year> on the interactive desktop and wait until the revit-ballet server listens. Prints PORT=<n>.
param([Parameter(Mandatory)][int]$Year, [int]$TimeoutSec = 300)
Get-Process Revit -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
& C:\ci\it.ps1 -Name "ci-revit$Year" -Exe "C:\Program Files\Autodesk\Revit $Year\Revit.exe"
$t0 = Get-Date; $port = $null
while (((Get-Date) - $t0).TotalSeconds -lt $TimeoutSec -and -not $port) {
  Start-Sleep 5
  $rev = Get-Process Revit -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($rev) {
    $c = Get-NetTCPConnection -State Listen -OwningProcess $rev.Id -ErrorAction SilentlyContinue |
         Where-Object { $_.LocalPort -ge 23717 -and $_.LocalPort -le 23817 } | Select-Object -First 1
    if ($c) { $port = $c.LocalPort }
  }
}
if (-not $port) { Write-Output "TIMEOUT after $TimeoutSec s: revit running=$([bool](Get-Process Revit -ErrorAction SilentlyContinue)), no server port"; exit 2 }
Write-Output "PORT=$port"
