# Start Revit <Year> on the interactive desktop and wait until ITS revit-ballet server listens. Prints PORT=<n>.
# Several years can be started and run side by side: only this year's Revit.exe is touched.
param([Parameter(Mandatory)][int]$Year, [int]$TimeoutSec = 300)
$exe = "C:\Program Files\Autodesk\Revit $Year\Revit.exe"
function Mine { Get-Process Revit -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $exe } }
Mine | Stop-Process -Force
Start-Sleep 2
& C:\ci\it.ps1 -Name "ci-revit$Year" -Exe $exe
$t0 = Get-Date; $port = $null
while (((Get-Date) - $t0).TotalSeconds -lt $TimeoutSec -and -not $port) {
  Start-Sleep 3
  $rev = Mine | Select-Object -First 1
  if ($rev) {
    $c = Get-NetTCPConnection -State Listen -OwningProcess $rev.Id -ErrorAction SilentlyContinue |
         Where-Object { $_.LocalPort -ge 23717 -and $_.LocalPort -le 23817 } | Select-Object -First 1
    if ($c) { $port = $c.LocalPort }
  }
}
if (-not $port) { Write-Output "TIMEOUT after $TimeoutSec s: revit $Year running=$([bool](Mine)), no server port"; exit 2 }
# the port opens before the agent runtime does (it starts when Revit is first idle, later on a loaded machine): wait for /v1/info
$tokFile = Join-Path $env:APPDATA "revit-ballet\runtime\network\token"
while (((Get-Date) - $t0).TotalSeconds -lt $TimeoutSec) {
  if (Test-Path $tokFile) {
    $tok = (Get-Content $tokFile -Raw).Trim()
    $code = curl.exe -sk -m 10 -o NUL -w "%{http_code}" -H "X-Auth-Token: $tok" "https://127.0.0.1:$port/v1/info"
    if ($code -eq "200") { break }
  }
  Start-Sleep 2
}
if (((Get-Date) - $t0).TotalSeconds -ge $TimeoutSec) { Write-Output "TIMEOUT after $TimeoutSec s: port $port open but the agent runtime did not start (Revit never idle?)"; exit 2 }
Write-Output "PORT=$port"
