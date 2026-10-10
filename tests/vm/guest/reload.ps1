# Ask the running agent to reload its runtime from the hot-reload folder; prints the build before and after.
param([Parameter(Mandatory)][int]$Port)
$tok = (Get-Content (Join-Path $env:APPDATA "revit-ballet\runtime\network\token") -Raw).Trim()
function Info { (curl.exe -sk -m 30 -H "X-Auth-Token: $tok" "https://127.0.0.1:$Port/v1/info" | ConvertFrom-Json).agent }
"before: $((Info).build)"
$r = curl.exe -sk -m 90 -X POST -H "X-Auth-Token: $tok" "https://127.0.0.1:$Port/v1/admin/reload"
$a = Info
"after:  $($a.build) from $($a.loadedFrom)"
if ($r -notmatch '"reloaded":true') { "RELOAD FAILED: $r"; exit 1 }
