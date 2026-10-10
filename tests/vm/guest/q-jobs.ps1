# Prints the agent's job list of the first Revit session as "engine state message | output" lines (for tests that start a command from outside).
$tok = (Get-Content (Join-Path $env:APPDATA "revit-ballet\runtime\network\token") -Raw).Trim()
$port = (Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Where-Object { $_.LocalPort -ge 23717 -and $_.LocalPort -le 23817 } | Sort-Object LocalPort | Select-Object -First 1).LocalPort
$j = (& curl.exe -sk -m 10 -H "X-Auth-Token: $tok" "https://127.0.0.1:$port/v1/jobs") | ConvertFrom-Json
foreach ($x in $j.jobs) { "{0} {1} {2} {3} | {4}" -f $x.id, $x.engine, $x.state, $x.error.message, ($x.output -replace "\s+", " ").Trim() }
