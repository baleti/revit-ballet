# Runs ON a target machine: invoke one add-in command in one Revit instance. Parameters arrive in $P (see revit-remote). Prints the job JSON.
# $P = { port, command, wait (ms, 0 = return at once), timeoutMs (optional), dialogPolicy (optional: "report"|"fail"|{respond}) }
$ErrorActionPreference = 'Stop'
$tokFile = Join-Path $env:APPDATA 'revit-ballet\runtime\network\token'
if (-not (Test-Path $tokFile)) { '{"error":{"kind":"no_token","message":"no agent token on this machine"}}'; exit 2 }
$tok = (Get-Content $tokFile -Raw).Trim()
$body = [ordered]@{ wait = [int]$P.wait }
if ($P.timeoutMs) { $body.timeoutMs = [int]$P.timeoutMs }
if ($P.dialogPolicy) { $body.dialogPolicy = $P.dialogPolicy }
$tmp = Join-Path $env:TEMP ("rb-remote-" + [guid]::NewGuid().ToString('N') + '.json')
[IO.File]::WriteAllText($tmp, ($body | ConvertTo-Json -Compress -Depth 4), (New-Object Text.UTF8Encoding($false)))
try {
  $name = [uri]::EscapeDataString([string]$P.command)
  & curl.exe -sk -m ([Math]::Max(30, [int]($P.wait / 1000) + 30)) -X POST -H "X-Auth-Token: $tok" -H 'Content-Type: application/json' --data-binary "@$tmp" "https://127.0.0.1:$($P.port)/v1/commands/$name/invoke"
} finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
