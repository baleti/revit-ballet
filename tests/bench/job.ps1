# Run one job from a file through the agent API and print the result. Used for ad-hoc measurements.
#   powershell -File job.ps1 -Port 23717 -Engine roslyn -File C:\bench\gfx.cs [-TimeoutSec 600]
param([Parameter(Mandatory)][int]$Port, [Parameter(Mandatory)][string]$Engine, [Parameter(Mandatory)][string]$File, [int]$TimeoutSec = 600)
Add-Type -AssemblyName System.Net.Http
Add-Type @"
using System.Net;
public static class TrustAll { public static void Enable() { ServicePointManager.ServerCertificateValidationCallback = (a, b, c, d) => true; } }
"@
[TrustAll]::Enable(); [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$tok = (Get-Content (Join-Path $env:APPDATA "revit-ballet\runtime\network\token") -Raw).Trim()
$client = New-Object System.Net.Http.HttpClient; $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSec + 30); $client.DefaultRequestHeaders.Add("X-Auth-Token", $tok)
$code = [IO.File]::ReadAllText($File)
$body = @{ engine = $Engine; code = $code; timeoutMs = ($TimeoutSec * 1000); wait = ($TimeoutSec * 1000); dialogPolicy = "fail" } | ConvertTo-Json -Compress
$sw = [Diagnostics.Stopwatch]::StartNew()
$t = $client.PostAsync("https://127.0.0.1:$Port/v1/jobs", (New-Object System.Net.Http.StringContent($body, [Text.Encoding]::UTF8, "application/json"))).Result.Content.ReadAsStringAsync().Result
$j = $t | ConvertFrom-Json
"state=$($j.state) seconds=$([math]::Round($sw.Elapsed.TotalSeconds, 1)) error=$($j.error.message)"
if ($j.output) { $j.output.TrimEnd() }
if ($j.value) { "value: $($j.value.text)" }
