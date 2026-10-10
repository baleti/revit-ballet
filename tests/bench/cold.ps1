# Cold start of one engine in a freshly started Revit: first job (loads the engine) and memory cost. Prints one JSON line.
#   powershell -File cold.ps1 -Port 23717 -Engine roslyn
param([Parameter(Mandatory)][int]$Port, [Parameter(Mandatory)][string]$Engine)
Add-Type -AssemblyName System.Net.Http
Add-Type @"
using System.Net;
public static class TrustAll { public static void Enable() { ServicePointManager.ServerCertificateValidationCallback = (a, b, c, d) => true; } }
"@
[TrustAll]::Enable(); [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$tok = (Get-Content (Join-Path $env:APPDATA "revit-ballet\runtime\network\token") -Raw).Trim()
$client = New-Object System.Net.Http.HttpClient; $client.DefaultRequestHeaders.Add("X-Auth-Token", $tok)
$code = if ($Engine -eq "roslyn") { "return 1;" } else { "result = 1" }
function Job { $body = @{ engine = $Engine; code = $code; timeoutMs = 120000; wait = 120000 } | ConvertTo-Json -Compress
  $c = New-Object System.Net.Http.StringContent($body, [Text.Encoding]::UTF8, "application/json"); $sw = [Diagnostics.Stopwatch]::StartNew()
  $t = $client.PostAsync("https://127.0.0.1:$Port/v1/jobs", $c).Result.Content.ReadAsStringAsync().Result; $sw.Stop(); [pscustomobject]@{ wall = $sw.Elapsed.TotalMilliseconds; job = ($t | ConvertFrom-Json) } }
$rev = Get-Process Revit | Select-Object -First 1
$before = $rev.WorkingSet64; $privBefore = $rev.PrivateMemorySize64
$first = Job; $rev.Refresh(); $after1 = $rev.WorkingSet64
$second = Job; $rev.Refresh(); $after2 = $rev.WorkingSet64; $priv = $rev.PrivateMemorySize64
[ordered]@{ engine = $Engine; state = $first.job.state; firstJobWallMs = [math]::Round($first.wall); firstCompileMs = $first.job.timings.compileMs; firstRunMs = $first.job.timings.runMs
  secondJobWallMs = [math]::Round($second.wall); secondCompileMs = $second.job.timings.compileMs; secondRunMs = $second.job.timings.runMs
  workingSetBeforeMB = [math]::Round($before / 1MB); workingSetAfterFirstMB = [math]::Round($after1 / 1MB); workingSetAfterSecondMB = [math]::Round($after2 / 1MB)
  engineCostMB = [math]::Round(($after2 - $before) / 1MB); privateCostMB = [math]::Round(($priv - $privBefore) / 1MB) } | ConvertTo-Json -Compress
