# Benchmark the three agent engines (roslyn, ironpython2, cpython3) inside a running Revit. Run in the guest, interactive desktop session.
#   powershell -File bench.ps1 -Port 23717 -Tasks C:\bench\tasks -Out C:\bench\out\bench.json [-Only loop,fib] [-Runs 10] [-NeedModel]
# For every task and engine: 2 warm-up runs, then -Runs measured runs. Reports client wall time (HTTP round trip) and the server's own split
# (compile / dispatch to the UI thread / run). Results are cross-checked: every engine must return the same value for the same task.
param([Parameter(Mandatory)][int]$Port, [Parameter(Mandatory)][string]$Tasks, [Parameter(Mandatory)][string]$Out,
      [string[]]$Only = @(), [int]$Runs = 10, [int]$Burst = 100)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Net.Http
Add-Type @"
using System.Net;
public static class TrustAll { public static void Enable() { ServicePointManager.ServerCertificateValidationCallback = (a, b, c, d) => true; } }
"@
[TrustAll]::Enable()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$tok = (Get-Content (Join-Path $env:APPDATA "revit-ballet\runtime\network\token") -Raw).Trim()
$client = New-Object System.Net.Http.HttpClient
$client.Timeout = [TimeSpan]::FromMinutes(15)
$client.DefaultRequestHeaders.Add("X-Auth-Token", $tok)
$client.DefaultRequestHeaders.Add("X-Agent-Name", "bench")
$base = "https://127.0.0.1:$Port"
$engines = @(
  @{ name = "roslyn"; ext = "cs" },
  @{ name = "ironpython2"; ext = "py" },
  @{ name = "cpython3"; ext = "py" })

function Post($engine, $code, $timeoutMs, [int]$wait) {
  $body = @{ engine = $engine; code = $code; timeoutMs = $timeoutMs; wait = $wait } | ConvertTo-Json -Compress
  $content = New-Object System.Net.Http.StringContent($body, [Text.Encoding]::UTF8, "application/json")
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $resp = $client.PostAsync("$base/v1/jobs", $content).Result
  $txt = $resp.Content.ReadAsStringAsync().Result
  $sw.Stop()
  [pscustomobject]@{ Wall = $sw.Elapsed.TotalMilliseconds; Job = ($txt | ConvertFrom-Json) }
}
function GetJob($id, $wait) { $client.GetStringAsync("$base/v1/jobs/${id}?wait=$wait").Result | ConvertFrom-Json }
function Stats($xs) {
  $s = @($xs | Sort-Object); if ($s.Count -eq 0) { return $null }
  $n = $s.Count
  [pscustomobject]@{ median = [math]::Round($s[[int][math]::Floor(($n - 1) / 2)], 1); min = [math]::Round($s[0], 1); p95 = [math]::Round($s[[int][math]::Min($n - 1, [math]::Ceiling(0.95 * $n) - 1)], 1); max = [math]::Round($s[$n - 1], 1); mean = [math]::Round(($s | Measure-Object -Average).Average, 1) }
}
function Split-Times($j, $wall) {
  $sub = [datetime]::Parse($j.submitted).ToUniversalTime(); $st = [datetime]::Parse($j.started).ToUniversalTime(); $fin = [datetime]::Parse($j.finished).ToUniversalTime()
  $dispatch = ($st - $sub).TotalMilliseconds - $j.timings.queueMs - $j.timings.compileMs
  [pscustomobject]@{ wall = $wall; compile = $j.timings.compileMs; dispatch = [math]::Max(0, [math]::Round($dispatch, 1)); run = $j.timings.runMs; total = ($fin - $sub).TotalMilliseconds }
}

$info = (($client.GetStringAsync("$base/v1/info").Result) | ConvertFrom-Json)
$rev = Get-Process Revit | Select-Object -First 1
$result = [ordered]@{
  meta = [ordered]@{ date = (Get-Date -Format s); revit = $info.revit.version; engines = $info.engines | ForEach-Object { "$($_.name) $($_.version)" }
                     cores = (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors; ramGB = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 1)
                     document = (@((& { $c = (Post "roslyn" 'return Doc == null ? "(none)" : Doc.Title + " / " + new FilteredElementCollector(Doc).WhereElementIsNotElementType().GetElementCount() + " elements";' 60000 30000).Job.value.text; $c })) -join "") }
  tasks = @(); throughput = @(); mismatches = @() }

# ---- per-task latency -------------------------------------------------------------------------------------------
$taskNames = Get-ChildItem $Tasks -Filter "*.cs" | ForEach-Object { $_.BaseName } | Where-Object { $_ -ne "noop" -or $true } | Sort-Object
if ($Only.Count) { $taskNames = $taskNames | Where-Object { $Only -contains $_ } }
foreach ($t in $taskNames) {
  $values = @{}
  foreach ($e in $engines) {
    $src = Join-Path $Tasks "$t.$($e.ext)"
    if (-not (Test-Path $src)) { continue }
    $code = [IO.File]::ReadAllText($src)
    $r = Post $e.name $code 300000 600000   # warm-up 1 (also loads the engine / JITs)
    if ($r.Job.state -ne "done") { Write-Output "SKIP $t $($e.name): $($r.Job.state) $($r.Job.error.message)"; $result.tasks += [ordered]@{ task = $t; engine = $e.name; error = "$($r.Job.state): $($r.Job.error.message)" }; continue }
    $null = Post $e.name $code 300000 600000
    $rows = @(1..$Runs | ForEach-Object { $r = Post $e.name $code 300000 600000; if ($r.Job.state -eq "done") { Split-Times $r.Job $r.Wall } })
    $values[$e.name] = if ($null -ne $r.Job.value.json) { ($r.Job.value.json | ConvertTo-Json -Compress -Depth 4) } else { "$($r.Job.value.text)" }
    $result.tasks += [ordered]@{ task = $t; engine = $e.name; runs = $rows.Count; value = $values[$e.name]; outputLength = $r.Job.outputLength
      wall = Stats ($rows | ForEach-Object wall); compile = Stats ($rows | ForEach-Object compile); dispatch = Stats ($rows | ForEach-Object dispatch); run = Stats ($rows | ForEach-Object run) }
    "{0,-12} {1,-12} wall {2,9} ms   compile {3,8}   dispatch {4,7}   run {5,9}   -> {6}" -f $t, $e.name, (Stats ($rows | ForEach-Object wall)).median, (Stats ($rows | ForEach-Object compile)).median, (Stats ($rows | ForEach-Object dispatch)).median, (Stats ($rows | ForEach-Object run)).median, $values[$e.name]
  }
  $distinct = @($values.Values | Select-Object -Unique)
  if ($distinct.Count -gt 1) { $m = "$t : " + (($values.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join " | "); $result.mismatches += $m; "MISMATCH $m" }
}

# ---- throughput -----------------------------------------------------------------------------------------------------
foreach ($e in $engines) {
  $code = [IO.File]::ReadAllText((Join-Path $Tasks "noop.$($e.ext)"))
  $null = Post $e.name $code 60000 60000
  $sw = [Diagnostics.Stopwatch]::StartNew()
  1..$Burst | ForEach-Object { $null = Post $e.name $code 60000 60000 }
  $sw.Stop(); $seqPerSec = $Burst / $sw.Elapsed.TotalSeconds
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $ids = 1..$Burst | ForEach-Object { (Post $e.name $code 60000 0).Job.id }
  $submitMs = $sw.Elapsed.TotalMilliseconds
  foreach ($id in $ids) { $null = GetJob $id 120000 }
  $sw.Stop(); $burstPerSec = $Burst / $sw.Elapsed.TotalSeconds
  $result.throughput += [ordered]@{ engine = $e.name; jobs = $Burst; sequentialJobsPerSec = [math]::Round($seqPerSec, 1); burstJobsPerSec = [math]::Round($burstPerSec, 1); burstSubmitMs = [math]::Round($submitMs) }
  "{0,-12} throughput: sequential {1,6:N1} jobs/s   burst {2,6:N1} jobs/s (submit {3} ms)" -f $e.name, $seqPerSec, $burstPerSec, [math]::Round($submitMs)
}
$rev.Refresh(); $result.meta.revitWorkingSetMB = [math]::Round($rev.WorkingSet64 / 1MB)
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
$result | ConvertTo-Json -Depth 8 | Set-Content $Out -Encoding UTF8
"written $Out"
