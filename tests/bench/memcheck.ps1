# Working set of Revit before/after a forced garbage collection (run through the agent API), to tell a leak from uncollected garbage.
param([Parameter(Mandatory)][int]$Port)
Add-Type -AssemblyName System.Net.Http
Add-Type @"
using System.Net;
public static class TrustAll { public static void Enable() { ServicePointManager.ServerCertificateValidationCallback = (a, b, c, d) => true; } }
"@
[TrustAll]::Enable(); [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$tok = (Get-Content (Join-Path $env:APPDATA "revit-ballet\runtime\network\token") -Raw).Trim()
$client = New-Object System.Net.Http.HttpClient; $client.Timeout = [TimeSpan]::FromMinutes(5); $client.DefaultRequestHeaders.Add("X-Auth-Token", $tok)
function Job($code) { $body = @{ engine = "roslyn"; code = $code; timeoutMs = 240000; wait = 240000 } | ConvertTo-Json -Compress
  ($client.PostAsync("https://127.0.0.1:$Port/v1/jobs", (New-Object System.Net.Http.StringContent($body, [Text.Encoding]::UTF8, "application/json"))).Result.Content.ReadAsStringAsync().Result | ConvertFrom-Json) }
$rev = Get-Process Revit | Select-Object -First 1
function Mb { $rev.Refresh(); [math]::Round($rev.WorkingSet64 / 1MB) }
"working set now:            $(Mb) MB"
$a = Job 'return System.GC.GetTotalMemory(false) / 1048576;'
"managed heap (no GC):       $($a.value.text) MB   state=$($a.state)"
$b = Job 'System.GC.Collect(); System.GC.WaitForPendingFinalizers(); System.GC.Collect(); return System.GC.GetTotalMemory(true) / 1048576;'
"managed heap after full GC: $($b.value.text) MB   state=$($b.state)"
Start-Sleep 3
"working set after GC:       $(Mb) MB"
$c = Job 'return AppDomain.CurrentDomain.GetAssemblies().Length + " assemblies loaded, " + AppDomain.CurrentDomain.GetAssemblies().Count(x => x.IsDynamic) + " dynamic";'
"loaded assemblies:          $($c.value.text)"
