# Open a model in the running Revit through the agent API (roslyn) and report how long it took.
#   powershell -File open-model.ps1 -Port 23717 -Path "C:\...\model.rvt"
param([Parameter(Mandatory)][int]$Port, [Parameter(Mandatory)][string]$Path, [int]$TimeoutSec = 600)
Add-Type -AssemblyName System.Net.Http
Add-Type @"
using System.Net;
public static class TrustAll { public static void Enable() { ServicePointManager.ServerCertificateValidationCallback = (a, b, c, d) => true; } }
"@
[TrustAll]::Enable(); [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$tok = (Get-Content (Join-Path $env:APPDATA "revit-ballet\runtime\network\token") -Raw).Trim()
$client = New-Object System.Net.Http.HttpClient; $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSec + 30); $client.DefaultRequestHeaders.Add("X-Auth-Token", $tok)
$escaped = $Path.Replace('\', '\\')
$code = 'var d = UIApp.OpenAndActivateDocument("' + $escaped + '"); return d.Document.Title + " / " + new FilteredElementCollector(d.Document).WhereElementIsNotElementType().GetElementCount() + " elements";'
$body = @{ engine = "roslyn"; code = $code; timeoutMs = ($TimeoutSec * 1000); wait = ($TimeoutSec * 1000) } | ConvertTo-Json -Compress
$sw = [Diagnostics.Stopwatch]::StartNew()
$t = $client.PostAsync("https://127.0.0.1:$Port/v1/jobs", (New-Object System.Net.Http.StringContent($body, [Text.Encoding]::UTF8, "application/json"))).Result.Content.ReadAsStringAsync().Result
$j = $t | ConvertFrom-Json
"state=$($j.state) seconds=$([math]::Round($sw.Elapsed.TotalSeconds, 1)) value=$($j.value.text) error=$($j.error.message)"
