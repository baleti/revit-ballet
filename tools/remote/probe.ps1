# Runs ON a target machine (sent over ssh, nothing is installed there): finds every Revit Ballet agent on this machine and lists its commands.
# Prints ONE line of JSON: { hostname, error, instances: [ {port,pid,year,version,build,state,responsive,dialog,queued,running,loadedFrom,protocol,hungCompiles,activeEngine,engines} ], commands: { "<year>": [ {name,displayName,input,output} ] } }
$ErrorActionPreference = 'Stop'
$out = [ordered]@{ hostname = $env:COMPUTERNAME; error = $null; instances = @(); commands = [ordered]@{} }
function Emit { $out | ConvertTo-Json -Depth 8 -Compress; exit 0 }
$tokFile = Join-Path $env:APPDATA 'revit-ballet\runtime\network\token'
if (-not (Test-Path $tokFile)) { $out.error = 'no agent token on this machine (Revit Ballet has not run here)'; Emit }
$tok = (Get-Content $tokFile -Raw).Trim()
function Api($port, $path) {
  $j = & curl.exe -sk -m 8 -H "X-Auth-Token: $tok" "https://127.0.0.1:$port$path" 2>$null
  if ($j) { try { $j | ConvertFrom-Json } catch { $null } } else { $null }
}
$ports = @(Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Where-Object { $_.LocalPort -ge 23717 -and $_.LocalPort -le 23817 } | ForEach-Object { $_.LocalPort } | Sort-Object -Unique)
$instances = @()
foreach ($p in $ports) {
  $i = Api $p '/v1/info'
  if (-not $i -or -not $i.revit) { continue }
  $year = [string]$i.revit.year
  if ($P.commands -ne $false -and -not $out.commands.Contains($year)) {
    $c = Api $p '/v1/commands'
    $out.commands[$year] = if ($c -and $c.commands) { @($c.commands | ForEach-Object { [ordered]@{ name = $_.name; displayName = $_.displayName; input = $_.input; output = $_.output } }) } else { @() }
  }
  $dlg = if ($i.ui.modalDialog) { [string]$i.ui.modalDialog.title } else { $null }
  # the engine of the job running right now (queue.running is its id); null when nothing is running
  $active = $null
  if ($i.queue.running) { $j = Api $p ('/v1/jobs/' + $i.queue.running); if ($j -and $j.engine) { $active = [string]$j.engine } }
  $instances += [ordered]@{
    port = [int]$p; pid = [int]$i.session.pid; year = [int]$i.revit.year; version = [string]$i.revit.version
    build = [string]$i.agent.build; state = [string]$i.ui.state; responsive = [bool]$i.ui.responsive; dialog = $dlg
    queued = [int]$i.queue.queued; running = $i.queue.running
    loadedFrom = [string]$i.agent.loadedFrom; protocol = [int]$i.protocol; hungCompiles = [int]$i.hungCompiles
    activeEngine = $active
    engines = @($i.engines | Where-Object { $_.available } | ForEach-Object { [string]$_.name })
  }
}
$out.instances = $instances
Emit
