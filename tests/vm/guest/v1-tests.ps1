# End-to-end checks of the /v1 agent API (docs/agent-protocol.md) against a running Revit. Prints PASS/FAIL lines; exit code = failures.
param([Parameter(Mandatory)][int]$Port, [Parameter(Mandatory)][int]$Year)
$ErrorActionPreference = "Stop"
$tok = (Get-Content (Join-Path $env:APPDATA "revit-ballet\runtime\network\token") -Raw).Trim()
$base = "https://127.0.0.1:$Port"
$tmp = Join-Path $env:TEMP "v1-body.json"
$script:fails = 0

function Api($method, $path, $body, [bool]$auth = $true, [int]$max = 60, [string[]]$extra = @()) {
  $args = @("-sk", "-m", "$max", "-X", $method, "-w", "`n%{http_code}") + $extra
  if ($auth) { $args += @("-H", "X-Auth-Token: $tok") }
  if ($null -ne $body) {
    $json = if ($body -is [string]) { $body } else { $body | ConvertTo-Json -Compress -Depth 6 }
    [IO.File]::WriteAllText($tmp, $json, (New-Object Text.UTF8Encoding($false)))
    $args += @("-H", "Content-Type: application/json", "--data-binary", "@$tmp")
  }
  $out = & curl.exe @args "$base$path"
  $lines = @($out)
  $status = [int]$lines[-1]
  $text = ($lines[0..($lines.Count - 2)] -join "`n")
  $obj = $null; try { $obj = $text | ConvertFrom-Json } catch { }
  [pscustomobject]@{ Status = $status; Body = $obj; Text = $text }
}
function Check($name, [scriptblock]$test) {
  try { $r = & $test; if ($r -eq $true) { "PASS $name" } else { "FAIL ${name}: $r"; $script:fails++ } }
  catch { "FAIL ${name}: exception $($_.Exception.Message)"; $script:fails++ }
}
function Run($engine, $code, $timeoutMs = 20000, $session = $null) {
  $b = @{ engine = $engine; code = $code; timeoutMs = $timeoutMs; wait = 30000 }
  if ($session) { $b.session = $session }
  (Api POST "/v1/jobs" $b).Body
}

Check "info" {
  $i = (Api GET "/v1/info").Body
  if ($i.protocol -ne 1) { return "protocol=$($i.protocol)" }
  if ($i.revit.year -ne $Year) { return "year=$($i.revit.year)" }
  $r = $i.engines | Where-Object name -eq "roslyn"
  if (-not $r.available) { return "roslyn not available: $($r.reason)" }
  if (-not $i.ui.responsive) { return "ui not responsive" }
  $true
}
Check "auth-required" { $r = Api GET "/v1/info" $null $false; if ($r.Status -eq 401 -and $r.Body.error.kind -eq "auth") { $true } else { "status=$($r.Status) $($r.Text)" } }
Check "job-value-and-output" {
  $j = Run roslyn 'Console.WriteLine("hi"); Out.WriteLine("there"); return 6*7;'
  if ($j.state -ne "done") { return "state=$($j.state) $($j.error.message)" }
  if ($j.output -ne "hi`nthere`n" -and $j.output -ne "hi`r`nthere`r`n") { return "output=[$($j.output)]" }
  if ($j.value.json -ne 42) { return "value=$($j.value | ConvertTo-Json -Compress)" }
  $true
}
Check "compile-error" {
  $j = Run roslyn 'int x = ;'
  if ($j.state -ne "failed" -or $j.error.kind -ne "compile") { return "state=$($j.state) kind=$($j.error.kind)" }
  $d = @($j.diagnostics | Where-Object severity -eq "error")[0]
  if ($d.line -ne 1) { return "diagnostic line=$($d.line)" }
  $true
}
Check "runtime-error" {
  $j = Run roslyn 'throw new System.InvalidOperationException("boom");'
  if ($j.state -ne "failed" -or $j.error.kind -ne "runtime" -or $j.error.message -notmatch "boom") { return "state=$($j.state) $($j.error | ConvertTo-Json -Compress)" }
  $true
}
Check "revit-context" {
  $j = Run roslyn 'return UIApp.Application.VersionNumber;'
  if ($j.state -ne "done" -or $j.value.text -ne "$Year") { return "state=$($j.state) value=$($j.value.text) $($j.error.message)" }
  $true
}
Check "non-ascii-body" {
  $j = Run roslyn ('return "' + [char]0x00e9 + [char]0x20ac + '".Length;')
  if ($j.state -ne "done" -or $j.value.json -ne 2) { return "state=$($j.state) value=$($j.value.json) $($j.error.message)" }
  $true
}
Check "session-state" {
  $a = Run roslyn 'var counter = 5;' 20000 "t1"
  $b = Run roslyn 'return counter + 1;' 20000 "t1"
  if ($b.state -ne "done" -or $b.value.json -ne 6) { return "continued job: state=$($b.state) value=$($b.value.json) $($b.error.message)" }
  $s = (Api GET "/v1/sessions").Body.sessions | Where-Object name -eq "t1"
  if (-not $s) { return "session not listed" }
  $d = Api DELETE "/v1/sessions/roslyn/t1"
  if ($d.Status -ne 200) { return "delete status=$($d.Status)" }
  $c = Run roslyn 'return counter;' 20000 "t1"
  if ($c.state -ne "failed") { return "state after drop=$($c.state)" }
  $true
}
Check "queue-keeps-order-no-lost-jobs" {
  $ids = 1..5 | ForEach-Object { (Api POST "/v1/jobs" @{ engine = "roslyn"; code = "System.Threading.Thread.Sleep(150); return $_;"; timeoutMs = 20000 }).Body.id }
  $vals = foreach ($id in $ids) { (Api GET "/v1/jobs/${id}?wait=30000").Body }
  $bad = $vals | Where-Object { $_.state -ne "done" }
  if ($bad) { return "not done: $(($vals | ForEach-Object state) -join ',')" }
  $got = ($vals | ForEach-Object { $_.value.json }) -join ","
  if ($got -ne "1,2,3,4,5") { return "values=$got" }
  $true
}
Check "cancel-queued" {
  $first = (Api POST "/v1/jobs" @{ engine = "roslyn"; code = "System.Threading.Thread.Sleep(2500); return 1;"; timeoutMs = 20000 }).Body
  $second = (Api POST "/v1/jobs" @{ engine = "roslyn"; code = "return 2;"; timeoutMs = 20000 }).Body
  $c = Api DELETE "/v1/jobs/$($second.id)"
  if ($c.Body.state -ne "cancelled") { return "state=$($c.Body.state)" }
  $f = (Api GET "/v1/jobs/$($first.id)?wait=30000").Body
  if ($f.state -ne "done") { return "first job state=$($f.state)" }
  $true
}
Check "timeout-and-cancel-token" {
  $j = Run roslyn 'while (!Cancel.IsCancellationRequested) { System.Threading.Thread.Sleep(20); } return 0;' 1500
  if ($j.state -ne "timeout") { return "state=$($j.state)" }
  $n = Run roslyn 'return 1;'
  if ($n.state -ne "done") { return "next job state=$($n.state) $($n.error.message)" }
  $true
}
Check "events-stream" {
  $b = (Api POST "/v1/jobs" @{ engine = "roslyn"; code = 'Out.WriteLine("streamed"); return 1;'; timeoutMs = 20000 }).Body
  $out = (& curl.exe -sk -m 30 -H "X-Auth-Token: $tok" "$base/v1/jobs/$($b.id)/events") -join "`n"
  if ($out -notmatch "event: done" -or $out -notmatch "streamed") { return $out }
  $true
}
Check "unknown-engine" { $r = Api POST "/v1/jobs" @{ engine = "nope"; code = "1" }; if ($r.Status -eq 404) { $true } else { "status=$($r.Status)" } }
Check "legacy-roslyn-shim" {
  [IO.File]::WriteAllText($tmp, 'Console.WriteLine("legacy"); return 3;', (New-Object Text.UTF8Encoding($false)))
  $o = & curl.exe -sk -m 60 -H "X-Auth-Token: $tok" --data-binary "@$tmp" "$base/roslyn"
  $r = $o | ConvertFrom-Json
  if (-not $r.Success -or $r.Output -notmatch "legacy" -or $r.Output -notmatch "Return value: 3") { return $o }
  $true
}


# ---- correctness batch: hardening, caps, audit log, idempotency, liveness fields ----
Check "origin-and-host-rejected" {
  $o = Api GET "/v1/info" $null $true 30 @("-H", "Origin: http://evil.example")
  $h = Api GET "/v1/info" $null $true 30 @("-H", "Host: evil.example")
  $ok = Api GET "/v1/info"
  if ($o.Status -ne 403) { return "Origin header: status=$($o.Status)" }
  if ($h.Status -ne 403) { return "foreign Host: status=$($h.Status)" }
  if ($ok.Status -ne 200) { return "normal request: status=$($ok.Status)" }
  $true
}
Check "ui-state-and-queue-age" {
  $i = (Api GET "/v1/info").Body
  if ($i.ui.state -ne "idle") { return "ui.state=$($i.ui.state)" }
  if ($null -eq $i.queue.oldestQueuedMs -or $i.queue.oldestQueuedMs -ne 0) { return "oldestQueuedMs=$($i.queue.oldestQueuedMs)" }
  $true
}
Check "output-cap-and-output-endpoint" {
  $j = Run roslyn 'for (int i = 0; i < 40000; i++) Out.WriteLine(new string((char)120, 20)); return 1;'
  if ($j.state -ne "done") { return "state=$($j.state) $($j.error.message)" }
  if (-not $j.outputTruncated) { return "not truncated, length=$($j.outputLength)" }
  if ($j.output.Length -gt 262144) { return "inline output too long: $($j.output.Length)" }
  if ($j.outputLength -lt 800000) { return "outputLength=$($j.outputLength)" }
  $s = (Api GET "/v1/jobs/$($j.id)/output?offset=600000&limit=50").Body
  if ($s.length -ne 50 -or $s.total -ne $j.outputLength -or $s.text -notmatch "^x+") { return "slice: length=$($s.length) total=$($s.total)" }
  $true
}
Check "value-revit-types-roslyn" {
  $a = Run roslyn 'return new XYZ(1, 2, 3);'
  if ($a.state -ne "done" -or $a.value.json.x -ne 1 -or $a.value.json.z -ne 3) { return "XYZ: $($a | ConvertTo-Json -Compress -Depth 5)" }
  $b = Run roslyn 'return new ElementId(5);'
  if ($b.state -ne "done" -or $b.value.json.id -ne 5) { return "ElementId: $($b.value | ConvertTo-Json -Compress)" }
  $c = Run roslyn 'return new List<object> { 1, "a", new XYZ(0, 0, 1) };'
  if ($c.state -ne "done" -or @($c.value.json).Count -ne 3) { return "list: $($c.value | ConvertTo-Json -Compress -Depth 5)" }
  $true
}
Check "audit-log-client-and-request-id" {
  $b = @{ engine = "roslyn"; code = 'return 11;'; wait = 20000; requestId = "ci-req-1" }
  $r1 = Api POST "/v1/jobs" $b $true 60 @("-H", "X-Agent-Name: ci-test")
  $r2 = Api POST "/v1/jobs" $b $true 60 @("-H", "X-Agent-Name: ci-test")
  if ($r1.Body.id -ne $r2.Body.id) { return "retry created a second job: $($r1.Body.id) vs $($r2.Body.id)" }
  if ($r1.Body.client -ne "ci-test") { return "client=$($r1.Body.client)" }
  $dir = Join-Path $env:APPDATA "revit-ballet\runtime\agent\jobs"
  $f = Get-ChildItem $dir -Filter "*-$($r1.Body.id).json" -ErrorAction SilentlyContinue | Select-Object -Last 1
  if (-not $f) { return "no audit file for $($r1.Body.id) in $dir" }
  $j = Get-Content $f.FullName -Raw | ConvertFrom-Json
  if ($j.code -ne "return 11;" -or $j.client -ne "ci-test" -or $j.state -ne "done") { return "audit content: $($j | ConvertTo-Json -Compress)" }
  if ((Get-Content $f.FullName -Raw) -match [regex]::Escape($tok)) { return "token leaked into the audit log" }
  $true
}

# ---- Python engines (each only when installed); the same snippets are valid Python 2 and 3 ----
$infoNow = (Api GET "/v1/info").Body
foreach ($eng in "ironpython2", "cpython3") {
$avail = $infoNow.engines | Where-Object { $_.name -eq $eng -and $_.available }
if ($avail) {
  Check "$eng-value-and-output" {
    $j = Run $eng "print('hi')`nout.write('there\n')`nresult = 6*7"
    if ($j.state -ne "done") { return "state=$($j.state) $($j.error.message)" }
    if (($j.output -replace "`r","") -ne "hi`nthere`n") { return "output=[$($j.output)]" }
    if ($j.value.json -ne 42) { return "value=$($j.value | ConvertTo-Json -Compress)" }
    $true
  }
  Check "$eng-revit-api" {
    $j = Run $eng "import clr`nclr.AddReference('RevitAPI')`nfrom Autodesk.Revit.DB import FilteredElementCollector, Wall`nresult = uiapp.Application.VersionNumber"
    if ($j.state -ne "done" -or $j.value.text -ne "$Year") { return "state=$($j.state) value=$($j.value.text) $($j.error.message)" }
    $true
  }
  Check "$eng-import-star-without-addreference" {
    $j = Run $eng "from Autodesk.Revit.DB import *`nresult = Document.__name__"
    if ($j.state -ne "done" -or $j.value.text -ne "Document") { return "state=$($j.state) value=$($j.value.text) $($j.error.message)" }
    $true
  }
  Check "$eng-readme-example-documents" {
    $j = Run $eng "result = [d.Title for d in uiapp.Application.Documents]"
    if ($j.state -ne "done") { return "state=$($j.state) $($j.error.message)" }
    $true
  }
  Check "$eng-value-revit-types" {
    $j = Run $eng "import clr`nclr.AddReference('RevitAPI')`nfrom Autodesk.Revit.DB import XYZ`nresult = XYZ(1, 2, 3)"
    if ($j.state -ne "done" -or $j.value.json.x -ne 1 -or $j.value.json.y -ne 2) { return "state=$($j.state) value=$($j.value | ConvertTo-Json -Compress -Depth 4) $($j.error.message)" }
    $true
  }
  Check "$eng-stdlib" {
    $j = Run $eng "import os, re, json, collections`nresult = json.dumps({'a': 1})"
    if ($j.state -ne "done" -or $j.value.text -ne '{"a": 1}') { return "state=$($j.state) value=$($j.value.text) $($j.error.message)" }
    $true
  }
  Check "$eng-syntax-error" {
    $j = Run $eng "x = ("
    if ($j.state -ne "failed" -or $j.error.kind -ne "compile") { return "state=$($j.state) kind=$($j.error.kind)" }
    if (-not $j.diagnostics -or $j.diagnostics[0].line -lt 1) { return "no line in diagnostics" }
    $true
  }
  Check "$eng-runtime-error" {
    $j = Run $eng "def f():`n    raise ValueError('bad')`nf()"
    if ($j.state -ne "failed" -or $j.error.kind -ne "runtime" -or $j.error.message -notmatch "bad" -or $j.error.stack -notmatch "line 2") { return "state=$($j.state) $($j.error | ConvertTo-Json -Compress)" }
    $true
  }
  Check "$eng-session-state" {
    $null = Run $eng "x = 5" 20000 "p1"
    $b = Run $eng "result = x + 1" 20000 "p1"
    $null = Api DELETE "/v1/sessions/$eng/p1"
    if ($b.state -ne "done" -or $b.value.json -ne 6) { return "state=$($b.state) value=$($b.value.json) $($b.error.message)" }
    $true
  }
  Check "$eng-cancel-token-and-hot-loop" {
    $a = Run $eng "while not cancel.is_set():`n    pass" 1500
    if ($a.state -ne "timeout") { return "cooperative loop: state=$($a.state)" }
    $b = Run $eng "while True:`n    pass" 1500
    if ($b.state -ne "timeout") { return "hot loop: state=$($b.state)" }
    $n = Run $eng "result = 1"
    if ($n.state -ne "done") { return "next job state=$($n.state) $($n.error.message)" }
    $true
  }
  Check "$eng-mixed-queue-with-roslyn" {
    $ids = @(); foreach ($i in 1..3) {
      $ids += (Api POST "/v1/jobs" @{ engine = "roslyn"; code = "return $i;" }).Body.id
      $ids += (Api POST "/v1/jobs" @{ engine = $eng; code = "result = $i" }).Body.id
    }
    $vals = foreach ($id in $ids) { (Api GET "/v1/jobs/${id}?wait=30000").Body }
    if (($vals | Where-Object state -ne "done")) { return ($vals | ForEach-Object state) -join "," }
    $got = ($vals | ForEach-Object { $_.value.json }) -join ","
    if ($got -ne "1,1,2,2,3,3") { return "values=$got" }
    $true
  }
} else { "SKIP $eng engine not installed" }
}
exit $script:fails
