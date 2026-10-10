# Checks the installer's opt-in "remote command picker" component in the guest. Prints PASS/FAIL lines; exit code = failures.
#   -Step default | on | keep | off | uninstall
param([Parameter(Mandatory)][ValidateSet("default", "on", "keep", "off", "uninstall")][string]$Step)
$ErrorActionPreference = "Stop"
$script:fails = 0
$dir = Join-Path $env:LOCALAPPDATA "revit-ballet\remote-picker"
$runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$files = "RevitBalletRemote.exe", "RevitBallet.Core.dll", "Newtonsoft.Json.dll"
function Check($name, [scriptblock]$test) {
  try { $r = & $test; if ($r -eq $true) { "PASS $name" } else { "FAIL ${name}: $r"; $script:fails++ } }
  catch { "FAIL ${name}: exception $($_.Exception.Message)"; $script:fails++ }
}
function Installer($argList) {
  $p = Start-Process "C:\ci\installer.exe" -ArgumentList $argList -Wait -PassThru -WindowStyle Hidden
  $p.ExitCode
}
function Present() {
  $run = (Get-ItemProperty $runKey -Name RevitBalletRemote -ErrorAction SilentlyContinue).RevitBalletRemote
  $have = @($files | Where-Object { Test-Path (Join-Path $dir $_) })
  [pscustomobject]@{ Run = $run; Files = $have.Count; Proc = @(Get-Process RevitBalletRemote -ErrorAction SilentlyContinue).Count }
}

switch ($Step) {
  "default" {
    Check "default-install-does-not-add-the-picker" { $s = Present; if ($s.Run -or $s.Files -ne 0) { "run=[$($s.Run)] files=$($s.Files)" } else { $true } }
  }
  "on" {
    Check "opt-in-installs-files-and-autostart" {
      $rc = Installer "--quiet --remote-picker on"; if ($rc -ne 0) { return "installer exit $rc" }
      $s = Present
      if ($s.Files -ne 3) { return "files=$($s.Files) of 3 in $dir" }
      if ($s.Run -ne ('"' + (Join-Path $dir "RevitBalletRemote.exe") + '"')) { return "Run value=[$($s.Run)]" }
      $true
    }
    Check "installed-exe-is-the-built-one" {
      $v = (Get-Item (Join-Path $dir "RevitBalletRemote.exe")).Length; if ($v -lt 10000) { "exe is $v bytes" } else { $true }
    }
  }
  "keep" {
    Check "quiet-run-without-the-flag-keeps-an-installed-picker" {
      $rc = Installer "--quiet"; if ($rc -ne 0) { return "installer exit $rc" }
      $s = Present; if ($s.Files -ne 3 -or -not $s.Run) { "files=$($s.Files) run=[$($s.Run)]" } else { $true }
    }
  }
  "off" {
    Check "opt-out-removes-files-autostart-and-process" {
      $rc = Installer "--quiet --remote-picker off"; if ($rc -ne 0) { return "installer exit $rc" }
      Start-Sleep 2; $s = Present
      if ($s.Run) { return "Run value still [$($s.Run)]" }
      if ($s.Files -ne 0 -or (Test-Path $dir)) { return "files left: $($s.Files); dir exists=$(Test-Path $dir)" }
      if ($s.Proc -ne 0) { return "process still running ($($s.Proc))" }
      $true
    }
    Check "opt-out-leaves-the-rest-of-revit-ballet-alone" {
      $dll = Get-ChildItem "$env:APPDATA\Autodesk\Revit\Addins\2026\revit-ballet\revit-ballet.dll" -ErrorAction SilentlyContinue
      if (-not $dll) { "the add-in was removed too" } else { $true }
    }
  }
  "uninstall" {
    Check "uninstall-removes-the-picker-too" {
      $rc = Installer "--quiet --remote-picker on"; if ($rc -ne 0) { return "install exit $rc" }
      if ((Present).Files -ne 3) { return "picker not installed before the uninstall" }
      $rc = Installer "/uninstall --quiet"; if ($rc -ne 0) { return "uninstall exit $rc" }
      Start-Sleep 2; $s = Present
      if ($s.Run -or $s.Files -ne 0 -or $s.Proc -ne 0) { "run=[$($s.Run)] files=$($s.Files) proc=$($s.Proc)" } else { $true }
    }
  }
}
exit $script:fails
