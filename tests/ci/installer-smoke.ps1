# Installer smoke test for a plain Windows machine WITHOUT Revit (GitHub-hosted runner): fakes just enough of a Revit install for the
# installer to detect (Revit.exe under Program Files + a per-user Uninstall entry), then checks what it lays down, the --engines flag,
# a re-install, and the uninstaller. Prints PASS/FAIL lines; exit code = number of failures. Cleans up after itself.
param([Parameter(Mandatory)][string]$Installer, [string[]]$Years = @("2022", "2025"))
$ErrorActionPreference = "Stop"
$Installer = (Resolve-Path $Installer).Path
$script:fails = 0
$created = @()
$regKeys = @()
$outDir = Join-Path ([IO.Path]::GetTempPath()) "installer-smoke"; New-Item -ItemType Directory -Force $outDir | Out-Null

function Check($name, [scriptblock]$test) {
  try { $r = & $test; if ($r -eq $true) { "PASS $name" } else { "FAIL ${name}: $r"; $script:fails++ } }
  catch { "FAIL ${name}: exception $($_.Exception.Message)"; $script:fails++ }
}
function Run-Installer([string[]]$ArgList) {
  $p = Start-Process -FilePath $Installer -ArgumentList $ArgList -Wait -PassThru -RedirectStandardOutput (Join-Path $outDir "out.txt") -RedirectStandardError (Join-Path $outDir "err.txt")
  $p.ExitCode
}
function AddinDir($y) { Join-Path $env:APPDATA "Autodesk\Revit\Addins\$y\revit-ballet" }
function EnginesInstalled($y) { @(Get-ChildItem (Join-Path (AddinDir $y) "engines") -Directory -ErrorAction SilentlyContinue | ForEach-Object Name | Sort-Object) -join "," }

try {
  # ---- fake Revit installs --------------------------------------------------------------------------
  foreach ($y in $Years) {
    $dir = Join-Path $env:ProgramFiles "Autodesk\Revit $y"
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force $dir | Out-Null; $created += (Join-Path $env:ProgramFiles "Autodesk\Revit $y") }
    $exe = Join-Path $dir "Revit.exe"
    if (-not (Test-Path $exe)) { Set-Content -Path $exe -Value "fake" }
    $key = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\RevitBalletSmoke$y"
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name DisplayName -Value "Autodesk Revit $y" -Force | Out-Null
    $regKeys += $key
  }

  # ---- install everything -----------------------------------------------------------------------------
  Check "install-exit-code" { $rc = Run-Installer @("--quiet", "--verbosity", "2"); if ($rc -eq 0) { $true } else { "exit=$rc $(Get-Content (Join-Path $outDir 'err.txt') -Raw)" } }
  foreach ($y in $Years) {
    Check "layout-$y" {
      $need = "revit-ballet.dll", "revit-ballet.agents.abstractions.dll",
              "engines\roslyn\engine.json", "engines\roslyn\revit-ballet-engine-roslyn.dll",
              "engines\ironpython2\engine.json", "engines\ironpython2\Lib.zip",
              "engines\cpython3\engine.json", "engines\cpython3\python312.dll"
      $missing = $need | Where-Object { -not (Test-Path (Join-Path (AddinDir $y) $_)) }
      if (-not (Test-Path (Join-Path $env:APPDATA "Autodesk\Revit\Addins\$y\revit-ballet.addin"))) { $missing += "revit-ballet.addin" }
      if ($missing) { "missing: $($missing -join ', ')" } else { $true }
    }
    Check "no-autodesk-binaries-shipped-$y" {
      $bad = Get-ChildItem (AddinDir $y) -Recurse -File | Where-Object { $_.Name -match '^(RevitAPI|RevitAPIUI|AdWindows|UIFramework)' }
      if ($bad) { "Autodesk assemblies in the install: $(($bad | ForEach-Object Name) -join ', ')" } else { $true }
    }
    Check "engine-manifests-$y" {
      foreach ($d in Get-ChildItem (Join-Path (AddinDir $y) "engines") -Directory) {
        $m = Get-Content (Join-Path $d.FullName "engine.json") -Raw | ConvertFrom-Json
        if ($m.name -ne $d.Name -or $m.protocol -ne 1) { return "$($d.Name): bad manifest" }
        if (-not (Test-Path (Join-Path $d.FullName $m.assembly))) { return "$($d.Name): assembly $($m.assembly) missing" }
      }
      $true
    }
  }
  Check "uninstaller-registered" { if (Test-Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\RevitBallet") { $true } else { "no HKCU uninstall entry" } }

  # ---- --engines flag -----------------------------------------------------------------------------------
  $y0 = $Years[0]
  foreach ($case in @(@("none", "none", ""), @("roslyn-only", "roslyn", "roslyn"), @("python-only", "ironpython2,cpython3", "cpython3,ironpython2"), @("all", "all", "cpython3,ironpython2,roslyn"))) {
    Check "engines-flag-$($case[0])" {
      $rc = Run-Installer @("--quiet", "--engines", $case[1]); $got = EnginesInstalled $y0
      if ($rc -eq 0 -and $got -eq $case[2]) { $true } else { "exit=$rc installed=[$got] expected=[$($case[2])]" }
    }
  }

  # ---- re-install over an existing install ----------------------------------------------------------
  Check "reinstall" { $rc = Run-Installer @("--quiet"); if ($rc -eq 0 -and (EnginesInstalled $y0) -eq "cpython3,ironpython2,roslyn") { $true } else { "exit=$rc" } }

  # ---- uninstall --------------------------------------------------------------------------------------
  Check "uninstall" {
    $rc = Run-Installer @("/uninstall", "--quiet")
    $left = @(); foreach ($y in $Years) { if (Test-Path (AddinDir $y)) { $left += "folder $y" }; if (Test-Path (Join-Path $env:APPDATA "Autodesk\Revit\Addins\$y\revit-ballet.addin")) { $left += "addin $y" } }
    if (Test-Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\RevitBallet") { $left += "uninstall registry entry" }
    if ($rc -eq 0 -and -not $left) { $true } else { "exit=$rc left behind: $($left -join ', ')" }
  }
}
finally {
  foreach ($k in $regKeys) { Remove-Item -Path $k -Recurse -Force -ErrorAction SilentlyContinue }
  foreach ($d in $created) { Remove-Item -Path $d -Recurse -Force -ErrorAction SilentlyContinue }
}
exit $script:fails
