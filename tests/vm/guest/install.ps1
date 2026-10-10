# Run the installer silently and verify what it laid down for each requested Revit year.
param([Parameter(Mandatory)][string[]]$Years, [string]$Installer = "C:\ci\installer.exe")
$Years = @($Years | ForEach-Object { $_ -split "," })
New-Item -ItemType Directory -Force C:\ci\out | Out-Null
$p = Start-Process -FilePath $Installer -ArgumentList "/quiet", "--verbosity", "2" -Wait -PassThru `
       -RedirectStandardOutput C:\ci\out\install.out -RedirectStandardError C:\ci\out\install.err
Write-Output "installer exit code: $($p.ExitCode)"
$bad = 0; if ($p.ExitCode -ne 0) { $bad++ }
foreach ($y in $Years) {
  $base = Join-Path $env:APPDATA "Autodesk\Revit\Addins\$y"
  foreach ($f in "revit-ballet.addin", "revit-ballet\revit-ballet.dll", "revit-ballet\revit-ballet.agents.abstractions.dll",
                 "revit-ballet\engines\roslyn\engine.json", "revit-ballet\engines\roslyn\revit-ballet-engine-roslyn.dll",
                 "revit-ballet\engines\ironpython2\engine.json", "revit-ballet\engines\ironpython2\Lib.zip",
                 "revit-ballet\engines\cpython3\engine.json", "revit-ballet\engines\cpython3\python312.dll") {
    $ok = Test-Path (Join-Path $base $f)
    Write-Output ("{0} {1} : {2}" -f $(if ($ok) { "OK  " } else { "FAIL" }), $y, $f)
    if (-not $ok) { $bad++ }
  }
}
if (-not (Test-Path (Join-Path $env:APPDATA "revit-ballet\uninstaller.exe"))) { Write-Output "FAIL uninstaller.exe missing"; $bad++ } else { Write-Output "OK   uninstaller.exe" }
exit $bad
