# Installer --engines flag: none / a subset / a different subset / all. Run when Revit is not running. Prints PASS/FAIL lines.
param([Parameter(Mandatory)][int]$Year, [string]$Installer = "C:\ci\installer.exe")
$dir = Join-Path $env:APPDATA "Autodesk\Revit\Addins\$Year\revit-ballet\engines"
$script:fails = 0
function Install($engines) {
  $p = Start-Process -FilePath $Installer -ArgumentList "/quiet", "--engines", $engines -Wait -PassThru -RedirectStandardOutput C:\ci\out\eng.out -RedirectStandardError C:\ci\out\eng.err
  $p.ExitCode
}
function Installed { @(Get-ChildItem $dir -Directory -ErrorAction SilentlyContinue | ForEach-Object Name | Sort-Object) -join "," }
function Case($name, $arg, $expect) {
  $rc = Install $arg; $got = Installed
  if ($rc -eq 0 -and $got -eq $expect) { "PASS engines-flag $name" } else { "FAIL engines-flag ${name}: exit=$rc installed=[$got] expected=[$expect]"; $script:fails++ }
}
Case "none" "none" ""
Case "roslyn-only" "roslyn" "roslyn"
Case "python-only" "ironpython2,cpython3" "cpython3,ironpython2"
Case "all" "all" "cpython3,ironpython2,roslyn"
exit $script:fails
