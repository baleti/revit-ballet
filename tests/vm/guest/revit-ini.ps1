# Per-year Revit.ini fixes so no modal dialog can block the UI thread (and with it every /roslyn script):
#  - FamilyTemplatePath empty (Revit 2020 in the baseline has no family template content): point it at an existing
#    directory, otherwise Revit opens "Default Family Template File Invalid" at start-up.
#  - -NoGpu (CI runs without GPU passthrough): UseGraphicsHardware=0, otherwise "Hardware Acceleration disabled".
param([Parameter(Mandatory)][string[]]$Years, [switch]$NoGpu)
$Years = @($Years | ForEach-Object { $_ -split "," })
New-Item -ItemType Directory -Force C:\ci\templates | Out-Null
foreach ($y in $Years) {
  $dir = Join-Path $env:APPDATA "Autodesk\Revit\Autodesk Revit $y"
  New-Item -ItemType Directory -Force $dir | Out-Null
  $ini = Join-Path $dir "Revit.ini"
  $text = if (Test-Path $ini) { Get-Content $ini -Raw } else { "" }
  $notes = @()
  $re = [regex]'(?m)^FamilyTemplatePath=[ \t]*\r?$'
  if ($re.IsMatch($text)) { $text = $re.Replace($text, "FamilyTemplatePath=C:\ci\templates\`r", 1); $notes += "FamilyTemplatePath" }
  if ($NoGpu) {
    if ($text -match '(?m)^UseGraphicsHardware=') { $text = $text -replace '(?m)^UseGraphicsHardware=\d', 'UseGraphicsHardware=0' }
    elseif ($text -match '(?m)^\[Graphics\]') { $text = $text -replace '(?m)^\[Graphics\]', "[Graphics]`r`nUseGraphicsHardware=0" }
    else { $text += "`r`n[Graphics]`r`nUseGraphicsHardware=0`r`n" }
    $notes += "UseGraphicsHardware=0"
  }
  Set-Content -Path $ini -Value $text -Encoding ASCII -NoNewline
  Write-Output ("OK   {0} : {1}" -f $y, $(if ($notes) { $notes -join ", " } else { "unchanged" }))
}
