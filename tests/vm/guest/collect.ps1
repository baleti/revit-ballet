# Gather evidence for a failed run into C:\ci\out: server log, tail of the newest Revit journal per year.
New-Item -ItemType Directory -Force C:\ci\out | Out-Null
Copy-Item (Join-Path $env:APPDATA "revit-ballet\runtime\server.log") C:\ci\out\server.log -ErrorAction SilentlyContinue
Get-ChildItem (Join-Path $env:LOCALAPPDATA "Autodesk\Revit\Autodesk Revit *\Journals") -Filter "journal.*.txt" -ErrorAction SilentlyContinue |
  Sort-Object LastWriteTime -Descending | Select-Object -First 3 | ForEach-Object {
    $n = ($_.Directory.Parent.Name -replace '\s', '_') + "-" + $_.Name
    Get-Content $_.FullName -Tail 200 | Out-File "C:\ci\out\$n" -Encoding utf8 }
