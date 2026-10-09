# POST a C# script file to the revit-ballet /roslyn endpoint and print the JSON response.
param([Parameter(Mandatory)][int]$Port, [Parameter(Mandatory)][string]$File)
$tok = (Get-Content (Join-Path $env:APPDATA "revit-ballet\runtime\network\token") -Raw).Trim()
curl.exe -sk -m 60 -H "X-Auth-Token: $tok" --data-binary "@$File" "https://127.0.0.1:$Port/roslyn"
