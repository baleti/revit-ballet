# Print the shared agent auth token of this machine's Revit sessions (for curl from outside).
(Get-Content (Join-Path $env:APPDATA "revit-ballet\runtime\network\token") -Raw).Trim()
