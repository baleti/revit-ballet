# Create the shared auth token before several Revit sessions start at once (the first session would otherwise generate it, racing the others).
$f = Join-Path $env:APPDATA "revit-ballet\runtime\network\token"
if (-not (Test-Path $f)) {
  New-Item -ItemType Directory -Force (Split-Path $f) | Out-Null
  $bytes = New-Object byte[] 32; (New-Object Security.Cryptography.RNGCryptoServiceProvider).GetBytes($bytes)
  Set-Content -Path $f -Value (($bytes | ForEach-Object { $_.ToString("x2") }) -join "") -Encoding ASCII -NoNewline
  Write-Output "token created"
} else { Write-Output "token exists" }
