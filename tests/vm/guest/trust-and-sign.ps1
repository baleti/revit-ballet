# CI stand-in for a real code-signing certificate: make a throwaway self-signed one, trust it machine-wide and sign the
# installed add-in DLLs, so Revit loads them without the "Security - Unsigned Add-In" prompt. Nothing leaves the VM.
param([Parameter(Mandatory)][string[]]$Years)
$Years = @($Years | ForEach-Object { $_ -split "," })
$ErrorActionPreference = "Stop"
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=Revit Ballet CI $(Get-Date -Format yyyyMMddHHmmss)" -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddDays(2)
$cer = "C:\ci\out\ci-signer.cer"
New-Item -ItemType Directory -Force C:\ci\out | Out-Null
Export-Certificate -Cert $cert -FilePath $cer | Out-Null
Import-Certificate -FilePath $cer -CertStoreLocation Cert:\LocalMachine\Root | Out-Null
Import-Certificate -FilePath $cer -CertStoreLocation Cert:\LocalMachine\TrustedPublisher | Out-Null
$bad = 0
foreach ($y in $Years) {
  $dll = Join-Path $env:APPDATA "Autodesk\Revit\Addins\$y\revit-ballet\revit-ballet.dll"
  if (-not (Test-Path $dll)) { Write-Output "FAIL $y : $dll missing"; $bad++; continue }
  $r = Set-AuthenticodeSignature -FilePath $dll -Certificate $cert -HashAlgorithm SHA256
  Write-Output ("{0} {1} : signature {2}" -f $(if ($r.Status -eq "Valid") { "OK  " } else { "FAIL" }), $y, $r.Status)
  if ($r.Status -ne "Valid") { $bad++ }
}
exit $bad
