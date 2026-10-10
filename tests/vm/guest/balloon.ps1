# Install the virtio-balloon driver and the blnsvr stats service in the guest (offline: the files are copied to C:\ci\balloon by ci.sh).
# Needs the VM started with BALLOON=1 (virtio-balloon-pci). Idempotent. Prints OK / FAIL lines; the autoscaler on the host needs the
# stats service, because the Windows driver only reports memory statistics through blnsvr.
param([string]$Dir = 'C:\ci\balloon')
$ErrorActionPreference = 'Stop'
$inf = Join-Path $Dir 'balloon.inf'
if (-not (Test-Path $inf)) { 'FAIL no balloon.inf in ' + $Dir; exit 1 }
$dev = Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object { $_.InstanceId -like 'PCI\VEN_1AF4&DEV_1002*' -or $_.InstanceId -like 'PCI\VEN_1AF4&DEV_1045*' }
if (-not $dev) { 'FAIL no virtio-balloon PCI device in the guest (start the VM with BALLOON=1)'; exit 1 }
& pnputil.exe /add-driver $inf /install | Out-String | Write-Output
# blnsvr.exe -i registers the "Balloon Service" (auto start) that feeds guest memory stats to the host
Copy-Item (Join-Path $Dir 'blnsvr.exe') 'C:\Windows\System32\blnsvr.exe' -Force
if (-not (Get-Service BalloonService -ErrorAction SilentlyContinue)) { & 'C:\Windows\System32\blnsvr.exe' -i | Out-Null }
Start-Service BalloonService -ErrorAction SilentlyContinue
Start-Sleep 2
$d = Get-PnpDevice | Where-Object { $_.InstanceId -like 'PCI\VEN_1AF4&DEV_1002*' -or $_.InstanceId -like 'PCI\VEN_1AF4&DEV_1045*' }
'device: ' + (($d | ForEach-Object { $_.Status + ' ' + $_.FriendlyName }) -join '; ')
'service: ' + (Get-Service BalloonService).Status
if ($d.Status -eq 'OK' -and (Get-Service BalloonService).Status -eq 'Running') { 'OK balloon' } else { 'FAIL balloon'; exit 1 }
