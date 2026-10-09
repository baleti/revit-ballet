# Run a program inside the logged-in (interactive) desktop session via a one-shot scheduled task.
# An ssh session is non-interactive: GUI apps (Revit, screenshots, mouse clicks) must be started this way.
param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Exe, [string]$Arguments = "")
$ErrorActionPreference = "Stop"
Unregister-ScheduledTask -TaskName $Name -Confirm:$false -ErrorAction SilentlyContinue
$action = if ($Arguments) { New-ScheduledTaskAction -Execute $Exe -Argument $Arguments } else { New-ScheduledTaskAction -Execute $Exe }
$principal = New-ScheduledTaskPrincipal -UserId "$env:COMPUTERNAME\$env:USERNAME" -LogonType Interactive -RunLevel Highest
Register-ScheduledTask -TaskName $Name -Action $action -Principal $principal -Force | Out-Null
Start-ScheduledTask -TaskName $Name
