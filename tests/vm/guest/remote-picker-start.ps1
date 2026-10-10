# Starts the picker test in the interactive session (a scheduled task: ssh sessions have no desktop) after stopping any running picker.
Get-Process RevitBalletRemote -ErrorAction SilentlyContinue | Stop-Process -Force
Remove-Item C:\ci\remote-picker.txt -ErrorAction SilentlyContinue
& C:\ci\it.ps1 -Name rb-remote-test -Exe powershell.exe -Arguments "-NoProfile -ExecutionPolicy Bypass -File C:\ci\remote-picker.ps1"
