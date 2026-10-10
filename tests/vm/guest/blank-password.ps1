$w = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
$p = Get-ItemProperty $w
"autologon: AutoAdminLogon=$($p.AutoAdminLogon) user=$($p.DefaultUserName) hasDefaultPassword=$([bool]$p.DefaultPassword)"
Set-LocalUser -Name User -Password (New-Object System.Security.SecureString)   # "net user User ``" over ssh silently does nothing
"blank password set"
Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' -Name LimitBlankPasswordUse -Value 0 -Type DWord
if ($p.AutoAdminLogon -eq '1') { Set-ItemProperty $w -Name DefaultPassword -Value '' }
"LimitBlankPasswordUse=$((Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa').LimitBlankPasswordUse)"
Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name UserAuthentication -Value 1 -Type DWord   # NLA stays on
