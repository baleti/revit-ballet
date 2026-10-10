$w = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
$p = Get-ItemProperty $w
"autologon: AutoAdminLogon=$($p.AutoAdminLogon) user=$($p.DefaultUserName) hasDefaultPassword=$([bool]$p.DefaultPassword)"
net user User "" | Out-Null; "net user rc=$LASTEXITCODE"
Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' -Name LimitBlankPasswordUse -Value 0 -Type DWord
if ($p.AutoAdminLogon -eq '1') { Set-ItemProperty $w -Name DefaultPassword -Value '' }
"LimitBlankPasswordUse=$((Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa').LimitBlankPasswordUse)"
