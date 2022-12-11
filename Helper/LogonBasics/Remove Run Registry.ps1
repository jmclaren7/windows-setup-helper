# Remove the run at logon registry value
$RegKey =  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
$RegName = Get-Item $RegKey | Select-Object -ExpandProperty Property | Where-Object { $_ -like "*Unattend*" } 
Remove-ItemProperty $RegKey -Name $RegName