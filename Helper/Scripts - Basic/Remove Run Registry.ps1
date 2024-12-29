# Remove the run at logon registry value
# This is the run command that triggers the logon scripts to run, once it's run it's no longer needed
$RegKey =  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
$RegName = Get-Item $RegKey | Select-Object -ExpandProperty Property | Where-Object { $_ -like "*Unattend*" } 
Remove-ItemProperty $RegKey -Name $RegName