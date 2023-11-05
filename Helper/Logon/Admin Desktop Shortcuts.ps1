$WshShell = New-Object -comObject WScript.Shell

$Shortcut = $WshShell.CreateShortcut("$Home\Desktop\System Properties.lnk")
$Shortcut.TargetPath = "SystemPropertiesComputerName.exe"
$Shortcut.Save()

$Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Users Managment.lnk")
$Shortcut.TargetPath = "lusrmgr.msc"
$Shortcut.Save()