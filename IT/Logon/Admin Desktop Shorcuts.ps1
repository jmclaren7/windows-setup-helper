$WshShell = New-Object -comObject WScript.Shell

$Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Domain & Name.lnk")
$Shortcut.TargetPath = "SystemPropertiesComputerName.exe"
$Shortcut.Save()

$Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Users Manager.lnk")
$Shortcut.TargetPath = "lusrmgr.msc"
$Shortcut.Save()