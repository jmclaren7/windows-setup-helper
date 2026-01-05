#include-once
;===============================================================================
; Miscellaneous Helper Functions
;===============================================================================
#include <Array.au3>

Func _Win11Bypass()
	RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\Setup\LabConfig", "BypassSecureBootCheck", "REG_DWORD", 1)
	RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\Setup\LabConfig", "BypassTPMCheck", "REG_DWORD", 1)
	RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\Setup\LabConfig", "BypassCPUCheck", "REG_DWORD", 1)
	RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\Setup\LabConfig", "BypassRAMCheck", "REG_DWORD", 1)
	RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\Setup\LabConfig", "BypassStorageCheck", "REG_DWORD", 1)
	RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\Setup\MoSetup", "AllowUpgradesWithUnsupportedTPMOrCPU", "REG_DWORD", 1)
	RegWrite("HKEY_CURRENT_USER\Control Panel\UnsupportedHardwareNotificationCache", "SV1", "REG_DWORD", 0)
	RegWrite("HKEY_CURRENT_USER\Control Panel\UnsupportedHardwareNotificationCache", "SV2", "REG_DWORD", 0)
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE", "BypassNRO", "REG_DWORD", 1)
EndFunc   ;==>_Win11Bypass

Func _GetDisks()
	Local $WMI = ObjGet('winmgmts:\\.\root\cimv2')

	Local $Partitions[0][3]
	Local $Query = $WMI.ExecQuery("Select DeviceID,Bootable,Type From Win32_DiskPartition")
	For $Item In $Query
		$Index = UBound($Partitions)
		ReDim $Partitions[$Index + 1][UBound($Partitions, 2)]
		$Partitions[$Index][0] = $Item.DeviceID
		$Partitions[$Index][0] = StringReplace(StringMid($Partitions[$Index][0], StringInStr($Partitions[$Index][0], "#") + 1, 2), ",", "")
		$Partitions[$Index][1] = $Item.Bootable
		$Partitions[$Index][2] = $Item.Type
	Next

	Local $Disks[0][5]
	Local $Query = $WMI.ExecQuery("Select DeviceID,Caption,InterfaceType,Size From Win32_DiskDrive")
	For $Item In $Query
		$Index = UBound($Disks)
		ReDim $Disks[$Index + 1][UBound($Disks, 2)]
		$Disks[$Index][0] = $Item.DeviceID
		$Disks[$Index][0] = StringMid($Disks[$Index][0], StringInStr($Disks[$Index][0], "DRIVE") + 5, 2)
		$Disks[$Index][1] = $Item.Caption
		$Disks[$Index][2] = $Item.InterfaceType
		$Disks[$Index][3] = Round($Item.Size / (1024 ^ 3)) & "GB"
		$Disks[$Index][4] = UBound(_ArrayFindAll($Partitions, $Disks[$Index][0], Default, Default, Default, Default, Default, 0))
	Next

	_ArraySort($Disks)

	Return $Disks
EndFunc   ;==>_GetDisks