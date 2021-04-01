$Msg = MsgBox(1, @ScriptName, "ParagonHDM causes an issue with the Windows installer, reboot if you need to run the installer."&@LF&@LF&"Use the ""EndTask"" button in the top left to exit ParagonHDM without rebooting.")
If $Msg <> 1 Then Exit

$SystemDrive = Stringleft(@SystemDir, 2)
$Pid = ShellExecute($SystemDrive & "\Programs\Paragon Software\program\hdm17.exe")

$CloseForm = GUICreate("", 52, 18, 0, 0, 0x80000000)
$ForceCloseButton = GUICtrlCreateButton("EndTask", 0, 0, 52, 18)
GUISetState(@SW_SHOW, $CloseForm)
WinSetOnTop ($CloseForm, "", 1)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $ForceCloseButton
			ProcessClose($Pid)
			Exit
	EndSwitch

	If NOT ProcessExists($Pid) Then Exit

	Sleep(20)
Wend