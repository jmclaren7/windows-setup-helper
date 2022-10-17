$Title = @ScriptFullPath

$Msg = MsgBox(1, $Title, "Use the ""EndTask"" button in the top left to exit without rebooting.")
If $Msg <> 1 Then Exit

$Pid = ShellExecute("refixer.exe", "", @ScriptDir)


$CloseForm = GUICreate("", 52, 18, 0, 0, 0x80000000)
$ForceCloseButton = GUICtrlCreateButton("EndTask", 0, 0, 52, 18)
GUISetState(@SW_SHOW, $CloseForm)
WinSetOnTop ($CloseForm, "", 1)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $ForceCloseButton
			$Msg = MsgBox(4, $Title, "Are you sure you want to exit?"&@LF&@LF&"Note: you can also use alt+tab to switch to other windows.")
			If $Msg <> 6 Then ContinueLoop ;6=Yes
			ProcessClose($Pid)

	EndSwitch

	If NOT ProcessExists($Pid) Then
		Exit
	EndIf
	Sleep(20)
Wend