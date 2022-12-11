#include "include\ButtonConstants.au3"
#include "include\ComboConstants.au3"
#include "include\Crypt.au3"
#include "include\EditConstants.au3"
#include "include\File.au3"
#include "include\FileConstants.au3"
#include "include\GuiConstantsEx.au3"
#include "include\GuiEdit.au3"
#include "include\GuiListView.au3"
#include "include\GuiTab.au3"
#include "include\GuiTreeView.au3"
#include "include\GuiStatusBar.au3"
#include "include\Inet.au3"
#include "include\InetConstants.au3"
#include "include\ListViewConstants.au3"
#include "include\Process.au3"
#include "include\StaticConstants.au3"
#include "include\TabConstants.au3"
#include "include\TreeViewConstants.au3"
#include "include\WindowsConstants.au3"
#include "include\WinAPI.au3"
#include "include\WinAPISys.au3"
#include "include\WinAPIFiles.au3"
#include "include\Date.au3"

#include "includeExt\Json.au3"
#include "includeExt\WinHttp.au3"
#include "includeExt\ActivationStatus.au3"
#include "includeExt\Custom.au3"
#include "includeExt\_Zip.au3"


OnAutoItExitRegister("_Exit")
_WinAPI_Wow64EnableWow64FsRedirection(False)
Opt("WinTitleMatchMode", -2)
Opt("TrayIconHide", 1)

Global $LogFullPath = StringReplace(@TempDir & "\" & @ScriptName, ".au3", ".log")
Global $MainSize = FileGetSize(@ScriptFullPath)
Global $Version = "5.0.0." & $MainSize

Global $Title = "Windows Setup Helper v" & $Version
Global $GUIMain
Global $oCommError = ObjEvent("AutoIt.Error", "_CommError")
Global $StatusBar1
Global $IsPE = StringInStr(@WindowsDir, "x")

_Log($Title)
_Log("$CmdLineRaw=" & $CmdLineRaw)

_Log("@UserName=" & @UserName)
_Log("@UserProfileDir=" & @UserProfileDir)
_Log("@AppDataDir=" & @AppDataDir)
_Log("@HomeDrive=" & @HomeDrive)
_Log("@WindowsDir=" & @WindowsDir)
_Log("@SystemDir=" & @SystemDir)
_Log("@TempDir=" & @TempDir)
_Log("@WorkingDir=" & @WorkingDir)
_Log("PATH=" & EnvGet("PATH"))

If $CmdLine[0] >= 1 Then
	$Command = $CmdLine[1]
Else
	$Command = "winpe"
EndIf

_Log("Command: " & $Command)

Switch $Command
	Case "winpe"

		#Region ### START Koda GUI section ###
		$GUIMain = GUICreate("$Title", 767, 543, -1, -1)
		$MenuItem2 = GUICtrlCreateMenu("&File")
		$MenuExitButton = GUICtrlCreateMenuItem("Exit", $MenuItem2)
		$MenuItem1 = GUICtrlCreateMenu("&Advanced")
		$MenuShowConsole = GUICtrlCreateMenuItem("Show Console", $MenuItem1)
		$MenuOpenLog = GUICtrlCreateMenuItem("Open Log File", $MenuItem1)
		$ShowHiddenTools = GUICtrlCreateMenuItem("Show Hidden Tools", $MenuItem1)
		$Tab1 = GUICtrlCreateTab(7, 4, 753, 495)
		$BootTabSheet = GUICtrlCreateTabItem("&")
		$Group5 = GUICtrlCreateGroup("WinPE Tools", 19, 37, 360, 452)
		$PERunButton = GUICtrlCreateButton("Run", 257, 454, 107, 25)
		$PEScriptTreeView = GUICtrlCreateTreeView(35, 61, 330, 385, BitOR($GUI_SS_DEFAULT_TREEVIEW,$TVS_CHECKBOXES))
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group6 = GUICtrlCreateGroup("Post-Install Scripts", 387, 37, 360, 452)
		$NormalInstallButton = GUICtrlCreateButton("Normal Install", 403, 454, 131, 25)
		$AutomatedInstallButton = GUICtrlCreateButton("Automated Install", 563, 454, 163, 25, $BS_DEFPUSHBUTTON)
		$PEInstallTreeView = GUICtrlCreateTreeView(404, 61, 330, 345, BitOR($GUI_SS_DEFAULT_TREEVIEW,$TVS_CHECKBOXES))
		$PEComputerNameInput = GUICtrlCreateInput("", 556, 422, 169, 21)
		$Label4 = GUICtrlCreateLabel("Computer Name", 476, 425, 80, 17)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		GUICtrlCreateTabItem("")
		$StatusBar1 = _GUICtrlStatusBar_Create($GUIMain)
		GUISetState(@SW_SHOW)
		#EndRegion ### END Koda GUI section ###

		; GUI Post Creation Setup
		WinSetTitle($GUIMain, "", $Title)
		GUICtrlSendMsg($PEComputerNameInput, $EM_SETCUEBANNER, False, "(Optional)")

		; Generate Script List
		_PopulateScripts($PEInstallTreeView, "Logon")
		_PopulateScripts($PEScriptTreeView, "Tools")

		Local $hSetup
		Local $RebootPrompt = False
		Local $CopyAutoLogonFiles = False
		Local $Reboot = False
		Local $BootDrive = StringLeft(@SystemDir, 3)

		; Start network
		If $IsPE Then Run(@ComSpec & " /c " & 'wpeinit.exe', @SystemDir, @SW_HIDE, $RUN_CREATE_NEW_CONSOLE)

		; Set GUI Icon
		GUISetIcon($BootDrive & "sources\setup.exe")

		; Run automatic setup scripts
		If $IsPE Then _RunMulti("PEAutoRun")

		; Hide console windows
		_Log("Hide console window")
		WinSetState($Title&" Log", "", @SW_HIDE)

		_Log("Ready", True)

		; Setup statusbar prerequisites
		$StatusBar = ""
		Global $InternetPing = 0
		AdlibRegister ("_InternetPing", 1000)
		$MemStats = MemGetStats()
		$StatusBar_Once = ""
		$StatusBar_Once &= "  |  " & Round($MemStats[1] / 1024 / 1024) & "GB"
		$StatusBar_Once &= "  |  " & _WMIC("cpu get NumberOfCores") & "/" & _WMIC("cpu get NumberOfLogicalProcessors") & " Cores"
		$Serial = _WMIC("bios get serialnumber")
		If $Serial <> "System Serial Number" Then $StatusBar_Once &= "  |  " & $Serial

		;GUI Loop
		While 1
			$nMsgA = GUIGetMsg(1)
			$nMsg = $nMsgA[0]
			Switch $nMsg
				Case $GUI_EVENT_CLOSE, $MenuExitButton
					If $nMsgA[1] <> $GUIMain Then ContinueLoop
					If $IsPE And MsgBox(1, $Title, "Closing the program will reboot the system while in WinPE.") <> 1 Then ContinueLoop
					Exit

				Case $ShowHiddenTools
					_Log("ShowHiddenTools")
					_PopulateScripts($PEScriptTreeView, "Advanced")
					$ShowHiddenTools = -1

				Case $MenuShowConsole
					WinSetState($Title&" Log", "", @SW_SHOW)

				Case $MenuOpenLog
					_Log("MenuOpenLog")
					_RunFile($LogFullPath)

				Case $PERunButton
					_RunTreeView($GUIMain, $PEScriptTreeView)

				Case $NormalInstallButton
					$hSetup = _RunFile($BootDrive & "sources\setup.exe", "/noreboot")
					$CopyAutoLogonFiles = False

				Case $AutomatedInstallButton
					If IsDeclared($hSetup) And ProcessExists($hSetup) Then
						MsgBox(0, "Error - " & $Title, "Setup is already running, please close it first")
						ContinueLoop
					EndIf

					$aAutoLogonCopy = _RunTreeView($GUIMain, $PEInstallTreeView, True)
					For $b = 0 To UBound($aAutoLogonCopy) - 1
						_Log("TreeItem: " & $aAutoLogonCopy[$b])
					Next

					$AutounattendPath = @ScriptDir & "\autounattend.xml"
					$ComputerName = GUICtrlRead($PEComputerNameInput)

					If $ComputerName <> "" Then
						_Log("$ComputerName=" & $ComputerName)
						$AutounattendPath_New = @TempDir & "\autounattend.xml"
						$sFileData = FileRead($AutounattendPath)
						$sFileData = StringReplace($sFileData, "<ComputerName>*</ComputerName>", "<ComputerName>" & $ComputerName & "</ComputerName>")
						_Log("StringReplace @extended=" & @extended)

						$hAutounattend = FileOpen($AutounattendPath_New, $FO_OVERWRITE)
						FileWrite($hAutounattend, $sFileData)
						_Log("FileWrite @error=" & @error)
						FileClose($hAutounattend)

						$AutounattendPath = $AutounattendPath_New
					EndIf

					_Log("$AutounattendPath=" & $AutounattendPath)
					$hSetup = _RunFile($BootDrive & "sources\setup.exe", "/noreboot /unattend:" & $AutounattendPath)

					$CopyAutoLogonFiles = True

			EndSwitch

			If $CopyAutoLogonFiles And Not ProcessExists($hSetup) Then
				_Log("Copy AutoLogon Files")
				For $i = 65 To 91
					$Drive = Chr($i) & ":"
					$TestFile = $Drive & "\Windows\System32\Config\SYSTEM"
					$Target = $Drive & "\Temp\Helper\"

					If FileExists($TestFile) And _FileModifiedAge($TestFile) < 600000 Then
						_Log("Found: " & $TestFile)

						; Copy the answers file so it can be used with registry key method during oobe
						; This is to deal with WDS overriding our answers file
						; Doing this also requieres that have a registry value set in the install image
						$Return = FileCopy($AutounattendPath, $Target, 1 + 8)
						_Log("FileCopy: " & $AutounattendPath & " (" & $Return & ")")

						If UBound($aAutoLogonCopy) Then
							$AutoLogonSource = StringLeft($aAutoLogonCopy[0], StringInStr($aAutoLogonCopy[0], "\", 0, -1)) & ".Autorun.ps1" ;hack because we dont know the directory
							$Return = FileCopy($AutoLogonSource, $Target, 1 + 8)
							_Log("FileCopy: " & $AutoLogonSource & " (" & $Return & ")")

						EndIf

						For $iFile = 0 To UBound($aAutoLogonCopy) - 1
							$Return = FileCopy($aAutoLogonCopy[$iFile], $Target, 1 + 8)
							_Log("FileCopy: " & $aAutoLogonCopy[$iFile] & " (" & $Return & ")")
						Next

						FileCopy($LogFullPath, $Target & "pe.log")

					EndIf
				Next
				If $i = 91 Then
					_Log("Could not find windows install")
					ContinueLoop
				EndIf

				$RebootPrompt = True
				$CopyAutoLogonFiles = False
			EndIf

			If $RebootPrompt Then
				_Log("Reboot")
				Beep(500, 1000)
				$Return = MsgBox(1 + 48 + 262144, $Title, "Rebooting in 15 seconds", 15)
				If $Return = $IDTIMEOUT Or $Return = $IDOK Then Exit
				$RebootPrompt = False
			EndIf

			; Update status bar
			$Connected = "Offline"
			If $InternetPing Then $Connected = "Online ("&$InternetPing&"ms)"
			$StatusBar_New = $Connected

			$IPAddresses = @IPAddress1 & " " & @IPAddress2 & " " & @IPAddress3 & " " & @IPAddress4
			$StatusBar_New &= "  |  " & StringStripWS(StringReplace($IPAddresses, "0.0.0.0", ""), 1+2+4)

			$StatusBar_New &= $StatusBar_Once

			If $StatusBar <> $StatusBar_New Then
				_Log("Updating statusbar")
				$StatusBar = $StatusBar_New
				_GUICtrlStatusBar_SetText($StatusBar1, $StatusBar)
			EndIf

			Sleep(1)
		WEnd

	Case Else
		_Log("Command unknown")

EndSwitch

;=========== =========== =========== =========== =========== =========== =========== ===========

Func _PopulateScripts($TreeID, $Folder)
	_Log("_PopulateScripts " & $Folder)

	Local $FileArray[0]

	_Log(@ScriptDir & "\" & $Folder & "\.Defaults.txt")
	Local $sDefaults = FileRead(@ScriptDir & "\" & $Folder & "\.Defaults.txt")
	If Not @error Then _Log("Defaults list: " & $sDefaults)

	$FileArray = _FileListToArray(@ScriptDir & "\" & $Folder & "\", "*", $FLTA_FILESFOLDERS, True) ;switched from $FLTA_FILES for allowing main.au3 in folder
	If Not @error Then
		_Log($Folder & " Files (no filter): " & $FileArray[0])
		Local $FolderTreeItem = GUICtrlCreateTreeViewItem($Folder, $TreeID)

		For $i = 1 To $FileArray[0]
			Local $FileName = StringTrimLeft($FileArray[$i], StringInStr($FileArray[$i], "\", 0, -1))

			If StringInStr($FileArray[$i], "\.") Then ContinueLoop ;Use . for hidden
			If StringInStr(FileGetAttrib($FileArray[$i]), "D") And Not FileExists($FileArray[$i] & "\main.au3") Then ContinueLoop ;allow folders only if they contain main.au3

			_Log("Adding: " & $FileArray[$i])

			; Create item
			GUICtrlCreateTreeViewItem($FileName, $FolderTreeItem)

			; If item is in defaults file the check it
			If StringInStr($sDefaults, $FileName) Then
				_Log("Set state checked")
				GUICtrlSetState(-1, $GUI_CHECKED)
			EndIf

		Next

		GUICtrlSetState($FolderTreeItem, $GUI_EXPAND)

	Else
		_Log($Folder & " No files or missing")

	EndIf

	; If this the current folder isnt a custom folder, then try to process a custom folder
	If StringRight($Folder, 6) <> "Custom" Then _PopulateScripts($TreeID, $Folder & "Custom")

	Return $FileArray

EndFunc   ;==>_PopulateScripts

Func _NotAdminMsg($hwnd = "")
	_Log("_NotAdminMsg")
	MsgBox($MB_OK, $Title, "Not running with admin rights.", 0, $hwnd)

EndFunc   ;==>_NotAdminMsg

Func _RunTreeView($hWindow, $hTreeView, $ListOnly = False)
	_Log("_RunTreeView")

	Local $aList[0]

	For $iTop = 0 To ControlTreeView($hWindow, "", $hTreeView, "GetItemCount", "") - 1
		$Folder = ControlTreeView($hWindow, "", $hTreeView, "GetText", "#" & $iTop)

		For $iSub = 0 To ControlTreeView($hWindow, "", $hTreeView, "GetItemCount", "#" & $iTop) - 1
			$File = ControlTreeView($hWindow, "", $hTreeView, "GetText", "#" & $iTop & "|#" & $iSub)
			$FileChecked = ControlTreeView($hWindow, "", $hTreeView, "IsChecked", "#" & $iTop & "|#" & $iSub)

			If $FileChecked Then
				$RunFullPath = @ScriptDir & "\" & $Folder & "\" & $File
				_Log("Checked: $RunFullPath=" & $RunFullPath)
				If $ListOnly = False Then
					ControlTreeView($hWindow, "", $hTreeView, "Uncheck", "#" & $iTop & "|#" & $iSub)
					_RunFile($RunFullPath)
				EndIf
				_ArrayAdd($aList, $RunFullPath)
			EndIf
		Next

	Next

	Return $aList

EndFunc   ;==>_RunTreeView

Func _RunMulti($Folder)
	_Log("_RunMulti " & $Folder)

	$aFiles1 = _RunFolder(@ScriptDir & "\" & $Folder)
	$aFiles2 = _RunFolder(@ScriptDir & "\" & $Folder & "Custom")

	$iCount = _ArrayConcatenate($aFiles1, $aFiles2, 1)

	;$aFiles1[0] = $iCount

	Return $aFiles1
EndFunc   ;==>_RunMulti

Func _RunFolder($Path)
	_Log("_RunFolder " & $Path)
	$FileArray = _FileListToArray($Path, "*", $FLTA_FILESFOLDERS, True) ;switched from $FLTA_FILES for allowing main.au3 in folder
	If Not @error Then
		_Log("Files: " & $FileArray[0])
		For $i = 1 To $FileArray[0]
			If StringInStr($FileArray[$i], "\.") Then ContinueLoop
			If StringInStr(FileGetAttrib($FileArray[$i]), "D") And Not FileExists($FileArray[$i] & "\main.au3") Then ContinueLoop
			_Log($FileArray[$i])
			_RunFile($FileArray[$i])
		Next
		Return $FileArray[0]
	Else
		_Log("No files")
	EndIf

EndFunc   ;==>_RunFolder

Func _RunFile($File, $Params = "", $WorkingDir = "")
	_Log("_RunFile " & $File & " " & $Params)

	If StringInStr(FileGetAttrib($File), "D") And FileExists($File & "\main.au3") Then
		$File = $File & "\main.au3"
	EndIf

	$Extension = StringTrimLeft($File, StringInStr($File, ".", 0, -1))
	Switch $Extension
		Case "au3", "a3x"
			_Log("  au3")
			$RunLine = @AutoItExe & " /AutoIt3ExecuteScript """ & $File & """ " & $Params
			_Log("$RunLine=" & $RunLine)
			Return Run($RunLine, $WorkingDir, @SW_SHOW, $STDIO_INHERIT_PARENT)

		Case "ps1"
			_Log("  ps1")
			$RunLine = @ComSpec & " /c " & "powershell.exe -ExecutionPolicy Unrestricted -File """ & $File & """ " & $Params
			_Log("$RunLine=" & $RunLine)
			Return Run($RunLine, $WorkingDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)

		Case "reg"
			_Log("  reg")
			$RunLine = @ComSpec & " /c " & "reg import """ & $File & """"

			Local $Data = FileRead($File)
			If StringInStr($Data, ";32") Then
				$RunLine = $RunLine & " /reg:32"
			ElseIf StringInStr($Data, ";64") Then
				$RunLine = $RunLine & " /reg:64"
			ElseIf @CPUArch = "X64" Then
				$RunLine = $RunLine & " /reg:64"
			EndIf

			_Log("$RunLine=" & $RunLine)
			Return Run($RunLine, $WorkingDir, @SW_SHOW, $STDERR_CHILD + $STDOUT_CHILD)

		Case Else
			_Log("  other")
			Return ShellExecute($File, $Params, $WorkingDir)

	EndSwitch

EndFunc   ;==>_RunFile

Func _InternetPing()
	$InternetPing = Ping( "8.8.8.8", 400)

EndFunc

Func _WMIC($Command, $Search="", $Trim="=")
	Local $CMDOut = @TempDir & "\wmic_out.txt"
	RunWait(@ComSpec & ' /c ' & 'wmic ' & $Command & ' /value > ' & $CMDOut, "", @SW_HIDE)

	Local $hFile = FileOpen($CMDOut, 0)
	Local $Return

	If $Search<>"" Then
		While 1
			$Return = FileReadLine($hFile)
			If @error Then Return
			If StringLeft($Return,StringLen($Search)) = $Search Then
				ExitLoop
			Endif

		WEnd
	Else
		$Return = FileRead($hFile)
		$Return = StringStripWS($Return, 1+2)
	Endif

	If $Trim Then
		$Return = StringTrimLeft($Return, StringInStr($Return, $Trim) + StringLen($Trim) - 1)
		$Return = StringStripWS($Return, 1+2)
	Endif

	FileClose($hFile)

	Return($Return)
EndFunc

Func _Log($Msg, $Statusbar = False)
	Local $sTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "> "

	If Not IsDeclared("LogEdit") Then
		Global $LogWindow = GUICreate($Title & " Log", 750, 450, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX))
		Global $LogEdit = GUICtrlCreateEdit("", 0, 0, 750, 450, BitOR($ES_MULTILINE, $ES_WANTRETURN, $WS_VSCROLL, $WS_HSCROLL))
		GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
		GUICtrlSetColor(-1, 0xFFFFFF)
		GUICtrlSetBkColor(-1, 0x000000)
		GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
		GUISetState(@SW_SHOW)
		_GUICtrlEdit_AppendText($LogEdit, $Msg)
	Else

		_GUICtrlEdit_BeginUpdate($LogEdit)
		_GUICtrlEdit_AppendText($LogEdit, @CRLF & $Msg)
		_GUICtrlEdit_LineScroll($LogEdit, -StringLen($Msg), _GUICtrlEdit_GetLineCount($LogEdit))
		_GUICtrlEdit_EndUpdate($LogEdit)

	EndIf

	If $Statusbar Then _GUICtrlStatusBar_SetText($Statusbar, $Msg)
	ConsoleWrite($sTime & $Msg & @CRLF)
	If IsDeclared("LogFullPath") Then
		FileWrite($LogFullPath, $sTime & $Msg & @CRLF)
	EndIf

	Return $Msg
EndFunc   ;==>_Log

Func _CommError()
	Local $HexNumber
	Local $strMsg

	$HexNumber = Hex($oCommError.Number, 8)
	$strMsg = "Error: " & $HexNumber
	$strMsg &= "  Desc: " & $oCommError.WinDescription
	$strMsg &= "  Line: " & $oCommError.ScriptLine

	_Log($strMsg)

EndFunc   ;==>_CommError

Func _Exit()
	_Log("End script " & $CmdLineRaw)

EndFunc   ;==>_Exit
