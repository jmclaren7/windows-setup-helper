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
#include <include\GuiToolTip.au3>
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

; Register a function to run whenever the script exits
OnAutoItExitRegister("_Exit")

; Fix for issues working in a 64-bit only environment
_WinAPI_Wow64EnableWow64FsRedirection(False)

Opt("WinTitleMatchMode", -2)
Opt("TrayIconHide", 1)

FileChangeDir(@ScriptDir)

Global $LogFullPath = StringReplace(@TempDir & "\Helper_" & @ScriptName, ".au3", ".log")
Global $Date = StringTrimRight(FileGetTime(@ScriptFullPath, $FT_MODIFIED, $FT_STRING), 2)
Global $Version = "5.2.0." & $Date
Global $Title = "Windows Setup Helper v" & $Version
Global $GUIMain
Global $oCommError = ObjEvent("AutoIt.Error", "_CommError")
Global $StatusBar1
Global $StatusbarTimer2
Global $IsPE = StringInStr(@WindowsDir, "X:")
Global $DoubleClick = False
Global $Debug = Not $IsPE

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

; Check for command line parameters
If $CmdLine[0] >= 1 Then
	$Command = $CmdLine[1]
Else
	$Command = "winpe"
EndIf

_Log("Command: " & $Command)

Switch $Command
	Case "winpe"

		; Run automatic setup scripts
		If $IsPE Then _RunMulti("PEAutoRun")

		#Region ### START Koda GUI section ###
		$GUIMain = GUICreate("$Title", 767, 543, -1, -1)
		$FileMenu = GUICtrlCreateMenu("&File")
		$AdvancedMenu = GUICtrlCreateMenu("&Advanced")
		$Tab1 = GUICtrlCreateTab(7, 4, 753, 495)
		$BootTabSheet = GUICtrlCreateTabItem("&")
		$Group5 = GUICtrlCreateGroup("WinPE Tools", 19, 37, 360, 452)
		$PERunButton = GUICtrlCreateButton("Run", 257, 454, 107, 25)
		$PEScriptTreeView = GUICtrlCreateTreeView(35, 61, 330, 385)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group6 = GUICtrlCreateGroup("First Logon Scripts", 387, 37, 360, 452)
		$NormalInstallButton = GUICtrlCreateButton("Normal Install", 403, 454, 131, 25)
		$AutomatedInstallButton = GUICtrlCreateButton("Automated Install", 563, 454, 163, 25, $BS_DEFPUSHBUTTON)
		$PEInstallTreeView = GUICtrlCreateTreeView(404, 61, 330, 345, BitOR($GUI_SS_DEFAULT_TREEVIEW, $TVS_CHECKBOXES))
		$PEComputerNameInput = GUICtrlCreateInput("", 556, 422, 169, 21)
		$Label4 = GUICtrlCreateLabel("Computer Name", 476, 425, 80, 17)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		GUICtrlCreateTabItem("")
		$StatusBar1 = _GUICtrlStatusBar_Create($GUIMain)
		_GUICtrlStatusBar_SetSimple($StatusBar1)
		_GUICtrlStatusBar_SetText($StatusBar1, "")
		GUISetState(@SW_SHOW)
		#EndRegion ### END Koda GUI section ###

		; File Menu Items
		$MenuExitButton = GUICtrlCreateMenuItem("Exit", $FileMenu)

		; Advanced Menu Items
		$MenuShowConsole = GUICtrlCreateMenuItem("Show Console", $AdvancedMenu)
		$MenuOpenLog = GUICtrlCreateMenuItem("Open Log File", $AdvancedMenu)
		$MenuRunMain = GUICtrlCreateMenuItem("Run Main", $AdvancedMenu)
		$MenuRelistScripts = GUICtrlCreateMenuItem("Relist Tools && Scripts", $AdvancedMenu)
		$MenuListDebugTools = GUICtrlCreateMenuItem("List Debug && AutoRun Tools", $AdvancedMenu)
		$MenuInstallDisk0 = -1;GUICtrlCreateMenuItem("Auto Install To Disk 0", $AdvancedMenu)

		; GUI Post Creation Setup
		WinSetTitle($GUIMain, "", $Title)
		GUICtrlSendMsg($PEComputerNameInput, $EM_SETCUEBANNER, False, "(Optional)")
		GUICtrlSetLimit($PEComputerNameInput, 15)

		; Generate Script List
		_PopulateScripts($PEInstallTreeView, "Logon*")
		_PopulateScripts($PEScriptTreeView, "Tools*")

		Local $hSetup
		Local $RebootPrompt = False
		Local $CopyAutoLogonFiles = False
		Local $Reboot = False
		Local $BootDrive = StringLeft(@SystemDir, 3)

		; Start PE networking
		If $IsPE Then Run(@ComSpec & " /c " & 'wpeinit.exe', @SystemDir, @SW_HIDE, $RUN_CREATE_NEW_CONSOLE)

		; Set GUI Icon
		GUISetIcon($BootDrive & "sources\setup.exe")

		; Hide console windows
		_Log("Hide console window")
		WinSetState($Title & " Log", "", @SW_HIDE)

		; Setup statusbar updates
		Global $StatusBarToolTip = _GUIToolTip_Create(0);BitOr($TTS_ALWAYSTIP, $TTS_NOPREFIX, $TTS_BALLOON)
		_GUIToolTip_AddTool($StatusBarToolTip, 0, " ", $StatusBar1)
		_GUIToolTip_SetMaxTipWidth($StatusBarToolTip, 400)
		AdlibRegister("_StatusBarUpdate", 5000)
		_StatusBarUpdate()

		; Setup double click detection for $PEScriptTreeView
		$HUser32DLL = DllOpen(@WindowsDir & "\System32\user32.dll")
		Global $hPEScriptTreeView = GUICtrlGetHandle($PEScriptTreeView)
		GUIRegisterMsg($WM_NOTIFY, "_WM_NOTIFY")

		_Log("Ready", True)

		;GUI Loop
		While 1
			$nMsgA = GUIGetMsg(1)
			$nMsg = $nMsgA[0]
			If $Debug And $nMsg > 0 Then _Log("MSG $nMsg=" & $nMsg)
			Switch $nMsg
				Case $GUI_EVENT_CLOSE, $MenuExitButton
					If $nMsgA[1] <> $GUIMain Then ContinueLoop
					If $IsPE And MsgBox(1, $Title, "Closing the program will reboot the system while in WinPE.") <> 1 Then ContinueLoop
					Exit

				Case $MenuShowConsole
					WinSetState($Title & " Log", "", @SW_SHOW)

				Case $MenuRelistScripts
					_GUICtrlTreeView_DeleteAll($PEInstallTreeView)
					_GUICtrlTreeView_DeleteAll($PEScriptTreeView)
					_PopulateScripts($PEInstallTreeView, "Logon*")
					_PopulateScripts($PEScriptTreeView, "Tools*")

				Case $MenuListDebugTools
					_PopulateScripts($PEScriptTreeView, "Debug")
					_PopulateScripts($PEScriptTreeView, "PEAutoRun*")

				Case $MenuOpenLog
					_Log("MenuOpenLog")
					;_RunFile($LogFullPath) ; Not working in Win11 PE
					ShellExecute("notepad.exe", $LogFullPath)

				Case $MenuRunMain
					_Log("MenuRunMain")
					_RunFile("Main.au3")

				Case $PERunButton
					$Item = GUICtrlRead(GUICtrlRead($PEScriptTreeView), 1)
					$Parent = GUICtrlRead(_GUICtrlTreeView_GetParentParam($PEScriptTreeView, GUICtrlRead($PEScriptTreeView)), 1)

					If IsString($Parent) Then
						_RunFile(_GetTreeItemFullPath($Parent, $Item))
					Else
						_Log("Invalid treeitem")
					EndIf

				Case $NormalInstallButton
					$hSetup = _RunFile($BootDrive & "sources\setup.exe", "/noreboot")
					$CopyAutoLogonFiles = False
					$RebootPrompt = True

				Case $AutomatedInstallButton, $MenuInstallDisk0
					If IsDeclared($hSetup) And ProcessExists($hSetup) Then
						MsgBox(0, "Error - " & $Title, "Setup is already running, please close it first")
						ContinueLoop
					EndIf

					If $nMsg = $MenuInstallDisk0 And MsgBox($MB_OKCANCEL, "Danger - " & $Title, "Setup will automaticly use disk 0 to install Windows") <> $IDOK Then ContinueLoop

					$aAutoLogonCopy = _RunTreeView($GUIMain, $PEInstallTreeView, True)
					For $b = 0 To UBound($aAutoLogonCopy) - 1
						_Log("TreeItem: " & $aAutoLogonCopy[$b])
					Next

					; Read the autounattend.xml file to memory
					$sFileData = FileRead(@ScriptDir & "\autounattend.xml")

					; Make modifications to autounattend.xml
					$ComputerName = GUICtrlRead($PEComputerNameInput)
					If $ComputerName <> "" Then
						_Log("$ComputerName=" & $ComputerName)
						$sFileData = StringReplace($sFileData, "<ComputerName>*</ComputerName>", "<ComputerName>" & $ComputerName & "</ComputerName>")
						_Log("StringReplace @extended=" & @extended)
					EndIf

					If $nMsg = $MenuInstallDisk0 Then
						;$sFileData = StringReplace($sFileData, "", "")
					EndIf

					If @OSVersion = "WIN_10" Then
						_Log("WIN_10")
						$sFileData = StringReplace($sFileData, "Windows 11", "Windows 10")
						_Log("StringReplace @extended=" & @extended)
					EndIf

					; Save modifications to autounattend.xml in new location
					$AutounattendPath = @TempDir & "\autounattend.xml"
					_Log("$AutounattendPath=" & $AutounattendPath)
					$hAutounattend = FileOpen($AutounattendPath, $FO_OVERWRITE)
					FileWrite($hAutounattend, $sFileData)
					_Log("FileWrite @error=" & @error)
					FileClose($hAutounattend)

					$hSetup = _RunFile($BootDrive & "sources\setup.exe", "/noreboot /unattend:" & $AutounattendPath)

					$CopyAutoLogonFiles = True

			EndSwitch

			; If a double click is detected on the PE Tools treeview run the script
			If $DoubleClick Then
				$DoubleClick = False
				$Item = GUICtrlRead(GUICtrlRead($PEScriptTreeView), 1)
				$Parent = GUICtrlRead(_GUICtrlTreeView_GetParentParam($PEScriptTreeView, GUICtrlRead($PEScriptTreeView)), 1)

				If IsString($Parent) Then
					_RunFile(_GetTreeItemFullPath($Parent, $Item))
				Else
					_Log("Invalid treeitem")
				EndIf

			EndIf

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
						; Doing this also requieres a registry value set in the install image
						$Return = FileCopy($AutounattendPath, $Target, 1 + 8)
						_Log("FileCopy: " & $AutounattendPath & " (" & $Return & ")")

						; Copy the script that is run at Logon and runs the other scripts we copy
						If UBound($aAutoLogonCopy) Then
							$AutoLogonSource = @ScriptDir & "\Logon\.Autorun.ps1"
							$Return = FileCopy($AutoLogonSource, $Target, 1 + 8)
							_Log("FileCopy: " & $AutoLogonSource & " (" & $Return & ")")
						EndIf

						; Copy selected files from the install list
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

			Sleep(10)
		WEnd

	Case Else
		_Log("Command unknown")

EndSwitch

;=========== =========== =========== =========== =========== =========== =========== ===========
;=========== =========== =========== =========== =========== =========== =========== ===========

; Proccess NOTIFY mesages to handle double clicks
Func _WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	$TagStruct1 = DllStructCreate($tagNMHDR, $lParam)
	$hWndFrom = HWnd(DllStructGetData($TagStruct1, "hWndFrom"))
	$IDFrom = DllStructGetData($TagStruct1, "IDFrom")
	$Code = DllStructGetData($TagStruct1, "Code")

	If $hWndFrom = $hPEScriptTreeView And $Code = -3 Then
		$DoubleClick = True
	EndIf

	$TagStruct1 = 0
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_NOTIFY

Func _GetTreeItemFullPath($Parent, $Item)
	Local $FullPath

	_Log("$Parent=" & $Parent & " $Item=" & $Item)

	; Test if path given is a full path
	If StringInStr($Parent, ":", 0, 1, 2, 1) Then
		$FullPath = $Parent & "\" & $Item
	Else
		$FullPath = @ScriptDir & "\" & $Parent & "\" & $Item
	EndIf

	;If Not FileExists($FullPath) Then Return SetError(1, 0, $FullPath)

	Return $FullPath

EndFunc   ;==>_GetTreeItemFullPath

Func _GetSimilarPaths($Folder, $Path = Default)
	If $Path = Default Then $Path = @ScriptDir

	_Log("_GetSimilarPaths: " & $Folder & " - " & $Path)

	; Get folders from the script path
	Local $aFolders = _FileListToArrayRec($Path, $Folder & "*", $FLTAR_FOLDERS, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)
	If @error Then _Log("@error=" & @error & " @extended=" & @extended)

	; Check other drives for similar folders and add them to the list
	Local $aDrivesLetters = DriveGetDrive($DT_ALL)
	For $i = 1 To $aDrivesLetters[0]
		$aDrivesLetters[$i] = StringUpper($aDrivesLetters[$i])
		_Log("  Drive: " & $aDrivesLetters[$i])
		$aOtherDrives = _FileListToArrayRec($aDrivesLetters[$i] & "\Helper", $Folder & "*", $FLTAR_FOLDERS, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)

		_ArrayConcatenate($aFolders, $aOtherDrives, 1)
	Next

	; Remove duplicated and update index 0 to reflect total number of values
	$aFolders = _ArrayUnique($aFolders, 0, 1)

	Return $aFolders

EndFunc   ;==>_GetSimilarPaths

Func _PopulateScripts($TreeID, $Folder)
	_Log("_PopulateScripts " & $Folder)

	Local $FolderFullPath, $CheckAll = False, $CollapseTree = False, $ResolveSimilarPaths = False

	If StringRight($Folder, 1) = "*" Then
		$Folder = StringTrimRight($Folder, 1)
		$ResolveSimilarPaths = True
	EndIf

	; Test if path given is a full path
	If StringInStr($Folder, ":", 0, 1, 2, 1) And StringInStr(FileGetAttrib($Folder), "D") Then
		$FolderFullPath = $Folder
		If StringRight($FolderFullPath, 1) = "\" Then $FolderFullPath = StringTrimRight($FolderFullPath, 1)
	Else
		$FolderFullPath = @ScriptDir & "\" & $Folder
	EndIf

	_Log("  $FolderFullPath=" & $FolderFullPath)

	Local $aFiles = _FileListToArray($FolderFullPath & "\", "*", $FLTA_FILESFOLDERS, True) ;switched from $FLTA_FILES for allowing main.au3 in folder
	If Not @error Then
		_Log("  " & $Folder & " Files (no filter): " & $aFiles[0])

		; Load options file
		Local $sOptions = FileRead($FolderFullPath & "\.Options.txt")
		If Not @error Then _Log("  Options list: " & $sOptions)

		; Create parent tree item
		Local $FolderTreeItem = GUICtrlCreateTreeViewItem($Folder, $TreeID)

		If StringInStr($sOptions, "CheckAll") Then _GUICtrlTreeView_SetChecked($TreeID, $FolderTreeItem)

		For $i = 1 To $aFiles[0]
			Local $FileName = StringTrimLeft($aFiles[$i], StringInStr($aFiles[$i], "\", 0, -1))

			If StringInStr($aFiles[$i], "\.") Then ContinueLoop ;use . for hidden
			If StringInStr(FileGetAttrib($aFiles[$i]), "D") And Not FileExists($aFiles[$i] & "\main.au3") Then ContinueLoop ;allow folders only if they contain main.au3

			_Log("  Adding: " & $aFiles[$i])

			; Create sub item
			$ThisItem = GUICtrlCreateTreeViewItem($FileName, $FolderTreeItem)

			; If item is in defaults file or $CheckAll is set then check it
			If StringInStr($sOptions, "CheckAll") Or StringInStr($sOptions, $FileName) Then
				_Log("  Set state checked")
				GUICtrlSetState($ThisItem, $GUI_CHECKED)

			EndIf

		Next

		; CollapseTree option
		If Not StringInStr($sOptions, "CollapseTree") Then _GUICtrlTreeView_Expand($TreeID, $FolderTreeItem)


	Else
		_Log("  " & $Folder & " No files or missing")

	EndIf


	If $ResolveSimilarPaths = True Then
		_Log("  ResolveSimilarPaths")
		;$FolderParentPath = StringLeft($FolderFullPath, StringInStr($FolderFullPath, "\", 0, -1) - 1)

		$aOtherFolders = _GetSimilarPaths($Folder)

		; Run _PopulateScripts on the similar folders
		For $i = 1 To $aOtherFolders[0]
			_Log("  " & $aOtherFolders[$i])

			; The list will include the folder we just proccessed so skip it
			If $aOtherFolders[$i] = $FolderFullPath Then ContinueLoop

			; If the folder is in the script path then treat it as relative (this is handled later when running a tool)
			$aOtherFolders[$i] = StringReplace($aOtherFolders[$i], @ScriptDir & "\", "")

			_Log("  Recurse _PopulateScripts for: " & $aOtherFolders[$i])
			_PopulateScripts($TreeID, $aOtherFolders[$i])
		Next


	EndIf

	_Log("End _PopulateScripts for: " & $Folder)

	Return $aFiles

EndFunc   ;==>_PopulateScripts

Func _NotAdminMsg($hWnd = "")
	_Log("_NotAdminMsg")
	MsgBox($MB_OK, $Title, "Not running with admin rights.", 0, $hWnd)

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

; Run all scripts in a folder and folders starting with the same name and on other drives
Func _RunMulti($Folder)
	_Log("_RunMulti " & $Folder)

	Local $Paths = _GetSimilarPaths($Folder)
	For $x = 1 To $Paths[0]
		_RunFolder($Paths[$x])

	Next

	Return $Paths
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

Func _StatusBarUpdate()
	Local $StatusbarText, $StatusbarToolTipText = "Additonal Details:"
	Local $Delimiter = "  |  "

	If $Debug Then
		Local $StatusBarTimer1 = TimerInit()
		Local $StatusBarTimer2Value = TimerDiff($StatusbarTimer2)
	EndIf

	; Get internet status/latency
	$InternetPing = Ping("8.8.8.8", 400)
	If @error Then $InternetPing = Ping("1.1.1.1", 200)

	$InternetPing = Round($InternetPing / 5) * 5 ; Reduces updates to GUI

	If $InternetPing Then
		$StatusbarText &= "Online (" & $InternetPing & "ms)"
	Else
		$StatusbarText &= "Offline"
	EndIf

	; Get gateway
	$StatusbarToolTipText &= @CR & "Gateway: " & _WMI("SELECT NextHop From Win32_IP4RouteTable WHERE Destination = '0.0.0.0'").NextHop

	; Get IP addresses
	$StatusbarText &= $Delimiter & _WMI("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True and DHCPEnabled = True").IPAddress[0]
	$StatusbarToolTipText &= @CR & "Other IPs: " & @IPAddress1 & ", " & @IPAddress2 & ", " & @IPAddress3 & ", " & @IPAddress4

	; Get memory information
	If Not IsDeclared("_MemStats") Then Global $_MemStats = MemGetStats()
	$StatusbarText &= $Delimiter & Round($_MemStats[1] / 1024 / 1024) & "GB"

	; Get CPU information
	$objItem = _WMI("SELECT NumberOfCores,NumberOfLogicalProcessors FROM Win32_Processor")
	$StatusbarText &= $Delimiter & $objItem.NumberOfCores & "/" & $objItem.NumberOfLogicalProcessors & " Cores"

	; Get motherboard bios information
	$objItem = _WMI("SELECT * FROM Win32_BIOS")
	If Not @error and $objItem.SerialNumber <> "" and $objItem.SerialNumber <> "System Serial Number" Then $StatusbarText &= $Delimiter & $objItem.SerialNumber
	$StatusbarText &= $Delimiter & "FW: " & $objItem.SMBIOSBIOSVersion & " (" & StringLeft($objItem.ReleaseDate, 8) & ")"

	; Get additional statusbar and tool tip text
	$HelperStatusFiles = _FileListToArray(@TempDir, "Helper_Status_*.txt", $FLTA_FILES, True)
	For $i = 1 To Ubound($HelperStatusFiles) - 1
		If _FileModifiedAge($HelperStatusFiles[$i]) < 10 * 1000 Then
			$FileText = FileReadLine($HelperStatusFiles[$i], 1)
			_Log("$FileText=" & $FileText)
			If Not @error Then $StatusbarText &= $Delimiter & $FileText

			$FileText = FileReadLine($HelperStatusFiles[$i], 2)
			_Log("$FileText=" & $FileText)
			If Not @error Then $StatusBarToolTipText &= @CRLF & $FileText

		Else
			FileDelete($HelperStatusFiles[$i])
		EndIf

	Next

	; Update statusbar if the text changed
	If _GUICtrlStatusBar_GetText($StatusBar1, 0) <> $StatusbarText Then
		_GUICtrlStatusBar_SetText($StatusBar1, $StatusbarText)
		If $Debug Then _Log("Statusbar Updated")
	EndIf

	; Update statusbar tool tip if the text changed
	If _GUIToolTip_GetText($StatusBarToolTip, 0, $StatusBar1) <> $StatusBarToolTipText Then
		_GUIToolTip_UpdateTipText($StatusBarToolTip, 0, $StatusBar1, $StatusBarToolTipText)
		If $Debug Then _Log("Statusbar Tooltip Updated")
	EndIf

	If $Debug Then
		_Log("  _StatusBarUpdate Timers: " & Round(TimerDiff($StatusBarTimer1)) & "ms " & Round($StatusBarTimer2Value) & "ms")
		$StatusBarTimer2 = TimerInit()
	EndIf

	Return
EndFunc   ;==>_StatusBarUpdate

Func _Log($Msg, $Statusbar = False)
	Local $sTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "> "

	If Not IsDeclared("_LogEdit") Then
		Global $_LogWindow = GUICreate($Title & " Log", 750, 450, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX))
		Global $_LogEdit = GUICtrlCreateEdit("", 0, 0, 750, 450, BitOR($ES_MULTILINE, $ES_WANTRETURN, $WS_VSCROLL, $WS_HSCROLL))
		GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
		GUICtrlSetColor(-1, 0xFFFFFF)
		GUICtrlSetBkColor(-1, 0x000000)
		GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
		GUISetState(@SW_SHOW)
		_GUICtrlEdit_AppendText($_LogEdit, $Msg)
	Else

		_GUICtrlEdit_BeginUpdate($_LogEdit)
		_GUICtrlEdit_AppendText($_LogEdit, @CRLF & $Msg)
		_GUICtrlEdit_LineScroll($_LogEdit, -StringLen($Msg), _GUICtrlEdit_GetLineCount($_LogEdit))
		_GUICtrlEdit_EndUpdate($_LogEdit)

	EndIf

	If $Statusbar Then _GUICtrlStatusBar_SetText($Statusbar, $Msg)

	ConsoleWrite($sTime & $Msg & @CRLF)

	If IsDeclared("LogFullPath") Then
		If Not IsDeclared("_hLogFile") Then Global $_hLogFile = FileOpen($LogFullPath, $FO_APPEND)
		FileWrite($_hLogFile, $sTime & $Msg & @CRLF)
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
