#include "include\ButtonConstants.au3"
#include "include\ComboConstants.au3"
#include "include\Crypt.au3"
#include "include\EditConstants.au3"
#include "include\File.au3"
#include "include\FileConstants.au3"
#include "include\GuiConstantsEx.au3"
#include "include\GuiListView.au3"
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
#include "include\WinAPIFiles.au3"
#include "includeExt\Json.au3"
#include "includeExt\WinHttp.au3"
#include "includeExt\ActivationStatus.au3"
#include "includeExt\Custom.au3"
#include "includeExt\_Zip.au3"

OnAutoItExitRegister("_Exit")
_WinAPI_Wow64EnableWow64FsRedirection(False)

If StringInStr(@ScriptFullPath, "$OEM$\$$\IT") Then
	Global $LogFullPath = StringReplace(@TempDir & "\" & @ScriptName, ".au3", ".log")
Else
	Global $LogFullPath = StringReplace(@ScriptFullPath, ".au3", ".log")
EndIf

Global $MainSize = FileGetSize(@ScriptFullPath)
Global $Version = "3.0.1." & $MainSize

Global $Title = "IT Setup Helper v" & $Version
Global $DownloadUpdatedCount = 0
Global $DownloadErrors = 0
Global $DownloadUpdated = ""
Global $GITURL = "https://github.com/jmclaren7/itdeployhelper"
Global $GITAPIURL = "https://api.github.com/repos/jmclaren7/itdeployhelper/contents"
Global $GITZIP = "https://github.com/jmclaren7/itdeployhelper/archive/master.zip"
Global $GUIMain
Global $oCommError = ObjEvent("AutoIt.Error", "_CommError")
Global $StatusBar1
Global $UserCreatedWithAdmin = False

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

;If FileExists(@ScriptDir & "\noexecute") Then Exit

Global $TokenAddHeader = IniRead(".token", "t", "t", "")
;If $TokenAddHeader = "" Then $TokenAddHeader = IniRead("git.token", "t", "t", "")
If $TokenAddHeader <> "" Then
	_Log("Token Added")
	$TokenAddHeader = "Authorization: token " & $TokenAddHeader
EndIf

If $CmdLine[0] >= 1 Then
	$Command = $CmdLine[1]
Else
	$Command = "main-gui"
EndIf

_Log("Command: " & $Command)

Switch $Command
	Case "bootmedia", "boot-gui"
		$BootDrive = StringLeft(@SystemDir, 3)

		; Start network
		Run(@ComSpec & " /c " & 'wpeinit.exe', @SystemDir, @SW_SHOW, $RUN_CREATE_NEW_CONSOLE)

		; Boot GUI
		$GUIBoot = GUICreate("$Title", 625, 442, -1, -1)
		GUISetBkColor(0xFFFFFF)
		$NormalInstallButton = GUICtrlCreateButton("Normal Install", 23, 392, 107, 25)
		$AutomatedInstallButton = GUICtrlCreateButton("Continue Automated Install", 439, 392, 155, 25, $BS_DEFPUSHBUTTON)
		$Label1 = GUICtrlCreateLabel("Automated Install Options", 24, 24, 213, 24)
		GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
		GUICtrlSetColor(-1, 0x3399FF)
		$Boot_ScriptsTree = GUICtrlCreateTreeView(360, 56, 233, 305, BitOR($GUI_SS_DEFAULT_TREEVIEW, $TVS_CHECKBOXES))
		$AdvancedButton = GUICtrlCreateButton("Advanced", 142, 392, 107, 25)
		$Label2 = GUICtrlCreateLabel("Select scripts to automaticly run once installation is complete", 48, 72, 288, 17)
		$Label3 = GUICtrlCreateLabel("Select normal Install to skip all automated install features", 48, 96, 268, 17)

		Opt("WinTitleMatchMode", 2)
		WinSetState("bootmedia.exe", "", @SW_MINIMIZE)

		;$TestButton = GUICtrlCreateButton("Test", 262, 392, 107, 25)

		GUISetState(@SW_SHOW)
		WinSetTitle($GUIBoot, "", $Title)
		GUISetIcon($BootDrive & "sources\setup.exe")

		; Generate script checkboxes
		_PopulateScripts($Boot_ScriptsTree, "OptLogin")
		_PopulateScripts($Boot_ScriptsTree, "OptCustom")

		; Loop
		While 1
			$nMsg = GUIGetMsg()

			Switch $nMsg
				Case $TestButton


				Case $GUI_EVENT_CLOSE
					Exit

				Case $AdvancedButton
					_RunFile("cmd.exe")


				Case $NormalInstallButton
					_RunFile($BootDrive & "sources\setup.exe")
					While 1
						For $i = 65 To 90
							$Path = Chr($i) & ":\Windows\IT"
							If FileExists($Path) Then
								$Return = DirRemove($Path, 1)
								_Log("Found: " & $Path & " DirRemove: " & $Return)
								Sleep(1000)
							EndIf
						Next

						Sleep(10)
					WEnd

				Case $AutomatedInstallButton
					$aList = _RunTreeView($GUIBoot, $Boot_ScriptsTree, True)
					For $b = 0 To UBound($aList) - 1
						_Log("TreeItem: " & $aList[$b])
					Next

					_RunFile($BootDrive & "sources\setup.exe", "/unattend:" & @ScriptDir & "\autounattend.xml")

					While 1
						For $i = 65 To 90
							$Path = Chr($i) & ":\Windows\IT"
							If FileExists($Path) Then
								_Log("Found: " & $Path)
								For $iFile = 0 To UBound($aList) - 1
									$Dest = $Path & "\AutoLogin\"
									$Return = FileCopy($aList[$iFile], $Dest, 1)
									If $Return Then _Log("FileCopy: " & $Dest)

								Next
								Sleep(4000)
							EndIf
						Next

						Sleep(10)
					WEnd

			EndSwitch

			Sleep(10)
		WEnd

	Case "system"
		_RunFolder(@ScriptDir & "\AutoSystem\")

	Case "login"
		ProcessWait("Explorer.exe", 60)
		Sleep(5000)

		If Not StringInStr($CmdLineRaw, "skipupdate") Then
			_GitUpdate()
			If StringInStr($DownloadUpdated, @ScriptName) Then
				_RunFile(@ScriptFullPath, "login skipupdate")
				Exit
			EndIf
		EndIf

		FileCreateShortcut(@AutoItExe, @DesktopDir & "\IT Setup Helper.lnk", @ScriptDir, "/AutoIt3ExecuteScript """ & @ScriptFullPath & """")
		FileCreateShortcut(@ScriptDir, @DesktopDir & "\IT Setup Folder")

		WinMinimizeAll()

		_RunFolder(@ScriptDir & "\AutoLogin\")
		_RunFolder(@ScriptDir & "\AutoCustom\")
		_RunFile(@ScriptFullPath)

	Case "main-gui", ""
		#Region ### START Koda GUI section ###
		$GUIMain = GUICreate("$Title", 823, 574, -1, -1)
		$MenuItem2 = GUICtrlCreateMenu("&File")
		$MenuExitButton = GUICtrlCreateMenuItem("Exit", $MenuItem2)
		$MenuItem1 = GUICtrlCreateMenu("&Advanced")
		$MenuUpdateButton = GUICtrlCreateMenuItem("Update from GitHub", $MenuItem1)
		$MenuVisitGitButton = GUICtrlCreateMenuItem("Visit GitHub Page", $MenuItem1)
		$MenuShowAllScriptsButton = GUICtrlCreateMenuItem("Show All Scripts", $MenuItem1)
		$MenuOpenLog = GUICtrlCreateMenuItem("Open Log", $MenuItem1)
		$MenuOpenFolder = GUICtrlCreateMenuItem("Open Program Folder", $MenuItem1)
		$Tab1 = GUICtrlCreateTab(7, 4, 809, 521)
		$TabSheet1 = GUICtrlCreateTabItem("Main")
		$Group1 = GUICtrlCreateGroup("Scripts", 399, 33, 401, 481)
		$Presets = GUICtrlCreateCombo("Presets", 415, 57, 369, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
		GUICtrlSetState(-1, $GUI_DISABLE)
		$ScriptsTree = GUICtrlCreateTreeView(415, 97, 369, 369, BitOR($GUI_SS_DEFAULT_TREEVIEW, $TVS_CHECKBOXES))
		$RunButton = GUICtrlCreateButton("Run", 711, 481, 75, 25)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group2 = GUICtrlCreateGroup("Information", 22, 32, 361, 257)
		$InfoList = GUICtrlCreateListView("", 31, 50, 346, 230, BitOR($GUI_SS_DEFAULT_LISTVIEW, $LVS_SMALLICON), 0)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group3 = GUICtrlCreateGroup("Create Local User", 22, 426, 361, 89)
		$CreateLocalUserButton = GUICtrlCreateButton("Create Local User", 235, 478, 131, 25)
		$UsernameInput = GUICtrlCreateInput("", 38, 448, 185, 21)
		$PasswordInput = GUICtrlCreateInput("", 38, 480, 185, 21)
		$AdminCheckBox = GUICtrlCreateCheckbox("Local Administrator", 238, 450, 113, 17)
		GUICtrlSetState(-1, $GUI_CHECKED)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group4 = GUICtrlCreateGroup("Actions", 22, 294, 361, 129)
		$JoinButton = GUICtrlCreateButton("Domain && Computer Name", 35, 319, 160, 25)
		$DisableAdminButton = GUICtrlCreateButton("Disable Administrator", 35, 354, 160, 25)
		$SignOutButton = GUICtrlCreateButton("Sign Out", 35, 389, 160, 25)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		GUICtrlCreateTabItem("")
		$StatusBar1 = _GUICtrlStatusBar_Create($GUIMain)
		_GUICtrlStatusBar_SetSimple($StatusBar1)
		_GUICtrlStatusBar_SetText($StatusBar1, "")
		GUISetState(@SW_SHOW)
		#EndRegion ### END Koda GUI section ###

		;GUI Post Creation Setup
		WinSetTitle($GUIMain, "", $Title)
		GUICtrlSendMsg($UsernameInput, $EM_SETCUEBANNER, False, "Username")
		GUICtrlSendMsg($PasswordInput, $EM_SETCUEBANNER, False, "Password (optional)")

		;Info List Generation
		If IsAdmin() Then
			GUICtrlCreateListViewItem("Running with admin rights", $InfoList)
			GUICtrlSetColor(-1, "0x00a500")
		Else
			GUICtrlCreateListViewItem("Running without admin rights", $InfoList)
			GUICtrlSetColor(-1, "0xff1000")
		EndIf

		GUICtrlCreateListViewItem("Current User: " & @UserName, $InfoList)
		If @UserName = "Administrator" Then
			GUICtrlSetColor(-1, "0xffa500")
		EndIf

		GUICtrlCreateListViewItem("Computer Name: " & @ComputerName, $InfoList)
		GUICtrlCreateListViewItem("Login Domain: " & @LogonDomain, $InfoList)
		$Manufacturer = RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "SystemManufacturer")
		If $Manufacturer = "System manufacturer" Then $Manufacturer = "Unknown"
		GUICtrlCreateListViewItem("Manufacturer: " & $Manufacturer, $InfoList)
		GUICtrlCreateListViewItem("Model: " & RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "SystemProductName"), $InfoList)
		GUICtrlCreateListViewItem("BIOS Version: " & RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "BIOSVersion"), $InfoList)
		GUICtrlCreateListViewItem("BIOS Mode: " & _WinAPI_GetFirmwareEnvironmentVariable(), $InfoList)
		$WinAPISystemInfo = _WinAPI_GetSystemInfo()
		GUICtrlCreateListViewItem("CPU Cores/Logical Cores: " & $WinAPISystemInfo[5] & "/" & EnvGet("NUMBER_OF_PROCESSORS"), $InfoList)
		$MemStats = MemGetStats()
		GUICtrlCreateListViewItem("Installed Memory: " & Round($MemStats[$MEM_TOTALPHYSRAM] / 1024 / 1024, 1) & "GB", $InfoList)
		GUICtrlCreateListViewItem("License: " & IsActivated(), $InfoList)
		$NetInfo = _NetAdapterInfo()
		GUICtrlCreateListViewItem("IP/Gateway: " & $NetInfo[3] & "/" & $NetInfo[4], $InfoList)
		GUICtrlCreateListViewItem("MAC: " & $NetInfo[2], $InfoList)

		;Generate Script List
		_PopulateScripts($ScriptsTree, "OptLogin")
		_PopulateScripts($ScriptsTree, "OptCustom")

		_Log("Ready", True)

		;GUI Loop
		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $GUI_EVENT_CLOSE
					Exit

				Case $DisableAdminButton
					_Log("DisableAdminButton")

					If @ComputerName = @LogonDomain And Not $UserCreatedWithAdmin Then
						If MsgBox($MB_YESNO, $Title, "Are you sure?" & @CRLF & @CRLF & "This computer might not be joined to a domain and it looks like you haven't created a local user with admin rights.", 0, $GUIMain) <> $IDYES Then
							ContinueLoop
						EndIf
					EndIf

					If IsAdmin() Then
						_Log("Disable admin command")
						Run(@ComSpec & " /c " & 'net user administrator /active:no', @SystemDir, @SW_SHOW)
					Else
						_NotAdminMsg($GUIMain)
					EndIf

				Case $SignOutButton
					_Log("SignOutButton")
					Shutdown(0)

				Case $RunButton
					_Log("RunButton")
					_RunTreeView($GUIMain, $ScriptsTree)

				Case $MenuUpdateButton
					_Log("MenuUpdateButton")
					$aUpdates = _GitUpdate(True)
					If @error Then ContinueLoop
					$UpdatesCount = UBound($aUpdates)

					If $UpdatesCount = 0 Then
						MsgBox(0, $Title, "No updates")

					ElseIf MsgBox($MB_YESNO, $Title, "Restart script?", 0, $GUIMain) = $IDYES Then
						_RunFile(@ScriptFullPath)
						Exit

					EndIf

				Case $MenuShowAllScriptsButton
					_Log("MenuUpdateButton")
					_PopulateScripts($ScriptsTree, "AutoLogin")
					_PopulateScripts($ScriptsTree, "AutoCustom")
					_PopulateScripts($ScriptsTree, "AutoSystem")
					_PopulateScripts($ScriptsTree, "OptAdvanced")

				Case $MenuOpenFolder
					_Log("MenuOpenFolder")
					ShellExecute(@ScriptDir)

				Case $MenuOpenLog
					_Log("MenuOpenLog")
					ShellExecute($LogFullPath)

				Case $MenuVisitGitButton
					_Log("Opening Browser...", True)
					$o_URL = ObjCreate("Shell.Application")
					$o_URL.Open($GITURL)

				Case $JoinButton
					Run("SystemPropertiesComputerName.exe")
					$hWindow = WinWait("System Properties")
					ControlClick($hWindow, "", "[CLASS:Button; INSTANCE:2]")

				Case $CreateLocalUserButton
					$sUser = GUICtrlRead($UsernameInput)
					$sPassword = GUICtrlRead($PasswordInput)
					$Admin = GUICtrlRead($AdminCheckBox)

					If $sUser <> "" Then
						$objSystem = ObjGet("WinNT://localhost")
						$objUser = $objSystem.Create("user", $sUser)
						$objUser.SetPassword($sPassword)
						$objUser.Put("UserFlags", BitOR($objUser.get("UserFlags"), 0x10000))
						$objUser.SetInfo
						If Not @error And $Admin = $GUI_CHECKED Then
							$objGroup = ObjGet("WinNT://localhost/Administrators")
							$objGroup.Add("WinNT://" & $sUser)
						EndIf

						If Not IsObj(ObjGet("WinNT://./" & $sUser & ", user")) Then
							MsgBox($MB_ICONWARNING, $Title, "Error creating user", 0, $GUIMain)
							_Log("Error Creating User", True)
						Else
							If $Admin = $GUI_CHECKED Then $UserCreatedWithAdmin = True
							_Log("User Created Successfully", True)
						EndIf
					Else
						_Log("Missing Username", True)

					EndIf

			EndSwitch
		WEnd

	Case Else
		_Log("Command unknown")

EndSwitch

Func _PopulateScripts($TreeID, $Folder)
	Local $FileArray = _FileListToArray(@ScriptDir & "\" & $Folder & "\", "*", $FLTA_FILES, True)

	If Not @error Then
		_Log($Folder & " Files (no filter): " & $FileArray[0])
		Local $FolderTreeItem = GUICtrlCreateTreeViewItem($Folder, $TreeID)
		GUICtrlSetState($FolderTreeItem, $GUI_CHECKED)

		For $i = 1 To $FileArray[0]
			If StringInStr($FileArray[$i], "\.") Then ContinueLoop
			_Log("Added: " & $FileArray[$i])
			$FileName = StringTrimLeft($FileArray[$i], StringInStr($FileArray[$i], "\", 0, -1))
			GUICtrlCreateTreeViewItem($FileName, $FolderTreeItem)
		Next

		GUICtrlSetState($FolderTreeItem, $GUI_EXPAND)
		GUICtrlSetState($FolderTreeItem, $GUI_UNCHECKED)

		Return $FileArray
	Else
		_Log("No files")
		Return 0
	EndIf

EndFunc   ;==>_PopulateScripts

Func _NotAdminMsg($hwnd = "")
	_Log("_NotAdminMsg")
	MsgBox($MB_OK, $Title, "Not running with admin rights.", 0, $hwnd)

EndFunc   ;==>_NotAdminMsg

Func _RunTreeView($hWindow, $hTreeView, $Dry = False)
	_Log("_RunTreeView")

	Local $aList[0]

	For $iTop = 0 To ControlTreeView($GUIBoot, "", $Boot_ScriptsTree, "GetItemCount", "") - 1
		$Folder = ControlTreeView($GUIBoot, "", $Boot_ScriptsTree, "GetText", "#" & $iTop)

		For $iSub = 0 To ControlTreeView($GUIBoot, "", $Boot_ScriptsTree, "GetItemCount", "#" & $iTop) - 1
			$File = ControlTreeView($GUIBoot, "", $Boot_ScriptsTree, "GetText", "#" & $iTop & "|#" & $iSub)
			$FileChecked = ControlTreeView($GUIBoot, "", $Boot_ScriptsTree, "IsChecked", "#" & $iTop & "|#" & $iSub)

			If $FileChecked Then
				$RunFullPath = @ScriptDir & "\" & $Folder & "\" & $File
				_Log("Checked: $RunFullPath=" & $RunFullPath)
				If $Dry = False Then _RunFile($RunFullPath)
				_ArrayAdd($aList, $RunFullPath)
			EndIf
		Next

	Next

	Return $aList

EndFunc   ;==>_RunTreeView

Func _RunFolder($Path)
	_Log("_RunFolder " & $Path)
	$FileArray = _FileListToArray($Path, "*", $FLTA_FILES, True)
	If Not @error Then
		_Log("Files: " & $FileArray[0])
		For $i = 1 To $FileArray[0]
			If StringInStr($FileArray[$i], "\.") Then ContinueLoop
			_Log($FileArray[$i])
			_RunFile($FileArray[$i])
		Next
		Return $FileArray[0]
	Else
		_Log("No files")
	EndIf

EndFunc   ;==>_RunFolder

Func _RunFile($File, $Params = "", $WorkingDir = "")
	_Log("_RunFile " & $File)
	$Extension = StringTrimLeft($File, StringInStr($File, ".", 0, -1))
	Switch $Extension
		Case "au3", "a3x"
			_Log("  au3")
			$RunLine = @AutoItExe & " /AutoIt3ExecuteScript """ & $File & """ " & $Params
			;Return ShellExecute(@AutoItExe, "/AutoIt3ExecuteScript """ & $File & """ " & $Params)
			Return Run($RunLine, $WorkingDir, @SW_SHOW, $STDIO_INHERIT_PARENT)

		Case "ps1"
			_Log("  ps1")
			;$File = StringReplace($File, "$", "`$")
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

Func _GitUpdate($Prompt = False)
	_Log("_GitUpdate")
	Local $Current = _RecSizeAndHash(@ScriptDir)
	Local $TempPath = _TempFile(Default, "itdeploy-", ".tmp")
	Local $TempZIP = $TempPath & "\itdeploy.zip"
	Local $TempPathExtracted = $TempPath & "\itdeployhelper-master"
	Local $aChanges[0][3]

	DirCreate($TempPath)

	Local $DownloadSize = InetGet($GITZIP, $TempZIP, $INET_FORCERELOAD)
	If @error Then
		_Log("Download Error " & @error, True)
		Return 0
	EndIf
	_Log("ZIP Download Size: " & $DownloadSize)

	;Extract zip
	_Zip_UnzipAll($TempZIP, $TempPath, 16 + 1024)
	If @error Then
		_Log("Unzip Error " & @error, True)
		Return 0
	EndIf

	Local $New = _RecSizeAndHash($TempPathExtracted)


	;Look for files that were changed or removed
	For $i = 0 To UBound($Current) - 1
		$Found = _ArraySearch($New, $Current[$i][0])
		If $Found >= 0 Then
			If $Current[$i][2] <> $New[$Found][2] Then
				_Log("Changed: " & $Current[$i][0])
				_ArrayAdd($aChanges, $Current[$i][0] & "|" & $Current[$i][1] & "|" & $New[$Found][1])
			EndIf
		Else

			If StringInStr($Current[$i][0], "\AutoLogin") Or StringInStr($Current[$i][0], "\OptLogin") Then
				_Log("Removed: " & $Current[$i][0])
				_ArrayAdd($aChanges, $Current[$i][0] & "|" & $Current[$i][1] & "|" & "(Removed)")
			EndIf
		EndIf
	Next

	;Look for files that were added
	For $i = 0 To UBound($New) - 1
		$Found = _ArraySearch($Current, $New[$i][0])
		If $Found = -1 Then
			_Log("Added: " & $New[$i][0])
			_ArrayAdd($aChanges, $New[$i][0] & "|" & "(Added)" & "|" & $New[$i][1])
		EndIf
	Next

	Local $ChangesCount = UBound($aChanges)
	Local $ChangesString = _ArrayToString($aChanges, ", ", Default, Default, @CRLF)
	_Log("Changes: " & $ChangesCount)

	If $ChangesCount = 0 Then
		FileDelete($TempZIP)
		DirRemove($TempPath, $DIR_REMOVE)
		Return $aChanges
	EndIf

	If $Prompt Then
		If MsgBox($MB_YESNO, $Title, "Apply the following changes?" & @CRLF & @CRLF & "File Name, Old Size, New Size" & @CRLF & $ChangesString) <> $IDYES Then
			FileDelete($TempZIP)
			DirRemove($TempPath, $DIR_REMOVE)
			SetError(1)
			Return $aChanges
		EndIf
	EndIf

	If FileExists($TempPathExtracted & "\AutoLogin") Then FileDelete(@ScriptDir & "\AutoLogin")
	If FileExists($TempPathExtracted & "\OptLogin") Then FileDelete(@ScriptDir & "\OptLogin")

	Local $CopyStatus = DirCopy($TempPathExtracted, @ScriptDir, $FC_OVERWRITE)
	_Log("Copied Files (" & $CopyStatus & ")")

	FileDelete($TempZIP)
	DirRemove($TempPath, $DIR_REMOVE)
	Return $aChanges

EndFunc   ;==>_GitUpdate

Func _RecSizeAndHash($Path) ; Return Array with RelativePath|Size|MD5
	_Log("_RecSizeAndHash - " & $Path)
	Local $aOutput[0][3]

	If StringRight($Path, 1) = "\" Then $Path = StringTrimRight($Path, 1)
	Local $aFiles = _FileListToArrayRec($Path, "*", $FLTAR_FILES + $FLTAR_NOHIDDEN + $FLTAR_NOSYSTEM + $FLTAR_NOLINK, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_RELPATH)

	If Not @error Then
		For $i = 1 To $aFiles[0]
			$ThisFileRelPath = $aFiles[$i]
			$ThisFileFullPath = $Path & "\" & $ThisFileRelPath
			$ThisSize = FileGetSize($ThisFileFullPath)
			$ThisHash = _Crypt_HashFile($ThisFileFullPath, $CALG_MD5)
			_ArrayAdd($aOutput, $ThisFileRelPath & "|" & $ThisSize & "|" & $ThisHash, 0, "|")
		Next
	EndIf

	Return $aOutput

EndFunc   ;==>_RecSizeAndHash

Func _Log($Message, $Statusbar = "")
	Local $sTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "> " ; Generate Timestamp
	ConsoleWrite($sTime & $Message & @CRLF)
	If $Statusbar Then _GUICtrlStatusBar_SetText($StatusBar1, $Message)

	FileWrite($LogFullPath, $sTime & $Message & @CRLF)
	Return $Message
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
