#include "include\ButtonConstants.au3"
#include "include\ComboConstants.au3"
#include "include\Crypt.au3"
#include "include\EditConstants.au3"
#include "include\File.au3"
#include "include\FileConstants.au3"
#include "include\GuiConstantsEx.au3"
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
#include "include\WinAPIFiles.au3"
#include "includeExt\Json.au3"
#include "includeExt\WinHttp.au3"
#include "includeExt\ActivationStatus.au3"
#include "includeExt\Custom.au3"
#include "includeExt\_Zip.au3"


OnAutoItExitRegister("_Exit")
_WinAPI_Wow64EnableWow64FsRedirection(False)

Global $LogFullPath = StringReplace(@TempDir & "\" & @ScriptName, ".au3", ".log")
Global $MainSize = FileGetSize(@ScriptFullPath)
Global $Version = "4.0.2." & $MainSize

Global $Title = "Windows Setup Helper v" & $Version
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

Global $TokenAddHeader = IniRead(".token", "t", "t", "")
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
	Case "system"
		_RunMulti("AutoSystem")

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

		_RunMulti("AutoLogin")
		_RunFile(@ScriptFullPath)

	Case "main-gui", "boot-gui"
		; Mode specific
		If $Command="boot-gui" Then
			; Hide console windows
			Sleep(500)
			WinSetState(@ComSpec, "", @SW_HIDE)
			Sleep(500)
			WinSetState(@ComSpec, "", @SW_HIDE) ;Something prevents it working if ran only once or too fast?
			WinSetState("winpehelper.exe", "", @SW_HIDE)

		EndIf

		#Region ### START Koda GUI section ### Form=main.kxf
		$GUIMain = GUICreate("$Title", 767, 543, -1, -1)
		$MenuItem2 = GUICtrlCreateMenu("&File")
		$MenuExitButton = GUICtrlCreateMenuItem("Exit", $MenuItem2)
		$MenuItem1 = GUICtrlCreateMenu("&Advanced")
		$MenuUpdateButton = GUICtrlCreateMenuItem("Update from GitHub", $MenuItem1)
		$MenuVisitGitButton = GUICtrlCreateMenuItem("Visit GitHub Page", $MenuItem1)
		$MenuShowConsole = GUICtrlCreateMenuItem("Show Console", $MenuItem1)
		$MenuShowAllScriptsButton = GUICtrlCreateMenuItem("Show All Scripts", $MenuItem1)
		$MenuOpenLog = GUICtrlCreateMenuItem("Open Log", $MenuItem1)
		$MenuOpenFolder = GUICtrlCreateMenuItem("Open Program Folder", $MenuItem1)
		$SwitchTabsMenu = GUICtrlCreateMenuItem("Switch Mode", $MenuItem1)
		$Tab1 = GUICtrlCreateTab(7, 4, 753, 495)
		$MainTabSheet = GUICtrlCreateTabItem("&")
		$Group1 = GUICtrlCreateGroup("Scripts", 393, 33, 356, 452)
		$ScriptsTreeView = GUICtrlCreateTreeView(403, 57, 335, 382, BitOR($GUI_SS_DEFAULT_TREEVIEW,$TVS_CHECKBOXES))
		$RunButton = GUICtrlCreateButton("Run", 633, 449, 107, 25)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group2 = GUICtrlCreateGroup("Information", 22, 32, 361, 257)
		$InfoList = GUICtrlCreateListView("", 27, 50, 350, 230, BitOR($GUI_SS_DEFAULT_LISTVIEW,$LVS_SMALLICON), 0)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group3 = GUICtrlCreateGroup("Create Local User", 22, 396, 361, 89)
		$CreateLocalUserButton = GUICtrlCreateButton("Create Local User", 235, 448, 131, 25)
		$UsernameInput = GUICtrlCreateInput("", 38, 418, 185, 21)
		$PasswordInput = GUICtrlCreateInput("", 38, 450, 185, 21)
		$AdminCheckBox = GUICtrlCreateCheckbox("Local Administrator", 238, 420, 113, 17)
		GUICtrlSetState(-1, $GUI_CHECKED)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group4 = GUICtrlCreateGroup("Actions", 22, 294, 361, 97)
		$JoinButton = GUICtrlCreateButton("Domain && Computer Name", 35, 319, 160, 25)
		$DisableAdminButton = GUICtrlCreateButton("Disable Administrator", 211, 319, 160, 25)
		$SignOutButton = GUICtrlCreateButton("Sign Out", 35, 354, 160, 25)
		$DeleteHelperButton = GUICtrlCreateButton("Delete Install Helper", 211, 354, 160, 25)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
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
		_GUICtrlStatusBar_SetSimple($StatusBar1)
		_GUICtrlStatusBar_SetText($StatusBar1, "")
		GUISetState(@SW_SHOW)
		#EndRegion ### END Koda GUI section ###

		; GUI Post Creation Setup
		WinSetTitle($GUIMain, "", $Title)
		GUICtrlSendMsg($UsernameInput, $EM_SETCUEBANNER, False, "Username")
		GUICtrlSendMsg($PasswordInput, $EM_SETCUEBANNER, False, "Password (optional)")

		; Info List Generation
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

		; Generate Script List
		_PopulateScripts($ScriptsTreeView, "OptLogin")
		_PopulateScripts($PEInstallTreeView, "OptLogin")
		_PopulateScripts($PEScriptTreeView, "OptSetup")

		Local $hSetup
		Local $CopyOptFiles = False
		Local $DeleteOEMFiles = False
		Local $Reboot = False
		Local $BootDrive = StringLeft(@SystemDir, 3)

		; Tab/Mode specific
		If $Command="boot-gui" Then
			; Start network
			Run(@ComSpec & " /c " & 'wpeinit.exe', @SystemDir, @SW_HIDE, $RUN_CREATE_NEW_CONSOLE)

			; Set GUI Icon
			GUISetIcon($BootDrive & "sources\setup.exe")

			; Run automatic setup scripts
			_RunMulti("AutoSetup")

			; ?
			Opt("WinTitleMatchMode", 2)
		EndIf

		_Log("Ready", True)

		;GUI Loop
		While 1
			$CurrentTab=GUICtrlRead($Tab1)
			If $Command="boot-gui" AND $CurrentTab=0 Then
				GUICtrlSetState($BootTabSheet, $GUI_SHOW)
			Elseif $Command="main-gui" AND $CurrentTab=1 Then
				GUICtrlSetState($MainTabSheet, $GUI_SHOW)
			Endif

			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $GUI_EVENT_CLOSE, $MenuExitButton
					If $Command="boot-gui" AND MsgBox(1, $Title, "Closing the program will reboot the system while in WinPE.") <> 1 Then ContinueLoop
					Exit

				Case $SwitchTabsMenu
					If $Command="boot-gui" Then
						$Command="main-gui"
					Elseif $Command="main-gui" Then
						$Command="boot-gui"
					Endif

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

				Case $DeleteHelperButton
					_Log("CleaupButton")
					FileDelete(@DesktopDir & "\IT Setup Folder.lnk")
					FileDelete(@DesktopDir & "\IT Setup Helper.lnk")
					If FileExists(@WindowsDir & "\IT\Main.au3") Then
						Run(@ComSpec & " /c " & "ping 127.0.0.1 -n 3 && rmdir /s /q " & @WindowsDir & "\IT")
					EndIf
					ProcessClose(@AutoItPID)

				Case $RunButton
					_Log("RunButton")
					_RunTreeView($GUIMain, $ScriptsTreeView)

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
					_Log("MenuShowAllScriptsButton")
					_PopulateScripts($ScriptsTreeView, "*")
					_PopulateScripts($PEScriptTreeView, "*")
					_PopulateScripts($PEInstallTreeView, "*")

				Case $MenuShowConsole
					WinSetState(@ComSpec, "", @SW_SHOW)
					WinSetState("winpehelper.exe", "", @SW_SHOW)

				Case $MenuOpenFolder
					_Log("MenuOpenFolder")
					_RunFile("explorer.exe", @ScriptDir)

				Case $MenuOpenLog
					_Log("MenuOpenLog")
					_RunFile($LogFullPath)

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

				Case $PERunButton
					_RunTreeView($GUIMain, $PEScriptTreeView)

				Case $NormalInstallButton
					$hSetup = _RunFile($BootDrive & "sources\setup.exe")
					$DeleteOEMFiles = True
					$CopyOptFiles = False

				Case $AutomatedInstallButton
					$aList = _RunTreeView($GUIMain, $PEInstallTreeView, True)
					For $b = 0 To UBound($aList) - 1
						_Log("TreeItem: " & $aList[$b])
					Next

					$AutounattendPath = @ScriptDir & "\autounattend.xml"
					$ComputerName = GUICtrlRead($PEComputerNameInput)

					If $ComputerName <> "" Then
						_Log("$ComputerName=" & $ComputerName)
						$AutounattendPath_New = @TempDir & "\autounattend.xml"
						$sFileData = FileRead($AutounattendPath)
						$sFileData = StringReplace($sFileData, "<ComputerName>*</ComputerName>", "<ComputerName>"&$ComputerName&"</ComputerName>")
						_Log("StringReplace @extended=" & @extended)

						$hAutounattend = FileOpen($AutounattendPath_New, $FO_OVERWRITE)
						FileWrite($hAutounattend, $sFileData)
						_Log("FileWrite @error=" & @error)
						FileClose($hAutounattend)

						$AutounattendPath = $AutounattendPath_New
					EndIf

					_Log("$AutounattendPath=" & $AutounattendPath)
					$hSetup = _RunFile($BootDrive & "sources\setup.exe", "/noreboot /unattend:" & $AutounattendPath)
					$CopyOptFiles = True
					$DeleteOEMFiles = False

				EndSwitch

				If $CopyOptFiles AND NOT ProcessExists($hSetup) Then
					_Log("CopyOptFiles")
					For $i = 65 To 90
						$Path = Chr($i) & ":\Windows\IT"
						If FileExists($Path) Then
							$Dest = $Path & "\AutoLogin\"
							_Log("Found: " & $Path)
							For $iFile = 0 To UBound($aList) - 1
								$Return = FileCopy($aList[$iFile], $Dest, 1)
								_Log("FileCopy: " & $aList[$iFile] & " (" & $Return & ")")
							Next

							FileCopy($LogFullPath, $Path & "\pe.log")

						EndIf
					Next

					$Reboot = True
					$CopyOptFiles = False
				Endif

				If $DeleteOEMFiles AND NOT ProcessExists($hSetup) Then
					_Log("DeleteOEMFiles")
					For $i = 65 To 90
						$Path = Chr($i) & ":\Windows\IT"
						If FileExists($Path) Then
							_Log("Found: " & $Path)
							$Return = DirRemove($Path, 1)
							_Log("DirRemove: " & $Path & " (" & $Return & ")")
							Sleep(1000)
						EndIf
					Next

					$Reboot = True
					$DeleteOEMFiles = False
				EndIf

				If $Reboot Then
					_Log("Reboot")
					Beep(500, 1000)
					$Return = Msgbox(1, $Title, "Rebooting in 15 seconds", 15)
					If $Return = $IDTIMEOUT OR $Return = $IDOK Then Exit
					$Reboot = False
				Endif

		WEnd

	Case Else
		_Log("Command unknown")

EndSwitch

Func _PopulateScripts($TreeID, $Folder)
	_Log("_PopulateScripts " & $Folder)

	Local $FileArray[0]

	If $Folder = "*" Then
		_Log("Wildcard")
		_GUICtrlTreeView_DeleteAll($TreeID)

		Local $aList = _FileListToArrayRec(@ScriptDir, "auto*;opt*|*custom", $FLTAR_FOLDERS, $FLTAR_NORECUR, $FLTAR_NOSORT, $FLTAR_NOPATH)
		_Log("$aList=" & _ArrayToString($aList))

		Local $aFileList

		For $f=1 to $aList[0]
			$List = _PopulateScripts($TreeID, $aList[$f])
			_ArrayConcatenate($FileArray, $List, 1)
		Next

		Return $FileArray

	Else
		$FileArray = _FileListToArray(@ScriptDir & "\" & $Folder & "\", "*", $FLTA_FILESFOLDERS, True) ;switched from $FLTA_FILES for allowing main.au3 in folder
		If Not @error Then
			_Log($Folder & " Files (no filter): " & $FileArray[0])
			Local $FolderTreeItem = GUICtrlCreateTreeViewItem($Folder, $TreeID)

			For $i = 1 To $FileArray[0]
				If StringInStr($FileArray[$i], "\.") Then ContinueLoop ;Use . for hidden
				If StringInStr(FileGetAttrib($FileArray[$i]), "D") And NOT FileExists($FileArray[$i] & "\main.au3") Then ContinueLoop ;allow folders if then contain main.au3
				$FileName = StringTrimLeft($FileArray[$i], StringInStr($FileArray[$i], "\", 0, -1))
				GUICtrlCreateTreeViewItem($FileName, $FolderTreeItem)
				_Log("Added: " & $FileArray[$i])
			Next

			GUICtrlSetState($FolderTreeItem, $GUI_EXPAND)

		Else
			_Log($Folder & " No files or missing")

		EndIf

		If StringRight($Folder, 6) <> "Custom" Then _PopulateScripts($TreeID, $Folder & "Custom")
	EndIf

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

	$aFiles1 = _RunFolder(@ScriptDir & "\" & $Folder )
	$aFiles2 = _RunFolder(@ScriptDir & "\" & $Folder & "Custom")

	$iCount = _ArrayConcatenate ($aFiles1, $aFiles2, 1)

	;$aFiles1[0] = $iCount

	Return $aFiles1
EndFunc

Func _RunFolder($Path)
	_Log("_RunFolder " & $Path)
	$FileArray = _FileListToArray($Path, "*", $FLTA_FILESFOLDERS, True) ;switched from $FLTA_FILES for allowing main.au3 in folder
	If Not @error Then
		_Log("Files: " & $FileArray[0])
		For $i = 1 To $FileArray[0]
			If StringInStr($FileArray[$i], "\.") Then ContinueLoop
			If StringInStr(FileGetAttrib($FileArray[$i]), "D") And NOT FileExists($FileArray[$i] & "\main.au3") Then ContinueLoop
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
