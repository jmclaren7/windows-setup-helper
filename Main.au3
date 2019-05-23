#include <WinAPI.au3>
#include <File.au3>
#include <Process.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <TabConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include "includeExt\Json.au3"

OnAutoItExitRegister("_Exit")

Global $Log = @ScriptDir & "\ITSetupLog.txt"
Global $Title = "IT Setup Helper"

_Log("Start script " & $CmdLineRaw)
_Log("Username: " & @UserName)
_Log("@ScriptFullPath: " & @ScriptFullPath)

If $CmdLine[0] >= 1 Then
	$Command = $CmdLine[1]
Else
	$Command = ""
EndIf

Switch $Command
	Case "system"
		_RunFolder(@ScriptDir & "\AutoSystem\")

	Case "login"
		FileCreateShortcut(@AutoItExe, @DesktopDir & "\IT Setup Helper.lnk", @ScriptDir, "/AutoIt3ExecuteScript """ & @ScriptFullPath & """")
		FileCreateShortcut(@ScriptDir, @DesktopDir & "\IT Setup Folder")

		ProcessWait("Explorer.exe", 60)
		Sleep(5000)

		_RunFolder(@ScriptDir & "\AutoLogin\")
		ShellExecute(@AutoItExe, "/AutoIt3ExecuteScript """ & @ScriptFullPath & """")

	Case ""
		#Region ### START Koda GUI section ### Form=g:\windows 10 images\1809\windows 10 x64 12-21-18 1809\sources\$oem$\$$\it\general.kxf
		$Form1 = GUICreate("Form1", 824, 574, 353, 483)
		$Tab1 = GUICtrlCreateTab(8, 4, 809, 561)
		$TabSheet1 = GUICtrlCreateTabItem("Main")
		$Group1 = GUICtrlCreateGroup("Action Scripts", 400, 33, 401, 521)
		$Presets = GUICtrlCreateCombo("Presets", 416, 57, 369, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
		GUICtrlSetState(-1, $GUI_DISABLE)
		$TreeView1 = GUICtrlCreateTreeView(416, 97, 369, 417, BitOR($GUI_SS_DEFAULT_TREEVIEW,$TVS_CHECKBOXES))
		$RunButton = GUICtrlCreateButton("Run", 712, 521, 75, 25)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group2 = GUICtrlCreateGroup("Information", 23, 32, 361, 385)
		$InfoList = GUICtrlCreateListView("", 32, 50, 346, 358, BitOR($GUI_SS_DEFAULT_LISTVIEW,$LVS_SMALLICON), 0)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group3 = GUICtrlCreateGroup("Actions", 24, 426, 361, 129)
		$DisableAdminButton = GUICtrlCreateButton("Disable Administrator", 37, 521, 155, 25)
		$SignOutButton = GUICtrlCreateButton("Sign Out", 217, 521, 155, 25)
		$UpdateButton = GUICtrlCreateButton("Update From Git", 37, 448, 155, 25)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		GUICtrlCreateTabItem("")
		GUISetState(@SW_SHOW)
		#EndRegion ### END Koda GUI section ###

		;GUI Post Creation Setup
		WinSetTitle($Form1, "", $Title)
		$WindowPos = WinGetPos($Form1)
		WinMove($Form1, "", @DesktopWidth / 2 - $WindowPos[2] / 2, @DesktopHeight / 2 - $WindowPos[3] / 2)

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

		;Generate Script List
		$FileArray = _FileListToArray(@ScriptDir & "\OptLogin\", "*", $FLTA_FILES, True)
		If Not @error Then
			Local $OptLoginListItems[$FileArray[0] + 1]
			_Log("Files: " & $FileArray[0])
			For $i = 1 To $FileArray[0]
				_Log($FileArray[$i])
				$FileName = StringTrimLeft($FileArray[$i], StringInStr($FileArray[$i], "\", 0, -1))
				$OptLoginListItems[$i] = GUICtrlCreateTreeViewItem($FileName, $TreeView1)

			Next
		Else
			_Log("No files")
		EndIf

		;$TabSheet2 = GUICtrlCreateTabItem("Test")
		;$Groupb = GUICtrlCreateGroup("Action Scripts", 400, 33, 401, 521)
		GUISetState(@SW_HIDE)
		GUISetState(@SW_SHOW)

		;GUI Loop
		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $GUI_EVENT_CLOSE
					Exit

				Case $DisableAdminButton
					_Log("DisableAdminButton")

					If @ComputerName = @LogonDomain And MsgBox($MB_YESNO, $Title, "Computer does not seem to be joined to a domain, are you sure you want disable the administrator account?") <> $IDYES Then ContinueLoop

					If IsAdmin() Then
						_Log("Disable admin command")
						Run(@ComSpec & " /c " & 'net user administrator /active:no', "", @SW_SHOW)
					Else
						_NotAdminMsg($Form1)
					EndIf

				Case $RunButton
					_Log("RunButton")
					For $x = 1 To UBound($OptLoginListItems) - 1
						If BitAND(GUICtrlRead($OptLoginListItems[$x]), $GUI_CHECKED) Then
							_Log("Checked: " & $FileArray[$x])
							_RunFile($FileArray[$x])

						EndIf
					Next

				Case $UpdateButton
					_DownloadGit("https://api.github.com/repos/jmclaren7/itdeployhelper/contents", @ScriptDir)

				Case $SignOutButton
					Run(@ComSpec & " /c " & 'logoff', "", @SW_SHOW)

			EndSwitch
		WEnd

	Case Else
		_Log("Command unknown")

EndSwitch


Func _NotAdminMsg($hwnd = "")
	_Log("_NotAdminMsg")
	MsgBox($MB_OK, $Title, "Not running with admin rights.", 0, $hwnd)
EndFunc   ;==>_NotAdminMsg

Func _RunFolder($Path)
	_Log("_RunFolder " & $Path)
	$FileArray = _FileListToArray($Path, "*", $FLTA_FILES, True)
	If Not @error Then
		_Log("Files: " & $FileArray[0])
		For $i = 1 To $FileArray[0]
			_Log($FileArray[$i])
			_RunFile($FileArray[$i])
		Next
		Return $FileArray[0]
	Else
		_Log("No files")
	EndIf
EndFunc   ;==>_RunFolder

Func _RunFile($File)
	_Log("_RunFile " & $File)
	$Extension = StringTrimLeft($File, StringInStr($File, ".", 0, -1))
	Switch $Extension
		Case "au3"
			Return ShellExecute(@AutoItExe, "/AutoIt3ExecuteScript """ & $File & """")
			Sleep(5000)

		Case "ps1"
			;$File = StringReplace($File, "$", "`$")
			$RunLine = @ComSpec & " /c " & "powershell.exe -ExecutionPolicy Unrestricted -File """ & $File & """"
			_Log("$RunLine=" & $RunLine)
			Return Run($RunLine)

		Case "reg"
			$RunLine = @ComSpec & " /c " & "reg import """ & $File & """"
			_Log("$RunLine=" & $RunLine)
			Return Run($RunLine)

		Case Else
			Return ShellExecute($File)

	EndSwitch
EndFunc   ;==>_RunFile

Func _DownloadGit($URL, $Destination)
	_Log($URL)
	$Data = BinaryToString(InetRead($URL, 1))
	$Object = json_decode($Data)

	Local $i = 0
	While 1
		$Name = json_get($Object, '[' & $i & '].name')
		If @error Then ExitLoop
		_Log($Name)

		$Path = json_get($Object, '[' & $i & '].path')
		$Type = json_get($Object, '[' & $i & '].type')
		$URL = json_get($Object, '[' & $i & '].url')
		$Download_url = json_get($Object, '[' & $i & '].download_url')

		$FullPath = $Destination&"\"&StringReplace($Path, "/", "\")
		$FolderPath = StringLeft($FullPath, StringInStr($FullPath, "\", 0, -1))
		_Log("$FullPath="&$FullPath)
		_Log("$FolderPath="&$FolderPath)

		If $type = "dir" Then
			_Log("Dir - "&$URL)
			;recurse
			_DownloadGit($URL, $Destination)

		Else
			;download
			DirCreate($FolderPath)
			InetGet($Download_url, $FullPath, 1)
		endif

		$i += 1
	WEnd

EndFunc

Func _Log($Message)
	Local $sTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "> " ; Generate Timestamp
	ConsoleWrite($sTime & $Message & @CRLF)
	FileWrite(@ScriptFullPath&"_log.txt", $sTime & $Message & @CRLF)
	Return $Message
EndFunc   ;==>_Log

Func _Exit()
	_Log("End script " & $CmdLineRaw)

EndFunc   ;==>_Exit
