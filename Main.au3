#include <WinAPI.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <Inet.au3>
#include <InetConstants.au3>
#include <Process.au3>
#include <Crypt.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <GuiTreeView.au3>
#include <GuiStatusBar.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <TabConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include "includeExt\Json.au3"
#include "includeExt\WinHttp.au3"
#include "includeExt\ActivationStatus.au3"
#include "includeExt\Custom.au3"

OnAutoItExitRegister("_Exit")

Global $Log = @ScriptDir & "\ITSetupLog.txt"
Global $MainSize = FileGetSize(@ScriptFullPath)
Global $Version = "2.3.0-"&$MainSize

Global $Title = "IT Setup Helper v"&$Version
Global $DownloadUpdatedCount = 0
Global $DownloadErrors = 0
Global $DownloadUpdated = ""
Global $GITURL = "https://github.com/jmclaren7/itdeployhelper"
Global $GITAPIURL = "https://api.github.com/repos/jmclaren7/itdeployhelper/contents"
Global $GUIMain
Global $oCommError = ObjEvent("AutoIt.Error","_CommError")
Global $StatusBar1
$UserCreated = False

_Log("Start Script " & $CmdLineRaw)
_Log("@UserName=" & @UserName)
_Log("@ScriptFullPath=" & @ScriptFullPath)

If $CmdLine[0] >= 1 Then
	$Command = $CmdLine[1]
Else
	$Command = ""
EndIf

Switch $Command
	Case "system"
		_RunFolder(@ScriptDir & "\AutoSystem\")

	Case "login"
		ProcessWait("Explorer.exe", 60)
		Sleep(5000)

		If Not StringInStr($CmdLineRaw,"skipupdate") Then
			_DownloadGitSetup($GITAPIURL, @ScriptDir)
			If StringInStr($DownloadUpdated, @ScriptName) Then
				_RunFile(@ScriptFullPath, "login skipupdate")
				Exit
			Endif
		Endif

		FileCreateShortcut(@AutoItExe, @DesktopDir & "\IT Setup Helper.lnk", @ScriptDir, "/AutoIt3ExecuteScript """ & @ScriptFullPath & """")
		FileCreateShortcut(@ScriptDir, @DesktopDir & "\IT Setup Folder")

		WinMinimizeAll ( )

		_RunFolder(@ScriptDir & "\AutoLogin\")
		_RunFile(@ScriptFullPath)

	Case ""
		#Region ### START Koda GUI section ###
		$GUIMain = GUICreate("$Title", 823, 574, -1, -1)
		$MenuItem2 = GUICtrlCreateMenu("File")
		$MenuExitButton = GUICtrlCreateMenuItem("Exit", $MenuItem2)
		$MenuItem1 = GUICtrlCreateMenu("Advanced")
		$MenuUpdateButton = GUICtrlCreateMenuItem("Update from GitHub", $MenuItem1)
		$MenuVisitGitButton = GUICtrlCreateMenuItem("Go To GitHub Page", $MenuItem1)
		$MenuShowLoginScriptsButton = GUICtrlCreateMenuItem("Show Login Scripts", $MenuItem1)
		$Tab1 = GUICtrlCreateTab(7, 4, 809, 521)
		$TabSheet1 = GUICtrlCreateTabItem("Main")
		$Group1 = GUICtrlCreateGroup("Scripts", 399, 33, 401, 481)
		$Presets = GUICtrlCreateCombo("Presets", 415, 57, 369, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
		GUICtrlSetState(-1, $GUI_DISABLE)
		$ScriptsTree = GUICtrlCreateTreeView(415, 97, 369, 369, BitOR($GUI_SS_DEFAULT_TREEVIEW,$TVS_CHECKBOXES))
		$RunButton = GUICtrlCreateButton("Run", 711, 481, 75, 25)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$Group2 = GUICtrlCreateGroup("Information", 22, 32, 361, 257)
		$InfoList = GUICtrlCreateListView("", 31, 50, 346, 230, BitOR($GUI_SS_DEFAULT_LISTVIEW,$LVS_SMALLICON), 0)
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
		$Manufacturer = RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS","SystemManufacturer")
		If $Manufacturer = "System manufacturer" Then $Manufacturer = "Unknown"
		GUICtrlCreateListViewItem("Manufacturer: " & $Manufacturer, $InfoList)
		GUICtrlCreateListViewItem("Model: " & RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "SystemProductName"), $InfoList)
		GUICtrlCreateListViewItem("BIOS: " & RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "BIOSVersion"), $InfoList)
		GUICtrlCreateListViewItem("CPU Logical Cores: " & EnvGet("NUMBER_OF_PROCESSORS"), $InfoList)
		$MemStats = MemGetStats()
		GUICtrlCreateListViewItem("Installed Memory: " & Round($MemStats[$MEM_TOTALPHYSRAM]/1024/1024,1)&"GB", $InfoList)
		GUICtrlCreateListViewItem("License: " & IsActivated(), $InfoList)
		$NetInfo = _NetAdapterInfo()
		GUICtrlCreateListViewItem("IP/Gateway: " & $NetInfo[3] & "/" & $NetInfo[4], $InfoList)
		GUICtrlCreateListViewItem("MAC: " & $NetInfo[2], $InfoList)


		;Generate Script List
		_PopulateScripts($ScriptsTree, "OptLogin")
		_PopulateScripts($ScriptsTree, "OptCustom")

		;$TabSheet2 = GUICtrlCreateTabItem("Test")
		;$GroupSheet2 = GUICtrlCreateGroup("Test Group", 400, 33, 401, 521)
		;GUISetState(@SW_HIDE)
		;GUISetState(@SW_SHOW)

		_Log("Ready", True)

		;GUI Loop
		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $GUI_EVENT_CLOSE
					Exit

				Case $DisableAdminButton
					_Log("DisableAdminButton")

					If @ComputerName = @LogonDomain AND Not $UserCreated Then
						If MsgBox($MB_YESNO, $Title, "Are you sure?"&@CRLF&@CRLF&"This computer might not be joined to a domain and you haven't created a local user.", 0, $GUIMain) <> $IDYES Then
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
					Shutdown (0)

				Case $RunButton
					_Log("RunButton")
					$TreeViewItemTotal = _GUICtrlTreeView_GetCount ($ScriptsTree)

					For $TreeItemCount = 1 To $TreeViewItemTotal
						If Not IsDeclared("hScriptsTreeItem") Then
							$hScriptsTreeItem = _GUICtrlTreeView_GetFirstItem ($ScriptsTree)
						Else
							$hScriptsTreeItem = _GUICtrlTreeView_GetNext ( $ScriptsTree, $hScriptsTreeItem)
						EndIf

						$hScriptsTreeItemParent = _GUICtrlTreeView_GetParentHandle ($ScriptsTree, $hScriptsTreeItem)

						If Not Int($hScriptsTreeItemParent) Then
							ContinueLoop

						Elseif _GUICtrlTreeView_GetChecked ($ScriptsTree, $hScriptsTreeItem) Then
							$Folder = _GUICtrlTreeView_GetText ($ScriptsTree, $hScriptsTreeItemParent)
							$File = _GUICtrlTreeView_GetText ($ScriptsTree, $hScriptsTreeItem)
							$RunFullPath = @ScriptDir & "\"&$Folder&"\"&$File
							_RunFile($RunFullPath)

						Endif
					Next

				Case $MenuUpdateButton
					_Log("MenuUpdateButton")
					_DownloadGitSetup($GITAPIURL, @ScriptDir)
					If MsgBox($MB_YESNO, $Title, "Updated "&$DownloadUpdatedCount&" files ("&$DownloadErrors&" errors). The following files were updated:"&@CRLF&$DownloadUpdated&@CRLF&"Restart script?", 0, $GUIMain) = $IDYES Then
						_RunFile(@ScriptFullPath)
						Exit
					EndIf

				Case $MenuShowLoginScriptsButton
					_PopulateScripts($ScriptsTree, "AutoLogin")

				Case $MenuVisitGitButton
					_Log("Opening Browser...", True)
					$o_URL = ObjCreate("Shell.Application")
					$o_URL.Open($GITURL)

				Case $JoinButton
					Run("SystemPropertiesComputerName.exe")
					$hWindow = WinWait("System Properties")
					ControlClick ( $hWindow, "", "[CLASS:Button; INSTANCE:2]")

				Case $CreateLocalUserButton
					$sUser = GUICtrlRead($UsernameInput)
					$sPassword = GUICtrlRead($PasswordInput)
					$Admin = GUICtrlRead($AdminCheckBox)

					If $sUser <> "" and $sPassword <> "" Then
						$objSystem = ObjGet("WinNT://localhost")
						$objUser = $objSystem.Create("user", $sUser)
						$objUser.SetPassword ($sPassword)
						$objUser.SetInfo
						If Not @error And $Admin = $GUI_CHECKED Then
							$objGroup = ObjGet("WinNT://localhost/Administrators")
							$objGroup.Add("WinNT://"&$sUser)
						EndIf

						If Not IsObj( ObjGet("WinNT://./" & $sUser & ", user") ) Then
							MsgBox($MB_ICONWARNING, $Title, "Error creating user", 0, $GUIMain)
						Else
							$UserCreated = True
							_Log("User Created Successfully", True)
						EndIf

					EndIf

			EndSwitch
		WEnd

	Case Else
		_Log("Command unknown")

EndSwitch

Func _PopulateScripts($TreeID, $Folder)
	Local $FileArray = _FileListToArray(@ScriptDir & "\"&$Folder&"\", "*", $FLTA_FILES, True)

	If Not @error Then
		;Local $OptLoginListItems[$FileArray[0] + 1]
		_Log("OptLogin Files: " & $FileArray[0])
		Local $FolderTreeItem = GUICtrlCreateTreeViewItem($Folder, $TreeID)
		;GUICtrlSetState($FolderTreeItem,$GUI_DISABLE)
		GUICtrlSetState($FolderTreeItem, $GUI_CHECKED)

		For $i = 1 To $FileArray[0]
			_Log("Added: "&$FileArray[$i])
			$FileName = StringTrimLeft($FileArray[$i], StringInStr($FileArray[$i], "\", 0, -1))
			;$OptLoginListItems[$i] = GUICtrlCreateTreeViewItem($FileName, $FolderTreeItem)
			GUICtrlCreateTreeViewItem($FileName, $FolderTreeItem)
		Next

		GUICtrlSetState($FolderTreeItem, $GUI_EXPAND)
		GUICtrlSetState($FolderTreeItem, $GUI_UNCHECKED)

		Return $FileArray
	Else
		_Log("No files")
		Return 0
	EndIf
EndFunc

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

Func _RunFile($File, $Params = "")
	_Log("_RunFile " & $File)
	$Extension = StringTrimLeft($File, StringInStr($File, ".", 0, -1))
	Switch $Extension
		Case "au3"
			$RunLine = @AutoItExe & " /AutoIt3ExecuteScript """ & $File & """ " & $Params
			;Return ShellExecute(@AutoItExe, "/AutoIt3ExecuteScript """ & $File & """ " & $Params)
			Return Run($RunLine, "", @SW_SHOW, $STDERR_CHILD + $STDOUT_CHILD)

		Case "ps1"
			;$File = StringReplace($File, "$", "`$")
			$RunLine = @ComSpec & " /c " & "powershell.exe -ExecutionPolicy Unrestricted -File """ & $File & """ " & $Params
			_Log("$RunLine=" & $RunLine)
			Return Run($RunLine, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)

		Case "reg"
			$RunLine = @ComSpec & " /c " & "reg import """ & $File & """"

			Local $Data = FileRead($File)
			If StringInStr($Data, ";32") Then
				$RunLine = $RunLine & " /reg:32"
			ElseIf StringInStr($Data, ";64") Then
				$RunLine = $RunLine & " /reg:64"
			ElseIf @CPUArch = "X64" Then
				$RunLine = $RunLine & " /reg:32"
			Endif

			_Log("$RunLine=" & $RunLine)
			Return Run($RunLine, "", @SW_SHOW, $STDERR_CHILD + $STDOUT_CHILD)

		Case Else
			Return ShellExecute($File, $Params)

	EndSwitch
EndFunc   ;==>_RunFile


Func _DownloadGitSetup($sURL, $Destination)
	FileSetAttrib ($Destination, "-R", $FT_RECURSIVE)

	Global $DownloadErrors = 0
	Global $DownloadUpdated = ""
	Global $DownloadUpdatedCount = 0

	Return _DownloadGit($sURL, $Destination)

Endfunc
Func _DownloadGit($sURL, $Destination)
	_Log("_DownloadGit - " & $sURL)
	Local $bData = _WinHTTPRead($sURL)
	If @error Then
		_Log("  API http error: "&@error)
		$DownloadErrors = $DownloadErrors + 1
		Return SetError(1, @error, 0)
	EndIf

	Local $sData = BinaryToString($bData)
	Local $Object = json_decode($sData)

	Local $i = -1
	While 1
		$i += 1
		Local $Name = json_get($Object, '[' & $i & '].name')
		If @error Then
			;_Log("JSON Error")
			Exitloop
		endif

		$oPath = json_get($Object, '[' & $i & '].path')
		$oType = json_get($Object, '[' & $i & '].type')
		$oURL = json_get($Object, '[' & $i & '].url')
		$oSize = json_get($Object, '[' & $i & '].size')
		$oDownload_url = json_get($Object, '[' & $i & '].download_url')



		If $oType = "dir" Then
			;recurse
			_DownloadGit($oURL, $Destination)

		Else
			;download
			_Log("Downloading "&$oPath)

			$FullPath = $Destination&"\"&StringReplace($oPath, "/", "\")
			$FolderPath = StringLeft($FullPath, StringInStr($FullPath, "\", 0, -1))
			$FileName = StringTrimLeft($FullPath, StringInStr($FullPath,"\",0,-1))


			$InetData = _WinHTTPRead($oDownload_url)
			If @error Then
				_Log("  File download http error: "&@error)
				$DownloadErrors = $DownloadErrors + 1
				ContinueLoop
			Endif

			$DownloadSize = BinaryLen ($InetData)
			If @error Then
				_Log("  BinaryLen error: "&@error)
				$DownloadErrors = $DownloadErrors + 1
				ContinueLoop

			ElseIf $DownloadSize = $oSize Then
				_Log("  Download API size match ("&$DownloadSize&")")

				If FileExists($FullPath) Then
					$FileHash = _Crypt_HashFile ($FullPath, $CALG_MD5)
					$DataHash = _Crypt_HashData ($InetData, $CALG_MD5)
					If $FileHash = $DataHash Then
						_Log("  File unchanged, skipping ("&$FileHash&")")
						ContinueLoop
					Else
						_Log("  File changed, writing... ("&$FileHash&"/"&$DataHash&")")
						If FileDelete($FullPath) = 0 Then _Log("  Couldn't delete file")
					EndIf

				EndIf

				$hOutFile = FileOpen($FullPath, $FO_OVERWRITE + $FO_CREATEPATH)
				If NOT @error Then
					$FileWrite = FileWrite($hOutFile, $InetData)
					If Not @error Then
						_Log("  File write success")
						$DownloadUpdated = $DownloadUpdated & $FileName & @CRLF
						$DownloadUpdatedCount = $DownloadUpdatedCount + 1

					Else
						_Log("  File write error: "&@error)
						$DownloadErrors = $DownloadErrors + 1

					EndIf
					FileClose($hOutFile)

				Else
					_Log("  File open error: "&@error)
				Endif


			Else
				_Log("  Bad Size, Downloaded " & $DownloadSize & " But Expected " & $oSize)
				$DownloadErrors = $DownloadErrors + 1

			EndIf

		endif


	WEnd

EndFunc

Func _WinHTTPRead($sURL, $Agent = "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:15.0) Gecko/20100101 Firefox/15.0.1")
	_Log("_WinHTTPRead " & $sURL)
	; Open needed handles
	Local $hOpen = _WinHttpOpen($Agent)

	Local $iStart = StringInStr($sURL,"/",0,2)+1
	Local $Connect = StringMid($sURL, $iStart, StringInStr($sURL,"/",0,3) - $iStart)

	Local $hConnect = _WinHttpConnect($hOpen, $Connect)

	; Specify the reguest:
	Local $RequestURL = StringTrimLeft($sURL,StringInStr($sURL,"/",0,3))
	Local $hRequest = _WinHttpOpenRequest($hConnect, "GET", $RequestURL, Default, Default, Default, $WINHTTP_FLAG_SECURE + $WINHTTP_FLAG_ESCAPE_DISABLE + $WINHTTP_FLAG_BYPASS_PROXY_CACHE)

	_WinHttpAddRequestHeaders ($hRequest, "Cache-Control: no-cache")
	_WinHttpAddRequestHeaders ($hRequest, "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3")
	_WinHttpAddRequestHeaders ($hRequest, "content-type: application/json")

	Local $Token = IniRead("git.token","t","t","")
	If $Token <> "" Then
		_Log("Using Token")
		_WinHttpAddRequestHeaders ($hRequest, "Authorization: token "&$Token)
	EndIf

	; Send request
	_WinHttpSendRequest($hRequest)
	If @error Then
		_WinHttpCloseHandle($hRequest)
		_WinHttpCloseHandle($hConnect)
		_WinHttpCloseHandle($hOpen)
		_Log("Connection error (Send)")
		Return SetError(1, 0, 0)
	Endif

	; Wait for the response
	_WinHttpReceiveResponse($hRequest)
	If @error Then
		_WinHttpCloseHandle($hRequest)
		_WinHttpCloseHandle($hConnect)
		_WinHttpCloseHandle($hOpen)
		_Log("Connection error (Receive)")
		Return SetError(2, 0, 0)
	Endif

	Local $sHeader = _WinHttpQueryHeaders($hRequest) ; ...get full header

	Local $bData, $bChunk
	While 1
		$bChunk = _WinHttpReadData($hRequest, 2)
		If @error Then ExitLoop
		$bData = _WinHttpBinaryConcat($bData, $bChunk)
	WEnd

	; Clean
	_WinHttpCloseHandle($hRequest)
	_WinHttpCloseHandle($hConnect)
	_WinHttpCloseHandle($hOpen)

	Return $bData

EndFunc

Func _Log($Message, $Statusbar = "")
	Local $sTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "> " ; Generate Timestamp
	ConsoleWrite($sTime & $Message & @CRLF)
	If $Statusbar Then _GUICtrlStatusBar_SetText($StatusBar1, $Message)

	FileWrite(@ScriptFullPath&"_log.txt", $sTime & $Message & @CRLF)
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

EndFunc

Func _Exit()
	_Log("End script " & $CmdLineRaw)

EndFunc   ;==>_Exit
