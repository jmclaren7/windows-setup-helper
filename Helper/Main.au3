#RequireAdmin
#include "include\AutoItConstants.au3"
#include "include\ButtonConstants.au3"
#include "include\ComboConstants.au3"
#include "include\Crypt.au3"
#include "include\Date.au3"
#include "include\EditConstants.au3"
#include "include\File.au3"
#include "include\FileConstants.au3"
#include "include\GuiComboBox.au3"
#include "include\GuiConstantsEx.au3"
#include "include\GuiEdit.au3"
#include "include\GuiListBox.au3"
#include "include\GuiListView.au3"
#include "include\GuiTab.au3"
#include "include\GuiToolTip.au3"
#include "include\GuiTreeView.au3"
#include "include\GuiStatusBar.au3"
#include "include\Inet.au3"
#include "include\InetConstants.au3"
#include "include\ListViewConstants.au3"
#include "include\Process.au3"
#include "include\StaticConstants.au3"
#include "include\String.au3"
#include "include\TabConstants.au3"
#include "include\TreeViewConstants.au3"
#include "include\WindowsConstants.au3"
#include "include\WinAPI.au3"
#include "include\WinAPISys.au3"
#include "include\WinAPIFiles.au3"

; https://github.com/jmclaren7/autoit-scripts/blob/master/CommonFunctions.au3
#include "includeExt\CommonFunctions.au3"
#include "includeExt\Misc.au3"

; Register a function to run whenever the script exits
OnAutoItExitRegister("_Exit")

; Fix for issues working in a 64-bit only environment
_WinAPI_Wow64EnableWow64FsRedirection(False)

; AutoIT options
Opt("WinTitleMatchMode", -2)
Opt("TrayIconHide", 1)

; Make sure the working directory is the script directory
FileChangeDir(@ScriptDir)

; Miscellaneous global variables
Global $Date = StringTrimRight(FileGetTime(@ScriptFullPath, $FT_MODIFIED, $FT_STRING), 6)
Global $Version = "5.7"
Global $TitleShort = "Windows Setup Helper"
Global $Title = $TitleShort & " v" & $Version & " (" & $Date & ")"
Global $oCommError = ObjEvent("AutoIt.Error", "_CommError")
Global $DoubleClick = False
Global $SystemDrive = StringLeft(@SystemDir, 3)
Global $IsPE = StringInStr(@SystemDir, "X:")
Global $Debug = Not $IsPE
Global $LaunchFiles = StringSplit("main.au3,main.bat,main.exe,a.bat", ",") ; Used by _RunFile & _PopulateScripts
Global $MainConfig = "Config.ini"

; Default values used for automatic setup if missing from autounattend.xml
Global $DefaultComputerName = "WINDOWS-" & _RandomString(7, 7, "0123456789ABCDEFGHIJKLMNOPQRSTUV")
Global $DefaultWIMPath = "D:\sources\install.wim"
Global $DefaultAutounattend = @ScriptDir & "\autounattend.xml"
Global $SelectedAutounattendFile = $DefaultAutounattend
Global $DefaultEdition = "Windows 11 Pro"
Global $DefaultAdminPassword = "1234"
Global $DefaultLanguage = "en-US"

; Globals used by _Log function
Global $LogFullPath = StringReplace(@TempDir & "\Helper_" & @ScriptName, ".au3", ".log")
Global $LogTitle = $Title & " Log"
Global $LogFlushAlways = False
Global $LogLevel = 1
If $Debug Then $LogLevel = 3

; Some diagnostic information
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
_Log("@OSBuild=" & @OSBuild)
_Log("@ComputerName=" & @ComputerName)
If $IsPE Then _Log("PATH=" & EnvGet("PATH"))

; Access restriction
Local $AccessSecret = IniRead($MainConfig, "Access", "Secret", "")
Local $AccessSalt = IniRead($MainConfig, "Access", "Salt", "3b194da2")
Local $AccessPasswordHash = IniRead($MainConfig, "Access", "PasswordSHA256", "")
Local $AccessNetworking = IniRead($MainConfig, "Access", "PromptStage", "False")
Local $AccessPEAutoRun = IniRead($MainConfig, "Access", "PEAutoRun", "False")
Local $NetworkStarted = False
Local $PEAutoRunStarted = False

; Start PE networking before auth if configured
If $IsPE And $AccessNetworking = "True" Then
	Run(@ComSpec & " /c " & 'wpeinit.exe', @SystemDir, @SW_HIDE, $RUN_CREATE_NEW_CONSOLE)
	$NetworkStarted = True
EndIf

; Run automatic setup scripts before auth if configured
If $IsPE And $AccessPEAutoRun = "True" Then
	_RunFolder("PEAutoRun")
	$PEAutoRunStarted = True
EndIf

; Prompt for password
If $AccessPasswordHash <> "" Or $AccessSecret <> "" Then
	Local $AccessChallenge = _RandomString(8, 8, "1234567890ABCDEFGHJKMNPQRSTUVWXYZ")
	Local $AccessChallengeDisplay = StringLeft($AccessChallenge, 4) & "-" & StringRight($AccessChallenge, 4)
	_Log("$AccessChallengeDisplay= " & $AccessChallengeDisplay)

	Local $AccessCode = StringRight(_Crypt_HashData($AccessSecret & $AccessChallenge & $AccessSalt, $CALG_SHA_256), 8)
	$Code = StringLeft($AccessCode, 4) & "-" & StringRight($AccessCode, 4)


	Local $AccessResponseMessage = ""
	While 1
		Local $AccessInput = InputBox("Access Restricted - " & $TitleShort, "Enter the password to continue, canceling will restart the system, timeout in 10 minutes." & @CRLF & @CRLF & $AccessChallengeDisplay & $AccessResponseMessage, "", "*", 330, 170, Default, Default, 600, $_hLogWindow)
		If @error = 1 Then Exit

		Local $AccessCodeInput = StringStripWS($AccessInput, $STR_STRIPALL)
		$AccessCodeInput = StringReplace($AccessInput, "-", "")
		If $AccessSecret <> "" And $AccessCodeInput = $AccessCode Then ExitLoop

		Local $AccessPasswordHashInput = StringTrimLeft(_Crypt_HashData($AccessInput & $AccessSalt, $CALG_SHA_256), 2)
		_Log("$PasswordHashInput=" & $AccessPasswordHashInput)
		If $AccessPasswordHash <> "" And $AccessPasswordHashInput = $AccessPasswordHash Then ExitLoop

		$AccessResponseMessage = "     Response was invalid, try again."
		Sleep(10)
	WEnd
EndIf

; Start PE networking if it wasn't before auth
If $IsPE And Not $NetworkStarted Then Run(@ComSpec & " /c " & 'wpeinit.exe', @SystemDir, @SW_HIDE, $RUN_CREATE_NEW_CONSOLE)

; Run automatic setup scripts if it wasn't before auth
If $IsPE And Not $PEAutoRunStarted Then _RunFolder("PEAutoRun")

; Globals used by GUI
Global $GUIMain
Global $StatusBar1
Global $StatusbarTimer2
Global $GUIMainWidth = 753
Global $GUIMainHeight = 513

; Create main GUI
#Region ### START Koda GUI section ###
$GUIMain = GUICreate("Title", $GUIMainWidth, $GUIMainHeight, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_TABSTOP))
$FileMenu = GUICtrlCreateMenu("&File")
$AdvancedMenu = GUICtrlCreateMenu("&Advanced")
$AboutMenuItem = GUICtrlCreateMenuItem("A&bout", -1)
GUISetBkColor(0xF9F9F9)
$Group5 = GUICtrlCreateGroup("WinPE Tools", 12, 7, 354, 452)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
$PEScriptTreeView = GUICtrlCreateTreeView(24, 31, 330, 385)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
$PERunButton = GUICtrlCreateButton("Run", 242, 424, 107, 25)
GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Run the selected tool (you can also double click on the list item)")
$TaskMgrButton = GUICtrlCreateButton("", 28, 424, 29, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Task Manager")
$RegeditButton = GUICtrlCreateButton("", 68, 424, 29, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Registry Editor")
$NotepadButton = GUICtrlCreateButton("", 108, 424, 29, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Notepad")
$CMDButton = GUICtrlCreateButton("", 148, 424, 29, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Command Prompt")
$ShellButton = GUICtrlCreateButton("", 188, 424, 29, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "File Explorer (Explorer++)")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group6 = GUICtrlCreateGroup("First Logon Scripts", 384, 7, 354, 452)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
$NormalInstallButton = GUICtrlCreateButton("Normal Install", 400, 424, 131, 25)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Starts the installer with no modifications or automations")
$AutomatedInstallButton = GUICtrlCreateButton("Automated Install...", 560, 424, 163, 25, $BS_DEFPUSHBUTTON)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Prompts for setup options before starting")
$PEInstallTreeView = GUICtrlCreateTreeView(396, 31, 330, 385, BitOR($GUI_SS_DEFAULT_TREEVIEW,$TVS_CHECKBOXES))
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$StatusBar1 = _GUICtrlStatusBar_Create($GUIMain)
_GUICtrlStatusBar_SetSimple($StatusBar1)
_GUICtrlStatusBar_SetText($StatusBar1, "")
#EndRegion ### END Koda GUI section ###

; Button Icons
GUICtrlSetStyle($TaskMgrButton, $BS_ICON)
GUICtrlSetImage($TaskMgrButton, @WindowsDir & "\System32\taskmgr.exe", 1, 0)
GUICtrlSetStyle($RegeditButton, $BS_ICON)
GUICtrlSetImage($RegeditButton, @WindowsDir & "\regedit.exe", 1, 0)
GUICtrlSetStyle($NotepadButton, $BS_ICON)
GUICtrlSetImage($NotepadButton, @WindowsDir & "\System32\notepad.exe", 1, 0)
If FileExists(@WindowsDir & "\System32\WindowsPowerShell\v1.0\powershell.exe") Then
	GUICtrlSetStyle($CMDButton, $BS_ICON)
	GUICtrlSetImage($CMDButton, @WindowsDir & "\System32\WindowsPowerShell\v1.0\powershell.exe", 1, 0)
Else
	GUICtrlSetStyle($CMDButton, $BS_ICON)
	GUICtrlSetImage($CMDButton, @WindowsDir & "\System32\cmd.exe", 1, 0)
EndIf
If FileExists("Tools\.Explorer++.exe") Then
	GUICtrlSetStyle($ShellButton, $BS_ICON)
	GUICtrlSetImage($ShellButton, "Tools\.Explorer++.exe", 1, 0)
Else
	GUICtrlDelete($ShellButton)
EndIf

; File Menu Items
$MenuExitButton = GUICtrlCreateMenuItem("Exit", $FileMenu)

; Advanced Menu Items
$MenuShowConsole = GUICtrlCreateMenuItem("Show Log Window", $AdvancedMenu)
$MenuOpenLog = GUICtrlCreateMenuItem("Open Log File", $AdvancedMenu)
$MenuRunMain = GUICtrlCreateMenuItem("Run Main", $AdvancedMenu)
$MenuRelistScripts = GUICtrlCreateMenuItem("Relist Tools && Scripts", $AdvancedMenu)
$MenuListDebugTools = GUICtrlCreateMenuItem("List Debug && AutoRun Tools", $AdvancedMenu)

; GUI Post Creation Setup
WinSetTitle($GUIMain, "", $Title)

; Generate Script List
_PopulateScripts($PEInstallTreeView, "Logon*")
_PopulateScripts($PEScriptTreeView, "Tools*")

; Variables used in GUI loop
Local $hSetup
Local $RebootPrompt = False
Local $AutoInstallWait = False
Local $NormalInstallWait = False
Local $Reboot = False
Local $EdditionChoice = "Windows 11 Pro"
If @OSVersion = "WIN_10" Then $EdditionChoice = "Windows 10 Pro"

; Set GUI Icon
GUISetIcon($SystemDrive & "sources\setup.exe")

; Hide console windows
_Log("Hide console window")
WinSetState($LogTitle, "", @SW_HIDE)

; Show GUI
GUISetState(@SW_SHOW)

; Setup statusbar updates
Global $StatusBarToolTip = _GUIToolTip_Create(0)
_GUIToolTip_AddTool($StatusBarToolTip, 0, " ", $StatusBar1)
_GUIToolTip_SetMaxTipWidth($StatusBarToolTip, 400)
AdlibRegister("_StatusBarUpdate", 4000)
_StatusBarUpdate()

; Setup double click detection for $PEScriptTreeView
$HUser32DLL = DllOpen(@WindowsDir & "\System32\user32.dll")
Global $hPEScriptTreeView = GUICtrlGetHandle($PEScriptTreeView)
GUIRegisterMsg($WM_NOTIFY, "_WM_NOTIFY")

; Setup window resize detection for status bar resize
GUIRegisterMsg($WM_SIZE, "_WM_SIZE")
;WinMove($GUIMain, "", Default, Default, 900, 700) ; Size the window as desired here

_Log("Ready")

;GUI Loop
While 1
	$nMsgA = GUIGetMsg(1)
	$nMsg = $nMsgA[0]
	If $nMsg > 0 Then _Log("MSG $nMsg=" & $nMsg, 3)
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
			_Log("NormalInstallButton")
			If ProcessExists($hSetup) Then
				_Log("Setup is already running")
				MsgBox(0, "Error - " & $Title, "Setup is already running, please close it first")
				ContinueLoop
			EndIf
			$hSetup = _RunFile($SystemDrive & "sources\setup.exe", "/noreboot")
			$NormalInstallWait = True
			$AutoInstallWait = False

		Case $AutomatedInstallButton
			_Log("AutomatedInstallButton")
			If ProcessExists($hSetup) Then
				_Log("Setup is already running")
				MsgBox(0, "Error - " & $Title, "Setup is already running, please close it first")
				ContinueLoop
			EndIf

			GUISetState(@SW_DISABLE, $GUIMain)

			#Region ### START Koda GUI section ###
			$AutoInstallForm = GUICreate("Install Options", 559, 409, -1, -1, -1, -1, $GUIMain)
			GUISetBkColor(0xF9F9F9)
			$CancelButton = GUICtrlCreateButton("Cancel", 446, 373, 91, 25)
			$InstallButton = GUICtrlCreateButton("Install", 338, 373, 91, 25)
			$LocalizationGroup = GUICtrlCreateGroup("Localization", 16, 292, 528, 69)
			$TimezoneCombo = GUICtrlCreateCombo("", 92, 320, 233, 25, BitOR($GUI_SS_DEFAULT_COMBO,$CBS_SIMPLE))
			$Label7 = GUICtrlCreateLabel("Timezone", 32, 324, 50, 17)
			$LanguageInput = GUICtrlCreateInput("", 408, 320, 121, 21)
			$Label8 = GUICtrlCreateLabel("Language", 348, 324, 52, 17)
			GUICtrlCreateGroup("", -99, -99, 1, 1)
			$SourcesGroup = GUICtrlCreateGroup("SourcesGroup", 16, 7, 528, 133)
			$EditionCombo = GUICtrlCreateCombo("", 143, 101, 281, 25, BitOR($GUI_SS_DEFAULT_COMBO,$CBS_SIMPLE))
			$Label3 = GUICtrlCreateLabel("Edition", 98, 105, 36, 17)
			$WIMBrowseButton = GUICtrlCreateButton("Browse...", 427, 63, 75, 25)
			$WIMInput = GUICtrlCreateInput("", 143, 65, 281, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_READONLY))
			$Label1 = GUICtrlCreateLabel("Image (WIM/ESD)", 42, 69, 92, 17)
			$AutounattendBrowseButton = GUICtrlCreateButton("Browse...", 427, 24, 75, 25)
			$AutounattendInput = GUICtrlCreateInput("", 143, 26, 281, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_READONLY))
			$Label2 = GUICtrlCreateLabel("Autounattend.xml", 48, 30, 86, 17)
			GUICtrlCreateGroup("", -99, -99, 1, 1)
			$GeneralGroup = GUICtrlCreateGroup("General", 16, 143, 528, 145)
			$ComputerNameInput = GUICtrlCreateInput("", 116, 169, 133, 21)
			$Label4 = GUICtrlCreateLabel("Computer Name", 28, 173, 80, 17)
			$DiskList = GUICtrlCreateList("", 116, 205, 412, 45, BitOR($GUI_SS_DEFAULT_LIST,$LBS_NOINTEGRALHEIGHT), 0)
			$AdminPasswordInput = GUICtrlCreateInput("", 396, 169, 133, 21)
			$Label6 = GUICtrlCreateLabel("Administrator Password", 272, 173, 113, 17)
			$WindowsDiskCheckbox = GUICtrlCreateCheckbox("Let Windows setup prompt for disk selection instead", 116, 260, 273, 17)
			$Label5 = GUICtrlCreateLabel("Select Disk", 51, 204, 58, 17)
			GUICtrlCreateGroup("", -99, -99, 1, 1)
			$Bypass11Checkbox = GUICtrlCreateCheckbox("Bypass Windows 11 System Requirements Check", 16, 376, 265, 17)
			#EndRegion ### END Koda GUI section ###

			_GUICtrlComboBox_SetDroppedWidth($TimezoneCombo, 400)
			GUICtrlSetLimit($AdminPasswordInput, 15)
			GUICtrlSetState($Bypass11Checkbox, $GUI_CHECKED)

			; Set GUI Icon
			GUISetIcon($SystemDrive & "sources\setup.exe")

			; Add timezones
			$sTimezones = FileRead("IncludeExt\tz.txt")
			$sTimezones = StringReplace($sTimezones, @CRLF & "(", "|(")
			$sTimezones = StringReplace($sTimezones, " " & @CRLF, "^")
			Global $aTimezones = _ArrayFromString($sTimezones, "^", "|", True)
;~ 			Global $aTimezones = StringRegExp($sTimezones, "(.*)\r\n(.*)\r\n\r\n", 3) ; Returns a 1d array, alternating utc/windows tz name
			If Not @error And IsArray($aTimezones) Then
				$ComboString = _ArrayToString($aTimezones, " (", Default, Default, "|", 0, 1)
				$ComboString = StringReplace($ComboString, "|", ")|") & ")"
				GUICtrlSetData($TimezoneCombo, $ComboString)
			EndIf

			; Read the autounattend.xml file
			$sAutounattendData = FileRead($SelectedAutounattendFile)
			_UpdateXMLDependents($sAutounattendData)

			; Set Autounattend.xml path in GUI
			GUICtrlSetData($AutounattendInput, $SelectedAutounattendFile)

			; Add disks to GUI
			$aDiskInfo = _GetDisks()
			Local $aDriveListItems[1]
			For $i = UBound($aDiskInfo) - 1 To 0 Step -1
				$ListItem = "Disk " & $aDiskInfo[$i][0] & " (" & $aDiskInfo[$i][4] & " Partitions)" & "  " & $aDiskInfo[$i][3] & "  " & $aDiskInfo[$i][1]
				GUICtrlSetData($DiskList, $ListItem)
				If $aDiskInfo[$i][0] = 0 Then _GUICtrlListBox_SelectString($DiskList, $ListItem)
			Next

			GUISetState(@SW_SHOW, $AutoInstallForm)

			While 1
				$nAutoInstallFormMsg = GUIGetMsg()
				Switch $nAutoInstallFormMsg
					Case $GUI_EVENT_CLOSE, $CancelButton
						GUISetState(@SW_ENABLE, $GUIMain)
						GUIDelete($AutoInstallForm)
						ContinueLoop 2

					; WIM Browse Button
					Case $WIMBrowseButton
						$SaveWorkingDir = @WorkingDir
						$FileSelection = FileOpenDialog("Select Windows Install Image", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "Windows Images (*.wim;*.esd)|All (*.*)", $FD_FILEMUSTEXIST, "", $AutoInstallForm)
						If FileExists($FileSelection) Then
							GUICtrlSetData($WIMInput, $FileSelection)
							_UpdateWIMDependents()
						EndIf
						FileChangeDir($SaveWorkingDir)

					; Autounattend.xml Browse Button
					Case $AutounattendBrowseButton
						$SaveWorkingDir = @WorkingDir
						$FileSelection = FileOpenDialog("Select Windows Answer File", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "XML File (*.xml)|All (*.*)", $FD_FILEMUSTEXIST, "", $AutoInstallForm)
						If FileExists($FileSelection) Then
							$SelectedAutounattendFile = $FileSelection
							GUICtrlSetData($AutounattendInput, $SelectedAutounattendFile)
							; Read the autounattend.xml file
							$sAutounattendData = FileRead($SelectedAutounattendFile)
							_UpdateXMLDependents($sAutounattendData)
							_UpdateWIMDependents()
						EndIf
						FileChangeDir($SaveWorkingDir)

					; Disable/Enable Disk Selection
					Case $WindowsDiskCheckbox
						If GUICtrlRead($WindowsDiskCheckbox) = $GUI_UNCHECKED Then
							GUICtrlSetState($DiskList, $GUI_ENABLE)
						Else
							GUICtrlSetState($DiskList, $GUI_DISABLE)
						EndIf

					Case $InstallButton
						GUISetState(@SW_DISABLE, $AutoInstallForm)

						; Get the list of scripts that need to be copied later
						$aAutoLogonCopy = _RunTreeView($GUIMain, $PEInstallTreeView, True)
						For $b = 0 To UBound($aAutoLogonCopy) - 1
							_Log("TreeItem: " & $aAutoLogonCopy[$b])
						Next

						; ================ Start modifications to autounattend.xml ================
						; Read the answers file again in case the user had an error during modification and had to start over
						$sAutounattendData = FileRead($SelectedAutounattendFile)

						; WIM/ESD Path
						; <InstallFrom>*<Path>?</Path>*</InstallFrom>
						$WIMInputText = GUICtrlRead($WIMInput)
						$WIMInputText = StringReplace($WIMInputText, "\", "\\")
						_Log("$WIMInputText=" & $WIMInputText)
						$sAutounattendData = StringRegExpReplace($sAutounattendData, "(?si)(<InstallFrom>.*?<Path>).*?(</Path>.*?</InstallFrom>)", "${1}" & $WIMInputText & "${2}")
						_Log("StringRegExpReplace @error=" & @error)

						; Edition and key
						$EdditionChoice = GUICtrlRead($EditionCombo)
						$RegExReplacement = "${1}${2}" & $EdditionChoice & "${3}${4}"
						If $EdditionChoice = "Not Specified" Then $RegExReplacement = "${1}${4}"
						$sAutounattendData = StringRegExpReplace($sAutounattendData, "(?si)(<InstallFrom>.*?)(<Value>).*?(</Value>)(.*?</InstallFrom>)", $RegExReplacement)
						If StringInStr($EdditionChoice, "Home") Then
							$sAutounattendData = StringReplace($sAutounattendData, "<!--KeyHome", "")
							$sAutounattendData = StringReplace($sAutounattendData, "KeyHome-->", "")
						ElseIf StringInStr($EdditionChoice, "Pro") Then
							$sAutounattendData = StringReplace($sAutounattendData, "<!--KeyPro", "")
							$sAutounattendData = StringReplace($sAutounattendData, "KeyPro-->", "")
						EndIf

						; Computer name
						; <ComputerName>?</ComputerName>
						$ComputerName = GUICtrlRead($ComputerNameInput)
						_Log("$ComputerName=" & $ComputerName)
						$sAutounattendData = StringRegExpReplace($sAutounattendData, "(?si)(<ComputerName>).*?(</ComputerName>)", "${1}" & $ComputerName & "${2}")
						_Log("StringRegExpReplace @error=" & @error)

						; Administrator password
						; <AdministratorPassword>*<Value>?</Value>*</AdministratorPassword>
						$AdminPassword = GUICtrlRead($AdminPasswordInput)
						_Log("$AdminPassword=" & $AdminPassword)
						$sAutounattendData = StringRegExpReplace($sAutounattendData, "(?si)(<Password>.*?<Value>).*?(</Value>.*?</Password>)", "${1}" & $AdminPassword & "${2}")
						$sAutounattendData = StringRegExpReplace($sAutounattendData, "(?si)(<AdministratorPassword>.*<Value>).*?(</Value>.*</AdministratorPassword>)", "${1}" & $AdminPassword & "${2}")
						_Log("StringRegExpReplace @error=" & @error)

						; Disk/format options
						If GUICtrlRead($WindowsDiskCheckbox) = $GUI_UNCHECKED Then
							If EnvGet("firmware_type") = "Legacy" Then
								$sAutounattendData = StringReplace($sAutounattendData, "<!--FormatBIOS", "")
								$sAutounattendData = StringReplace($sAutounattendData, "FormatBIOS-->", "")
							Else
								$sAutounattendData = StringReplace($sAutounattendData, "<!--FormatUEFI", "")
								$sAutounattendData = StringReplace($sAutounattendData, "FormatUEFI-->", "")
							EndIf

							; Set the target disk
							$DiskListIndex = _GUICtrlListBox_GetCurSel($DiskList)
							$DiskListText = _GUICtrlListBox_GetText($DiskList, $DiskListIndex)
							$aTargetDisk = StringRegExp($DiskListText, '(?i)Disk (\d{1}) ', $STR_REGEXPARRAYMATCH)
							If @error Then
								MsgBox(0, "Select Disk - " & $TitleShort, "Error selecting disk: " & @error)
								GUISetState(@SW_ENABLE, $AutoInstallForm)
								ContinueLoop
							EndIf
							$sAutounattendData = StringRegExpReplace($sAutounattendData, "(?si)(<DiskID>).*?(</DiskID>)", "${1}" & $aTargetDisk[0] & "${2}")
						EndIf

						; Timezone
						; <TimeZone>?</TimeZone>
						$sTimezoneText = GUICtrlRead($TimezoneCombo)
						_Log("$sTimezoneText=" & $sTimezoneText)
						$aTimezoneText = StringRegExp($sTimezoneText, "\(([^)(]*(?:\((?:[^)(]+|\([^)(]*\))*\)[^)(]*)*)\)(?!.*\()", 1)
						If Not @error Then
							_Log("$aTimezoneText[0]=" & $aTimezoneText[0])
							$sAutounattendData = StringRegExpReplace($sAutounattendData, "(?si)(<TimeZone>).*?(</TimeZone>)", "${1}" & $aTimezoneText[0] & "${2}")
							_Log("StringRegExpReplace @error=" & @error)
						Else
							_Log("$aTimezoneText @error=" & @error)
						EndIf

						; <SystemLocale>en-US</SystemLocale>
						; <UILanguage>en-US</UILanguage>
						; <UserLocale>en-US</UserLocale>
						$LanguageText = GUICtrlRead($LanguageInput)
						_Log("$LanguageText=" & $LanguageText)
						$sAutounattendData = StringRegExpReplace($sAutounattendData, "(?si)(<SystemLocale>).*?(</SystemLocale>)", "${1}" & $LanguageText & "${2}")
						$sAutounattendData = StringRegExpReplace($sAutounattendData, "(?si)(<UILanguage>).*?(</UILanguage>)", "${1}" & $LanguageText & "${2}")
						$sAutounattendData = StringRegExpReplace($sAutounattendData, "(?si)(<UserLocale>).*?(</UserLocale>)", "${1}" & $LanguageText & "${2}")
						_Log("StringRegExpReplace @error=" & @error)

						; Windows 11 requirements bypass
						If GUICtrlRead($Bypass11Checkbox) = $GUI_CHECKED And $IsPE And @OSVersion = "WIN_11" Then _Win11Bypass()

						; (Legacy) Replace instances of Windows 11 if running a Windows 10 ISO
						If @OSVersion = "WIN_10" Then $sAutounattendData = StringReplace($sAutounattendData, "Windows 11", "Windows 10")
						; ================ End modifications to autounattend.xml ================


						If $IsPE Then
							; Save modifications to autounattend.xml in new location
							$AutounattendPath = @TempDir & "\autounattend.xml"
							_Log("$AutounattendPath=" & $AutounattendPath)
							$hAutounattend = FileOpen($AutounattendPath, $FO_OVERWRITE)
							FileWrite($hAutounattend, $sAutounattendData)
							_Log("FileWrite @error=" & @error)
							FileClose($hAutounattend)

							$hSetup = _RunFile($SystemDrive & "sources\setup.exe", "/noreboot /unattend:" & $AutounattendPath)

							$NormalInstallWait = False
							$AutoInstallWait = True
						Else
							_Log($sAutounattendData)
						EndIf

						GUISetState(@SW_ENABLE, $GUIMain)
						GUIDelete($AutoInstallForm)
						ContinueLoop 2
				EndSwitch
			WEnd


		Case $TaskMgrButton
			_RunFile("taskmgr.exe")

		Case $RegeditButton
			_RunFile("regedit.exe")

		Case $NotepadButton
			_RunFile("notepad.exe")

		Case $CMDButton
			If FileExists(@WindowsDir & "\System32\WindowsPowerShell\v1.0\powershell.exe") Then
				_RunFile("powershell.exe")
			Else
				_RunFile("cmd.exe")
			EndIf

		Case $ShellButton
			_RunFile("Tools\.Explorer++.exe")

		Case $AboutMenuItem
			_Log($AboutMenuItem)
			_MsgBox($MB_ICONINFORMATION, "About - " & $TitleShort, "Windows Setup Helper" & @CRLF & "John Mclaren" & @CRLF & "https://github.com/jmclaren7/windows-setup-helper" & @CRLF & @CRLF & "Included 3rd party software is subject to its respective licensing")

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

	; Wait for auotmatic install to finish
	If $AutoInstallWait And Not ProcessExists($hSetup) Then
		_Log("Copy AutoLogon Files")
		For $i = 65 To 90 ; 65=A 90=Z
			$Drive = Chr($i) & ":"
			$TestFile = $Drive & "\Windows\System32\Config\SYSTEM"
			$Target = $Drive & "\Temp\Helper"

			; If the Windows install is less than 10 minutes old it must be the target
			If FileExists($TestFile) And _FileModifiedAge($TestFile) < 600000 Then
				_Log("Test file found on drive " & $Drive)

				; Add autorun script that's executed on first logon
				_ArrayAdd($aAutoLogonCopy, @ScriptDir & "\Logon\.Autorun.ps1")

				; Add log file for diagnostics
				_ArrayAdd($aAutoLogonCopy, $LogFullPath)

				; Copy files to new Windows installation
				For $iFile = 0 To UBound($aAutoLogonCopy) - 1
					$ThisFile = $aAutoLogonCopy[$iFile]
					If StringInStr(FileGetAttrib($ThisFile), "D") > 0 Then
						$DirName = StringTrimLeft($ThisFile, StringInStr($ThisFile, "\", 0, -1))
						$Return = DirCopy($ThisFile, $Target & "\" & $DirName, 1)
					Else
						$Return = FileCopy($ThisFile, $Target & "\", 1 + 8)
					EndIf
					_Log("Copy: " & $ThisFile & " (" & $Return & ")")
				Next

				ExitLoop

			EndIf
		Next
		If $i = 90 Then
			_Log("Could not find windows install")
			ContinueLoop
		EndIf

		$RebootPrompt = True ; Will trigger a prompt that will reboot on timeout
		$AutoInstallWait = False
	EndIf

	; Wait for normal install to finish
	If $NormalInstallWait And Not ProcessExists($hSetup) Then
		_Log("Normal install finished")
		$RebootPrompt = True  ; Will trigger a prompt that will reboot on timeout
		$NormalInstallWait = False
	EndIf

	; Reboot
	If $RebootPrompt Then
		_Log("Reboot")
		Beep(500, 1000)
		$Return = MsgBox(1 + 48 + 262144, $Title, "Rebooting in 15 seconds", 15)
		If $Return = $IDTIMEOUT Or $Return = $IDOK Then Exit
		$RebootPrompt = False
	EndIf

	; Window minimum size
	$GUIMainPos = WinGetPos($GUIMain)
	If $GUIMainPos[2] < $GUIMainWidth Then WinMove($GUIMain, "", Default, Default, $GUIMainWidth, Default)
	If $GUIMainPos[3] < $GUIMainHeight Then WinMove($GUIMain, "", Default, Default, Default, $GUIMainHeight)

	Sleep(10)
WEnd


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

; Proccess NOTIFY mesages to handle window resize for status bar
Func _WM_SIZE($hWnd, $iMsg, $iwParam, $ilParam)
	_GUICtrlStatusBar_Resize($StatusBar1)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_SIZE

; Update installer GUI items based on WIM contents
Func _UpdateWIMDependents()
	Global $WIMInput

	; Get editions from wim/esd
	Local $sWIMPath = GUICtrlRead($WIMInput)
	Local $sCommand = "dism /Get-ImageInfo /ImageFile:""" & $sWIMPath & """"
	Local $sReturn = _RunWait(@ComSpec & " /c " & $sCommand)
	Local $aEditions = StringRegExp($sReturn, "Name : (.*?)\R", $STR_REGEXPARRAYGLOBALMATCH)
	If Not @error Then
		GUICtrlSetData($EditionCombo, "|" & _ArrayToString($aEditions) & "|Not Specified")
	Else
		GUICtrlSetData($EditionCombo, "")
	EndIf

	; Preselect edition
	; <InstallFrom>*<Value>?</Value>*</InstallFrom>
	Local $aMatch = StringRegExp($sAutounattendData, '(?si)<InstallFrom>.*<Value>(.*?)</Value>.*</InstallFrom>', $STR_REGEXPARRAYMATCH)
	If Not @error And $aMatch[0] <> "" Then
		_GUICtrlComboBox_SelectString($EditionCombo, $aMatch[0])
	Else
		If _GUICtrlComboBox_SelectString($EditionCombo, $DefaultEdition) = -1 Then
			_GUICtrlComboBox_SelectString($EditionCombo, "Not Specified")
		EndIf
	EndIf

EndFunc   ;==>_UpdateWIMDependents

; Update installer GUI items based on XML contents
Func _UpdateXMLDependents($sXML)
	Local $aMatch

	; WIM path
	; <InstallFrom>*<Path>?</Path>*</InstallFrom>
	$aMatch = StringRegExp($sAutounattendData, '(?si)<InstallFrom>.*<Path>(.*?)</Path>.*</InstallFrom>', $STR_REGEXPARRAYMATCH)
	If Not @error And $aMatch[0] <> "" Then
		$aMatch[0] = StringStripWS($aMatch[0], 1 + 2)
		GUICtrlSetData($WIMInput, $aMatch[0])
	Else
		Local $aDrivesLetters = DriveGetDrive($DT_ALL)
		For $i = 1 To $aDrivesLetters[0]
			$aDrivesLetters[$i] = StringUpper($aDrivesLetters[$i])
			_Log("  Drive: " & $aDrivesLetters[$i])
			Local $TestWIMPath = $aDrivesLetters[$i] & "\sources\install.wim"
			Local $TestESDPath = $aDrivesLetters[$i] & "\sources\install.esd"

			If FileExists($TestWIMPath) Then
				GUICtrlSetData($WIMInput, $TestWIMPath)
				ExitLoop
			ElseIf FileExists($TestESDPath) Then
				GUICtrlSetData($WIMInput, $TestESDPath)
				ExitLoop
			EndIf
		Next
	EndIf
	_UpdateWIMDependents()

	; Computer name
	; <ComputerName>?</ComputerName>
	$aMatch = StringRegExp($sAutounattendData, '(?i)<ComputerName>(.*?)</ComputerName>', $STR_REGEXPARRAYMATCH)
	If Not @error And $aMatch[0] <> "" And $aMatch[0] <> "*" Then
		GUICtrlSetData($ComputerNameInput, $aMatch[0])
	Else
		GUICtrlSetData($ComputerNameInput, $DefaultComputerName)
	EndIf

	; Administrator password
	; <AdministratorPassword>*<Value>?</Value>*</AdministratorPassword>
	$aMatch = StringRegExp($sAutounattendData, '(?si)<AdministratorPassword>.*<Value>(.*?)</Value>.*</AdministratorPassword>', $STR_REGEXPARRAYMATCH)
	If Not @error And $aMatch[0] <> "" Then
		GUICtrlSetData($AdminPasswordInput, $aMatch[0])
	Else
		GUICtrlSetData($AdminPasswordInput, $DefaultAdminPassword)
	EndIf

	; Select timezone
	; <TimeZone>?</TimeZone>
	$aMatch = StringRegExp($sAutounattendData, '(?i)<TimeZone>(.*?)</TimeZone>', $STR_REGEXPARRAYMATCH)
	If Not @error Then
		$SearchIndex = _ArraySearch($aTimezones, $aMatch[0], Default, Default, 0, 0, 1, 1)
		_GUICtrlComboBox_SetCurSel($TimezoneCombo, $SearchIndex)
	EndIf

	; Get language
	; <SystemLocale>?</SystemLocale>
	$aMatch = StringRegExp($sAutounattendData, '(?i)<SystemLocale>(.*?)</SystemLocale>', $STR_REGEXPARRAYMATCH)
	If Not @error And $aMatch[0] <> "" Then
		GUICtrlSetData($LanguageInput, $aMatch[0])
	Else
		GUICtrlSetData($LanguageInput, $DefaultLanguage)
	EndIf
EndFunc   ;==>_UpdateXMLDependents

; Get indexes from WIM or ESD
Func _GetImageNames($sPath)
	Local $sCommand = "dism /Get-ImageInfo /ImageFile:""" & $sPath & """"
	Local $sReturn = _RunWait(@ComSpec & " /c " & $sCommand)

	Local $aImageNames = StringRegExp($sReturn, "Name : (.*?)\R", $STR_REGEXPARRAYGLOBALMATCH)

	If Not @error And IsArray($aImageNames) Then
		Return $aImageNames
	Else
		Return SetError(1, 0, 0)
	EndIf

EndFunc   ;==>_GetImageNames

; Calculate the full path of an item from the GUI tree view
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

; Check to see if the same folder exists on other drives or if other folder start with the same name
Func _GetSimilarPaths($Folder, $Path = Default)
	If $Path = Default Then $Path = @ScriptDir
	_Log("_GetSimilarPaths: " & $Folder & " - " & $Path)

	; Get folders from the script path
	Local $aFolders = _FileListToArrayRec($Path, $Folder & "*", $FLTAR_FOLDERS, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)
	If @error Then _Log("  @error=" & @error & " @extended=" & @extended)

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

; List files from a folder and add them to a GUI tree view
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
	_ArraySort($aFiles, 0, 1)
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

			; Skip files starting with "." (treat them as hidden)
			If StringInStr($aFiles[$i], "\.") Then ContinueLoop

			; Folders that contain specific files can be added as a script
			If StringInStr(FileGetAttrib($aFiles[$i]), "D") Then
				Local $FileExists = 0
				For $b = 1 To $LaunchFiles[0]
					$FileExists += FileExists($aFiles[$i] & "\" & $LaunchFiles[$b])
				Next

				; If this folder does not have a reqiured file do not add it to the scripts
				If Not $FileExists Then ContinueLoop
			EndIf
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

	_Log("  End _PopulateScripts for: " & $Folder)

	Return $aFiles

EndFunc   ;==>_PopulateScripts

; Run the slected items from a tree view or return a list of the selected items
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
				_Log("  Checked: $RunFullPath=" & $RunFullPath)
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
Func _RunFolder($Folder)
	_Log("_RunFolder " & $Folder)

	Local $aPaths = _GetSimilarPaths($Folder)
	For $x = 1 To $aPaths[0]
		_Log("  $Paths[" & $x & "]=" & $aPaths[$x])
		Local $aFiles = _FileListToArray($aPaths[$x], "*", $FLTA_FILESFOLDERS, True) ;switched from $FLTA_FILES for allowing main.au3 in folder
		If Not @error Then
			_Log("  Files: " & $aFiles[0])
			For $i = 1 To $aFiles[0]
				If StringInStr($aFiles[$i], "\.") Then ContinueLoop
				_Log($aFiles[$i])
				_RunFile($aFiles[$i])
			Next
			Return $aFiles[0]
		Else
			_Log("  No files")
		EndIf

	Next

	Return $aPaths
EndFunc   ;==>_RunFolder

; Runs a file, automaticly handling file type and sub folders
Func _RunFile($File, $Params = "", $WorkingDir = "")
	_Log("_RunFile " & $File & " " & $Params)

	If StringInStr(FileGetAttrib($File), "D") Then
		; Folders that contain specific files can be executed
		Local $FileExists = 0
		For $b = 1 To $LaunchFiles[0]
			If FileExists($File & "\" & $LaunchFiles[$b]) Then
				$FileExists = 1
				$File = $File & "\" & $LaunchFiles[$b]
				ExitLoop
			EndIf
		Next
		; Return error because this is a directory without a valid file to run
		If Not $FileExists Then Return SetError(2, 0, 0)
	EndIf

	$Extension = StringTrimLeft($File, StringInStr($File, ".", 0, -1))
	Switch $Extension
		Case "au3", "a3x"
			_Log("  au3")
			$RunLine = @AutoItExe & " /AutoIt3ExecuteScript """ & $File & """ " & $Params
			_Log("  $RunLine=" & $RunLine)
			Return Run($RunLine, $WorkingDir, @SW_SHOW, $STDIO_INHERIT_PARENT)

		Case "ps1"
			_Log("  ps1")
			$RunLine = @ComSpec & " /c " & "powershell.exe -ExecutionPolicy Bypass -File """ & $File & """ " & $Params
			_Log("  $RunLine=" & $RunLine)
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

			_Log("  $RunLine=" & $RunLine)
			Return Run($RunLine, $WorkingDir, @SW_SHOW, $STDERR_CHILD + $STDOUT_CHILD)

		Case Else
			_Log("  Other file type")
			Return ShellExecute($File, $Params, $WorkingDir)

	EndSwitch

EndFunc   ;==>_RunFile

; Update the status bar with new infromation
Func _StatusBarUpdate()
	Local $StatusbarText, $StatusbarToolTipText
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

	; Get information for the first network adapter WMI provies (usually correct)
	$Win32_NetworkAdapterConfiguration = _WMI("SELECT IPAddress,DefaultIPGateway,DNSServerSearchOrder FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True and DHCPEnabled = True")
	If Not @error And IsObj($Win32_NetworkAdapterConfiguration) Then
		; If an IP hasn't been aquired it wont be an array
		If IsArray($Win32_NetworkAdapterConfiguration.IPAddress) Then $StatusbarText &= $Delimiter & $Win32_NetworkAdapterConfiguration.IPAddress[0]
		If IsArray($Win32_NetworkAdapterConfiguration.DefaultIPGateway) Then $StatusbarToolTipText &= "Gateway: " & $Win32_NetworkAdapterConfiguration.DefaultIPGateway[0] & "  "
		If IsArray($Win32_NetworkAdapterConfiguration.DNSServerSearchOrder) Then $StatusbarToolTipText &= "DNS: " & $Win32_NetworkAdapterConfiguration.DNSServerSearchOrder[0]
	EndIf
	$StatusbarToolTipText &= @CR & "Other IPs: " & @IPAddress1 & ", " & @IPAddress2 & ", " & @IPAddress3 & ", " & @IPAddress4

	; Get Hostname
	$StatusbarText &= "/" & @ComputerName

	; Get memory information
	If Not IsDeclared("_MemStats") Then
		$MemStats = MemGetStats()
		Static Local $_MemStatsText = $Delimiter & Round($MemStats[1] / 1024 / 1024) & "GB"
	EndIf
	If StringLen($_MemStatsText) > 2 Then $StatusbarText &= $_MemStatsText

	; Get CPU information
	If Not IsDeclared("_CPUStats") Then
		$Win32_Processor = _WMI("SELECT NumberOfCores,NumberOfLogicalProcessors FROM Win32_Processor")
		Static Local $_CPUStatsText = "/" & $Win32_Processor.NumberOfCores & "C/" & $Win32_Processor.NumberOfLogicalProcessors & "T"
	EndIf
	If StringLen($_CPUStatsText) > 4 Then $StatusbarText &= $_CPUStatsText

	; Get firmware type
	$Firmware_Type = EnvGet("firmware_type")
	If $Firmware_Type = "Legacy" Then $Firmware_Type = "BIOS"
	$StatusbarText &= $Delimiter & $Firmware_Type

	; Get motherboard bios information
	$Win32_BIOS = _WMI("SELECT SerialNumber,SMBIOSBIOSVersion,ReleaseDate FROM Win32_BIOS")
	If Not @error Then
		$StatusbarText &= "/" & StringLeft($Win32_BIOS.ReleaseDate, 8)
		$StatusbarToolTipText &= @CR & "Firmware: " & StringLeft($Win32_BIOS.SMBIOSBIOSVersion, 20) & " Date: " & StringLeft($Win32_BIOS.ReleaseDate, 8)
		If $Win32_BIOS.SerialNumber <> "" And $Win32_BIOS.SerialNumber <> "System Serial Number" Then $StatusbarText &= $Delimiter & StringLeft($Win32_BIOS.SerialNumber, 10)
	EndIf

	; Get time
	$StatusbarText &= $Delimiter & Int(@MON) & "/" & Int(@MDAY) & "/" & StringRight(@YEAR, 2) & " " & @HOUR & ":" & @MIN

	; Get additional statusbar and tool tip text
	$HelperStatusFiles = _FileListToArray(@TempDir, "Helper_Status_*.txt", $FLTA_FILES, True)
	For $i = 1 To UBound($HelperStatusFiles) - 1
		If _FileModifiedAge($HelperStatusFiles[$i]) < 10 * 1000 Then
			$FileText = FileReadLine($HelperStatusFiles[$i], 1)
			_Log("$FileText=" & $FileText, 3)
			If Not @error Then $StatusbarText &= $Delimiter & $FileText

			$FileText = FileReadLine($HelperStatusFiles[$i], 2)
			_Log("$FileText=" & $FileText, 3)
			If Not @error Then $StatusbarToolTipText &= @CRLF & $FileText

		Else
			FileDelete($HelperStatusFiles[$i])
		EndIf

	Next

	; Update statusbar if the text changed
	If _GUICtrlStatusBar_GetText($StatusBar1, 0) <> $StatusbarText Then
		_GUICtrlStatusBar_SetText($StatusBar1, $StatusbarText)
		_Log("Statusbar Updated", 3)
	EndIf

	; Update statusbar tool tip if the text changed
	If _GUIToolTip_GetText($StatusBarToolTip, 0, $StatusBar1) <> $StatusbarToolTipText Then
		_GUIToolTip_UpdateTipText($StatusBarToolTip, 0, $StatusBar1, $StatusbarToolTipText)
		_Log("Statusbar Tooltip Updated", 3)
	EndIf

	If $Debug Then
		$StatusbarTimer2 = TimerInit()
		_Log("_StatusBarUpdate Timer: " & Round(TimerDiff($StatusBarTimer1)) & "ms " & Round($StatusBarTimer2Value) & "ms")
	EndIf

	Return
EndFunc   ;==>_StatusBarUpdate

; Custom error handling, probably only good for running compiled
Func _CommError()
	Local $HexNumber
	Local $strMsg

	$HexNumber = Hex($oCommError.Number, 8)
	$strMsg = "Error: " & $HexNumber
	$strMsg &= "  Desc: " & $oCommError.WinDescription
	$strMsg &= "  Line: " & $oCommError.ScriptLine

	_Log($strMsg)

EndFunc   ;==>_CommError

; Custom exit function, called whenever a normal exit operation takes place
Func _Exit()
	_Log("End script " & $CmdLineRaw)

EndFunc   ;==>_Exit
