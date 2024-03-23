#include "include\AutoItConstants.au3"
#include "include\ButtonConstants.au3"
#include "include\ComboConstants.au3"
#include "include\Crypt.au3"
#include "include\Date.au3"
#include "include\EditConstants.au3"
#include "include\File.au3"
#include "include\FileConstants.au3"
#include "include\GuiConstantsEx.au3"
#include "include\GuiEdit.au3"
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
Global $Version = "5.3"
Global $Title = "Windows Setup Helper v" & $Version & " (" & $Date & ")"
Global $oCommError = ObjEvent("AutoIt.Error", "_CommError")
Global $DoubleClick = False
Global $SystemDrive = StringLeft(@SystemDir, 3)
Global $IsPE = StringInStr(@SystemDir, "X:")
Global $Debug = Not $IsPE
Global $FolderExecFiles = StringSplit("main.au3,main.bat,a.bat",",") ; Used by _RunFile & _PopulateScripts

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
_Log("PATH=" & EnvGet("PATH"))

; Run automatic setup scripts
If $IsPE Then _RunMulti("PEAutoRun")

; Globals used by GUI
Global $GUIMain
Global $StatusBar1
Global $StatusbarTimer2

; Create main GUI
#Region ### START Koda GUI section ###
$GUIMain = GUICreate("Title", 753, 513, -1, -1, BitOR($GUI_SS_DEFAULT_GUI,$WS_MAXIMIZEBOX,$WS_SIZEBOX,$WS_THICKFRAME,$WS_TABSTOP))
$FileMenu = GUICtrlCreateMenu("&File")
$AdvancedMenu = GUICtrlCreateMenu("&Advanced")
GUISetBkColor(0xF9F9F9)
$Group5 = GUICtrlCreateGroup("WinPE Tools", 12, 7, 354, 452)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
$PEScriptTreeView = GUICtrlCreateTreeView(24, 31, 330, 385)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
$PERunButton = GUICtrlCreateButton("Run", 237, 424, 107, 25)
GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
$TaskMgrButton = GUICtrlCreateButton("", 28, 424, 29, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Task Manager")
$RegeditButton = GUICtrlCreateButton("", 68, 424, 29, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Registry Editor")
$NotepadButton = GUICtrlCreateButton("", 108, 424, 29, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Notepad")
$CMDButton = GUICtrlCreateButton("", 149, 424, 29, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Command Prompt")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group6 = GUICtrlCreateGroup("First Logon Scripts", 384, 7, 354, 452)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
$NormalInstallButton = GUICtrlCreateButton("Normal Install", 400, 424, 131, 25)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
$AutomatedInstallButton = GUICtrlCreateButton("Automated Install", 560, 424, 131, 25, $BS_DEFPUSHBUTTON)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
$Label4 = GUICtrlCreateLabel("Computer Name", 471, 396, 80, 17)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
$PEInstallTreeView = GUICtrlCreateTreeView(396, 31, 330, 350, BitOR($GUI_SS_DEFAULT_TREEVIEW,$TVS_CHECKBOXES))
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
$PEComputerNameInput = GUICtrlCreateInput("", 553, 392, 169, 21)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
$FormatButton = GUICtrlCreateButton("", 698, 424, 29, 25)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Delete and install to disk 0")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$StatusBar1 = _GUICtrlStatusBar_Create($GUIMain)
_GUICtrlStatusBar_SetSimple($StatusBar1)
_GUICtrlStatusBar_SetText($StatusBar1, "")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

; Button Icons
GUICtrlSetStyle($TaskMgrButton, $BS_ICON)
GUICtrlSetImage($TaskMgrButton, @WindowsDir & "\System32\taskmgr.exe", 1, 0)
GUICtrlSetStyle($RegeditButton, $BS_ICON)
GUICtrlSetImage($RegeditButton, @WindowsDir & "\regedit.exe", 1, 0)
GUICtrlSetStyle($NotepadButton, $BS_ICON)
GUICtrlSetImage($NotepadButton, @WindowsDir & "\System32\notepad.exe", 1, 0)
GUICtrlSetStyle($CMDButton, $BS_ICON)
GUICtrlSetImage($CMDButton, @WindowsDir & "\System32\cmd.exe", 1, 0)
GUICtrlSetStyle($FormatButton, $BS_ICON)
GUICtrlSetImage($FormatButton, @WindowsDir & "\System32\shell32.dll", 240, 0)

; File Menu Items
$MenuExitButton = GUICtrlCreateMenuItem("Exit", $FileMenu)

; Advanced Menu Items
$MenuShowConsole = GUICtrlCreateMenuItem("Show Log Window", $AdvancedMenu)
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

; Variables used in GUI loop
Local $hSetup
Local $RebootPrompt = False
Local $CopyAutoLogonFiles = False
Local $Reboot = False

; Start PE networking
If $IsPE Then Run(@ComSpec & " /c " & 'wpeinit.exe', @SystemDir, @SW_HIDE, $RUN_CREATE_NEW_CONSOLE)

; Set GUI Icon
GUISetIcon($SystemDrive & "sources\setup.exe")

; Hide console windows
_Log("Hide console window")
WinSetState($LogTitle, "", @SW_HIDE)

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

		Case $WM_SIZE


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
			$hSetup = _RunFile($SystemDrive & "sources\setup.exe", "/noreboot")
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

			$hSetup = _RunFile($SystemDrive & "sources\setup.exe", "/noreboot /unattend:" & $AutounattendPath)

			$CopyAutoLogonFiles = True

		Case $TaskMgrButton
			_RunFile("taskmgr.exe")

		Case $RegeditButton
			_RunFile("regedit.exe")

		Case $NotepadButton
			_RunFile("notepad.exe")

		Case $CMDButton
			_RunFile("cmd.exe")

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
		For $i = 65 To 90 ; 65=A 90=Z
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
		If $i = 90 Then
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
EndFunc   ;==>MY_WM_SIZE

; Calculate the full path of a item from the GUI tree view
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
				For $b = 1 To $FolderExecFiles[0]
					$FileExists += FileExists($aFiles[$i] & "\" & $FolderExecFiles[$b])
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

	_Log("End _PopulateScripts for: " & $Folder)

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

; Run all the files in a folder
Func _RunFolder($Path)
	_Log("_RunFolder " & $Path)
	$FileArray = _FileListToArray($Path, "*", $FLTA_FILESFOLDERS, True) ;switched from $FLTA_FILES for allowing main.au3 in folder
	If Not @error Then
		_Log("Files: " & $FileArray[0])
		For $i = 1 To $FileArray[0]
			If StringInStr($FileArray[$i], "\.") Then ContinueLoop
			;We probably don't care if the folder is valid to run at this point, deal with it in the _RunFile function
			;If StringInStr(FileGetAttrib($FileArray[$i]), "D") And Not FileExists($FileArray[$i] & "\main.au3") Then ContinueLoop
			_Log($FileArray[$i])
			_RunFile($FileArray[$i])
		Next
		Return $FileArray[0]
	Else
		_Log("No files")
	EndIf

EndFunc   ;==>_RunFolder

; Runs a file, automaticly handling file type and sub folders
Func _RunFile($File, $Params = "", $WorkingDir = "")
	_Log("_RunFile " & $File & " " & $Params)

	If StringInStr(FileGetAttrib($File), "D") Then
		; Folders that contain specific files can be executed
		Local $FileExists = 0
		For $b = 1 To $FolderExecFiles[0]
			If FileExists($File & "\" & $FolderExecFiles[$b]) Then
				$FileExists = 1
				$File = $File & "\" & $FolderExecFiles[$b]
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
			_Log("$RunLine=" & $RunLine)
			Return Run($RunLine, $WorkingDir, @SW_SHOW, $STDIO_INHERIT_PARENT)

		Case "ps1"
			_Log("  ps1")
			$RunLine = @ComSpec & " /c " & "powershell.exe -ExecutionPolicy Bypass -File """ & $File & """ " & $Params
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
	If Not @error Then
		; If an IP hasn't been aquired it wont be an array
		If IsArray($Win32_NetworkAdapterConfiguration.IPAddress) Then $StatusbarText &= $Delimiter & $Win32_NetworkAdapterConfiguration.IPAddress[0]
		If IsArray($Win32_NetworkAdapterConfiguration.DefaultIPGateway) Then $StatusbarToolTipText &= "Gateway: " & $Win32_NetworkAdapterConfiguration.DefaultIPGateway[0] & "  "
		If IsArray($Win32_NetworkAdapterConfiguration.DNSServerSearchOrder) Then $StatusbarToolTipText &= "DNS: " & $Win32_NetworkAdapterConfiguration.DNSServerSearchOrder[0]
	EndIf
	$StatusbarToolTipText &= @CR & "Other IPs: " & @IPAddress1 & ", " & @IPAddress2 & ", " & @IPAddress3 & ", " & @IPAddress4

	; Get memory information
	If Not IsDeclared("_MemStats") Then Global $_MemStats = MemGetStats()
	$StatusbarText &= $Delimiter & Round($_MemStats[1] / 1024 / 1024) & "GB"

	; Get CPU information
	$Win32_Processor = _WMI("SELECT NumberOfCores,NumberOfLogicalProcessors FROM Win32_Processor")
	If Not @error Then $StatusbarText &= $Delimiter & $Win32_Processor.NumberOfCores & "/" & $Win32_Processor.NumberOfLogicalProcessors & " Cores"

	; Get motherboard bios information
	$Win32_BIOS = _WMI("SELECT SerialNumber,SMBIOSBIOSVersion,ReleaseDate FROM Win32_BIOS")
	If Not @error Then
		If $Win32_BIOS.SerialNumber <> "" and $Win32_BIOS.SerialNumber <> "System Serial Number" Then $StatusbarText &= $Delimiter & StringLeft($Win32_BIOS.SerialNumber, 20)
		$StatusbarText &= $Delimiter & "FW: " & StringLeft($Win32_BIOS.SMBIOSBIOSVersion, 20) & " (" & StringLeft($Win32_BIOS.ReleaseDate, 8) & ")"
	EndIf

	; Get additional statusbar and tool tip text
	$HelperStatusFiles = _FileListToArray(@TempDir, "Helper_Status_*.txt", $FLTA_FILES, True)
	For $i = 1 To Ubound($HelperStatusFiles) - 1
		If _FileModifiedAge($HelperStatusFiles[$i]) < 10 * 1000 Then
			$FileText = FileReadLine($HelperStatusFiles[$i], 1)
			_Log("$FileText=" & $FileText, 3)
			If Not @error Then $StatusbarText &= $Delimiter & $FileText

			$FileText = FileReadLine($HelperStatusFiles[$i], 2)
			_Log("$FileText=" & $FileText, 3)
			If Not @error Then $StatusBarToolTipText &= @CRLF & $FileText

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
	If _GUIToolTip_GetText($StatusBarToolTip, 0, $StatusBar1) <> $StatusBarToolTipText Then
		_GUIToolTip_UpdateTipText($StatusBarToolTip, 0, $StatusBar1, $StatusBarToolTipText)
		_Log("Statusbar Tooltip Updated", 3)
	EndIf



	If $Debug Then
		$StatusBarTimer2 = TimerInit()
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
