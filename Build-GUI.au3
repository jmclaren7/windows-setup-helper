#RequireAdmin
;===============================================================================
; Windows Setup Helper - Build Script Tool for Custom Windows ISOs
; Warning: This script uses Maps, a beta feature in the recent release versions of AutoIt
;===============================================================================
#include <AutoItConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiEdit.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <ScrollBarsConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#include <WinAPISys.au3>
#include <Array.au3>

#include "Helper\IncludeExt\Console.au3"
#include "Helper\IncludeExt\CommonFunctions.au3"

;===============================================================================
; Global Variables
;===============================================================================
Global $Title = "WSHelper Build Tool"
Global $Version = "1.0"
Global $TitleFull = $Title & " v" & $Version
Global $ConfigFile = @ScriptDir & "\Build.ini"
Global $HelperRepo = @ScriptDir
Global $DefaultADKPackages = "WinPE-WMI.cab|WinPE-NetFx.cab|WinPE-Scripting.cab|WinPE-PowerShell.cab|WinPE-StorageWMI.cab|WinPE-SecureBootCmdlets.cab|WinPE-SecureStartup.cab|WinPE-DismCmdlets.cab|WinPE-EnhancedStorage.cab|WinPE-Dot3Svc.cab|WinPE-FMAPI.cab|WinPE-FontSupport-WinRE.cab|WinPE-PlatformId.cab|WinPE-WDS-Tools.cab|WinPE-HTA.cab|WinPE-WinReCfg.cab"
Global $GUIMain
Global $IsRunning = False
Global $ProgramActive = True
Global $BootWIMMounted = False
Global $ADKVersionLabel
Global $ADKVersion = "Not detected"
Global $ADKPackagesPopulated = False

; Config Defaults
Global $SourceISOPath = "Windows11.iso"
Global $ISOTempPath = @ScriptDir & "\ISO-Temp"
Global $BootWIMPath = ""
Global $BootWIMIndex = "Microsoft Windows Setup (amd64)"
Global $WIMMountPath = @TempDir & "\WSHelper-WIMMount2"
Global $ADKPath = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Kits\Installed Roots", "KitsRoot10") 
Global $AddBootFilesPath = ""
Global $AddISOFilesPath = ""
Global $OutputISOPath = "Windows11-Output.iso"

;===============================================================================
; Start Program
;===============================================================================

; Change working directory to script location
FileChangeDir(@ScriptDir)

; Allocate console for logging
_Console_Alloc()
Local $hConsoleWnd = _Console_GetWindow()
_Console_SetTitle("Log - " & $Title)
WinMove($hConsoleWnd, "", 10, 10, 800, 400)

_Log("Starting " & $TitleFull)

; Register exit function
OnAutoItExitRegister("_Exit")

; Settings and GUI
_LoadSettings()
_CreateGUI()
GUISetState(@SW_SHOW, $GUIMain)

; Disable boot.wim input for now and auto-generate it
GUICtrlSetState($mBootWIMPath["Input"], $GUI_DISABLE)
GUICtrlSetState($mBootWIMPath["Button"], $GUI_DISABLE)

; !!!!!!  This build script is experimental !!!!!!!!
MsgBox(48, "Warning - " & $TitleFull, "This build script is experimental and incomplete" & @CRLF & @CRLF & "Use with caution and verify all specified paths are safe to work with.", 0, $GUIMain)

; Periodic GUI updates
Global $AdlibTimer = 1000
_ReadGUI()
_GUIChecks()


While 1
	Local $nMsg = GUIGetMsg()
	Local $ctrlFocus = _ControlGetFocus($GUIMain)

    Switch $nMsg
        Case 0, -11, -7, -9, -4
            ; Do nothing (idle/resize/focus messages)

		Case $GUI_EVENT_CLOSE
			If $BootWIMMounted Then
				Local $confirm = MsgBox(49, "WIM Still Mounted", "A WIM image is still mounted." & @CRLF & @CRLF & "Are you sure you want to exit?" & @CRLF & "(The mounted WIM will remain mounted)", 0, $GUIMain)
				If $confirm <> 1 Then ContinueLoop ; User clicked Cancel
			EndIf
			Exit

		Case $btnRun
			If Not $IsRunning Then _RunSelectedSteps()

		Case $btnSelectAll
			_SetAllCheckboxes($GUI_CHECKED)

		Case $btnSelectNone
			_SetAllCheckboxes($GUI_UNCHECKED)

		; Browse buttons
		Case $mSourceISOPath["Button"]
			Local $FileSelect = FileOpenDialog("Select Source ISO", "", "ISO Files (*.iso)|All Files (*.*)", $FD_FILEMUSTEXIST, "", $GUIMain)
			If Not @error Then
				GUICtrlSetData($mSourceISOPath["Input"], $FileSelect)
				; Auto-generate output ISO name if empty
				If GUICtrlRead($mOutputISOPath["Input"]) = "" Then
					_AutoGenerateOutputISO($FileSelect)
				EndIf
			EndIf

		Case $mISOTempPath["Button"]
			Local $folder = FileSelectFolder("Select Temp Folder", "", 0, GUICtrlRead($mISOTempPath["Input"]), $GUIMain)
			If Not @error Then
                GUICtrlSetData($mISOTempPath["Input"], $folder)
                ; Update boot.wim path based on new temp path
                GUICtrlSetData($mBootWIMPath["Input"], $folder & "\sources\boot.wim")
            EndIf

		Case $mBootWIMPath["Button"]
			Local $file = FileOpenDialog("Select Source WIM", "", "WIM Files (*.wim)|All Files (*.*)", $FD_FILEMUSTEXIST, "", $GUIMain)
			If Not @error Then GUICtrlSetData($mBootWIMPath["Input"], $file)

		Case $mAddBootFilesPath["Button"]
			Local $folder = FileSelectFolder("Select Extra Files Folder", "", 0, GUICtrlRead($mAddBootFilesPath["Input"]), $GUIMain)
			If Not @error Then GUICtrlSetData($mAddBootFilesPath["Input"], $folder)

		Case $mAddISOFilesPath["Button"]
			Local $folder = FileSelectFolder("Select Extra ISO Files Folder", "", 0, GUICtrlRead($mAddISOFilesPath["Input"]), $GUIMain)
			If Not @error Then GUICtrlSetData($mAddISOFilesPath["Input"], $folder)

		Case $mOutputISOPath["Button"]
			Local $file = FileSaveDialog("Select Output ISO", "", "ISO Files (*.iso)", $FD_PROMPTOVERWRITE, "", $GUIMain)
			If Not @error Then GUICtrlSetData($mOutputISOPath["Input"], $file)

		Case $mWIMMountPath["Button"]
			If $BootWIMMounted Then
				MsgBox(48, "Mount Path Locked", "The WIM mount path cannot be changed while a WIM image is mounted." & @CRLF & @CRLF & "Please unmount the current WIM before changing the mount path.", 0, $GUIMain)
			Else
				Local $folder = FileSelectFolder("Select WIM Mount Folder", "", 0, GUICtrlRead($mWIMMountPath["Input"]), $GUIMain)
				If Not @error Then GUICtrlSetData($mWIMMountPath["Input"], $folder)
			EndIf

		Case $mADKPath["Button"]
			Local $folder = FileSelectFolder("Select ADK Path", "", 0, GUICtrlRead($mADKPath["Input"]), $GUIMain)
			If Not @error Then GUICtrlSetData($mADKPath["Input"], $folder)

		; Tools buttons
		Case $ToolBrowseMountButton
			_BrowseMountFolder()

		Case $ToolDiscardUnmountButton
			_UnmountDiscard()

		Case $ToolGetInfoButton
			_GetImageInfo()

		Case $btnSave
			_SaveSettings()

		Case $btnExit
			If $BootWIMMounted Then
				Local $confirm = MsgBox(49, "WIM Still Mounted", "A WIM image is still mounted." & @CRLF & @CRLF & "Are you sure you want to exit?" & @CRLF & "(The mounted WIM will remain mounted)", 0, $GUIMain)
				If $confirm <> 1 Then ContinueLoop
			EndIf
			Exit

		; Go buttons for individual steps
		Case $mExtractISO["Button"]
            _ExtractISO()

        Case $mMountWIM["Button"]
            _MountBootWIM()

		Case $mCopyFiles["Button"]
            _CopyFiles()

		Case $mAddPackages["Button"]
            _AddPackages()

		Case $mDisableDPI["Button"]
            _DisableDPIScaling()

		Case $mUnmountCommit["Button"]
            _UnmountCommit()

		Case $mTrimImages["Button"]
            _TrimBootWIM()

		Case $mRemoveInstaller["Button"]
            _RemoveInstallWIM()

		Case $mMakeISO["Button"]
            _MakeISO()

		; Package ListView buttons
		Case $btnPkgSelectAll
			_GUICtrlListView_SetItemChecked($lvPackages, -1, True)

		Case $btnPkgSelectNone
			_GUICtrlListView_SetItemChecked($lvPackages, -1, False)

        Case $btnPkgSelectDefault
            _UpdateADKPackages()

		Case Else
			_Log("GUI Message: " & $nMsg)

	EndSwitch

	; Update GUI state on any non-idle message
	If $nMsg <> 0 And $nMsg <> -11 And $nMsg <> -7 And $nMsg <> -9 And $nMsg <> -4 Then
		_ReadGUI()
		_GUIChecks()
	EndIf

    ; Check for program window activation/deactivation and make sure both windows are visible when activating either
    $ActiveWindow = WinGetHandle("[ACTIVE]")
    If $ProgramActive = False and ($ActiveWindow = $hConsoleWnd or $ActiveWindow = $GUIMain) Then
        WinSetOnTop($hConsoleWnd, "", 1)
        WinSetOnTop($hConsoleWnd, "", 0)
        WinSetOnTop($GUIMain, "", 1)
        WinSetOnTop($GUIMain, "", 0)
        $ProgramActive = True
    ElseIf $ProgramActive = True and ($ActiveWindow <> $hConsoleWnd and $ActiveWindow <> $GUIMain) Then
        $ProgramActive = False
    EndIf
	Sleep(20)
WEnd

;===============================================================================
; GUI Creation
;===============================================================================
Func _CreateGUI()
	Local $guiWidth = 690, $guiHeight = 580

	$GUIMain = GUICreate($TitleFull, $guiWidth, $guiHeight, -1, -1)

	Local $yPos = 10
    Local $xPos = 20
	Local $stepHeight = 25
	Local $labelWidth = 80
	Local $inputWidth = 480
    Local $btnText = "Browse..."
	Local $btnWidth = 75

	; ============ Path Inputs Group ============
	$hGroup = GUICtrlCreateGroup("Paths", 10, $yPos, $guiWidth - 20, 260)
	$yPos += 20
    Global $mSourceISOPath = _CreateInputRow("Source ISO:", $xPos, $yPos, $inputWidth, $SourceISOPath, "Path to the Windows ISO file to customize", $btnText)
    $yPos += 30
    Global $mISOTempPath = _CreateInputRow("Temp Folder:", $xPos, $yPos, $inputWidth, $ISOTempPath, "Temporary folder where the ISO will be extracted and modified", $btnText)
    $yPos += 30
    Global $mBootWIMPath = _CreateInputRow("Boot.wim:", $xPos, $yPos, $inputWidth, $BootWIMPath, "Path to boot.wim file to mount and modify (usually in Temp\sources\)", $btnText)
    $yPos += 30
    Global $mBootWIMIndex = _CreateInputRow("Boot.wim Index:", $xPos, $yPos, $inputWidth, $BootWIMIndex, "Image index or name to mount from boot.wim (e.g., 2 or 'Microsoft Windows Setup (amd64)')", "")
    $yPos += 30
	Global $mWIMMountPath = _CreateInputRow("WIM Mount:", $xPos, $yPos, $inputWidth, $WIMMountPath, "Path where the boot.wim image will be mounted", $btnText)
    $yPos += 30
    Global $mADKPath = _CreateInputRow("ADK Path:", $xPos, $yPos, $inputWidth, $ADKPath, "Path to Windows Assessment and Deployment Kit (ADK) installation", $btnText)
    $yPos += 30
	$ADKVersionLabel = GUICtrlCreateLabel("Detected version: ", $xPos + 100, $yPos - 5, 200, 17)
	;GUICtrlSetFont($ADKVersionLabel, 8)
	$yPos += 20
    Global $mAddBootFilesPath = _CreateInputRow("Add to PE:", $xPos, $yPos, $inputWidth, $AddBootFilesPath, "Folder containing additional files to copy into the boot.wim image", $btnText)
    $yPos += 30
    Global $mAddISOFilesPath = _CreateInputRow("Add to ISO:", $xPos, $yPos, $inputWidth, $AddISOFilesPath, "Folder containing additional files to copy into the ISO root", $btnText)
    $yPos += 30
    Global $mOutputISOPath = _CreateInputRow("Output ISO:", $xPos, $yPos, $inputWidth, $OutputISOPath, "Path where the customized ISO will be saved", $btnText)
    $yPos += 20

	; Close group and set height
	GUICtrlCreateGroup("", -99, -99, 1, 1) ; Close group
	GUICtrlSetPos($hGroup, Default, Default, Default, $yPos)
	$yPos += 20

	; New row Y position for the next groups
	Local $RowStartY = $yPos

	; ============ Build Steps Group ============
	Local $chkX = 20, $chkWidth = 200, $goX = 220, $goWidth = 35
	Local $stepY = $RowStartY
	$hGroup = GUICtrlCreateGroup("Build Steps", 10, $stepY, 257, 100)
	$stepY += 20

    Global $mExtractISO = _CreateStepRow("1. Extract ISO to Temp folder", $chkX, $stepY, $chkWidth, True, "Extract the contents of the source ISO to the temp folder using 7-Zip", "Go", $goWidth)
    $stepY += $stepHeight

    Global $mMountWIM = _CreateStepRow("2. Mount boot.wim", $chkX, $stepY, $chkWidth, True, "Mount the boot.wim image so it can be modified", "Go", $goWidth)
    $stepY += $stepHeight

	Global $mCopyFiles = _CreateStepRow("3. Copy Helper files", $chkX, $stepY, $chkWidth, True, "Copy the Helper folder and extra files into the mounted WIM image", "Go", $goWidth)
    $stepY += $stepHeight

	Global $mAddPackages = _CreateStepRow("4. Add WinPE packages to image", $chkX, $stepY, $chkWidth, True, "Add selected WinPE optional components (packages) from the ADK", "Go", $goWidth)
    $stepY += $stepHeight

	Global $mDisableDPI = _CreateStepRow("5. Disable DPI scaling (registry)", $chkX, $stepY, $chkWidth, True, "Add registry keys to disable DPI scaling in WinPE (prevents blurry text on high-DPI displays)", "Go", $goWidth)
    $stepY += $stepHeight

	Global $mUnmountCommit = _CreateStepRow("6. Unmount and commit changes", $chkX, $stepY, $chkWidth, True, "Save all changes and unmount the WIM image", "Go", $goWidth)
    $stepY += $stepHeight

	Global $mTrimImages = _CreateStepRow("7. Trim boot.wim", $chkX, $stepY, $chkWidth, True, "Remove unused images from boot.wim to reduce file size", "Go", $goWidth)
    $stepY += $stepHeight

	Global $mRemoveInstaller = _CreateStepRow("8. Remove Install.wim", $chkX, $stepY, $chkWidth, False, "Delete install.wim/install.esd to create a boot-only ISO (no Windows installation)", "Go", $goWidth)
    $stepY += $stepHeight

	Global $mMakeISO = _CreateStepRow("9. Make ISO from Temp folder", $chkX, $stepY, $chkWidth, True, "Create a bootable ISO from the modified temp folder contents", "Go", $goWidth)
    $stepY += $stepHeight + 5

	; Buttons
	Global $btnSelectAll = GUICtrlCreateButton("All", $chkX, $stepY, 40, 25)
	GUICtrlSetTip(-1, "Check all build steps")
	Global $btnSelectNone = GUICtrlCreateButton("None", $chkX + 45, $stepY, 50, 25)
	GUICtrlSetTip(-1, "Uncheck all build steps")
	Global $btnRun = GUICtrlCreateButton("Run Selected", $goX - 110 + 35, $stepY, 110, 25)
	GUICtrlSetTip(-1, "Execute all checked build steps in order")
	$stepY += $stepHeight + 10

	; Close group and set height
	GUICtrlCreateGroup("", -99, -99, 1, 1) ; Close group
	Local $StepRowFinalHeight = $stepY - $RowStartY
	GUICtrlSetPos($hGroup, Default, Default, Default, $StepRowFinalHeight)
	$yPos = $stepY + 10

	; ============ ADK Packages Group ============ 
	Local $pkgGroupX = 273, $pkgGroupWidth = 260
	Local $pkgGroupHeight = $StepRowFinalHeight
	Local $pkgY = $RowStartY
	GUICtrlCreateGroup("WinPE Packages", $pkgGroupX, $pkgY, $pkgGroupWidth, $pkgGroupHeight)

	; Create ListView with checkboxes
	Local $ListViewHeight = $pkgGroupHeight - 65
	Global $lvPackages = GUICtrlCreateListView("Package Name", $pkgGroupX + 10, $pkgY + 20, $pkgGroupWidth - 20, $ListViewHeight, BitOR($LVS_REPORT, $LVS_SINGLESEL, $LVS_SHOWSELALWAYS))
	_GUICtrlListView_SetExtendedListViewStyle($lvPackages, BitOR($LVS_EX_CHECKBOXES, $LVS_EX_FULLROWSELECT))
	_GUICtrlListView_SetColumnWidth($lvPackages, 0, $pkgGroupWidth - 50)
	GUICtrlSetTip($lvPackages, "Select WinPE optional components to add to the boot image (Step 4)")
	$pkgY = $pkgY + $ListViewHeight + 30

	; Buttons for package selection
	Global $btnPkgSelectAll = GUICtrlCreateButton("All", $pkgGroupX + 10, $pkgY, 40, 25)
	GUICtrlSetTip(-1, "Select all available packages")
	Global $btnPkgSelectNone = GUICtrlCreateButton("None", $pkgGroupX + 55, $pkgY, 50, 25)
	GUICtrlSetTip(-1, "Deselect all packages")
    Global $btnPkgSelectDefault = GUICtrlCreateButton("Default", $pkgGroupX + 110, $pkgY, 60, 25)
	GUICtrlSetTip(-1, "Reset to default recommended packages")

	; Close group and set height
	GUICtrlCreateGroup("", -99, -99, 1, 1) ; Close group

	; ============  Tools Group ============ 
	Local $toolsGroupHeight = $StepRowFinalHeight
	Local $toolX = 550, $toolY = $RowStartY
	Local $toolBtnWidth = 120, $toolBtnHeight = 28
	GUICtrlCreateGroup("Tools", 540, $toolY, 140, $toolsGroupHeight)
	$toolY += 25

	Global $ToolBrowseMountButton = GUICtrlCreateButton("Browse Mount", $toolX, $toolY, $toolBtnWidth, $toolBtnHeight)
	GUICtrlSetState(-1, $GUI_DISABLE) ; Disabled until WIM is detected
	GUICtrlSetTip(-1, "Open the mounted WIM folder in File Explorer")
	$toolY += 35

	Global $ToolDiscardUnmountButton = GUICtrlCreateButton("Discard && Unmount", $toolX, $toolY, $toolBtnWidth, $toolBtnHeight)
	GUICtrlSetState(-1, $GUI_DISABLE) ; Disabled until WIM is detected
	GUICtrlSetTip(-1, "Unmount the WIM without saving changes (discard all modifications)")
	$toolY += 35

	Global $ToolGetInfoButton = GUICtrlCreateButton("Get Boot.wim Info", $toolX, $toolY, $toolBtnWidth, $toolBtnHeight)
	GUICtrlSetTip(-1, "Display information about images in the boot.wim file")

	; Populate the packages ListView
	_UpdateADKPackages()

	; Close group
	GUICtrlCreateGroup("", -99, -99, 1, 1) ; Close group

	; ============  Save/Exit Buttons ============ 
	Local $btnX = 500, $btnWidth = 80, $btnHeight = 28
	
	Global $btnSave = GUICtrlCreateButton("Save", $btnX, $yPos, $btnWidth, $btnHeight)
	GUICtrlSetTip(-1, "Save current settings to Build.ini")
	Global $btnExit = GUICtrlCreateButton("Exit", $btnX + $btnWidth + 10, $yPos, $btnWidth, $btnHeight)
	GUICtrlSetTip(-1, "Exit the program")
	$yPos += $btnHeight + 40

	; Final adjustment to GUI height based on content
	WinMove($GUIMain, "", Default, Default, Default, $yPos)

EndFunc   ;==>_CreateGUI

;===============================================================================
; Create Input Row (Label, Left, Top, Width, DefaultValue, Tip)
;===============================================================================
Func _CreateInputRow($Label, $Left, $Top, $Width, $Value = "", $Tip= "", $ButtonText = "")
    Local $mControls[]

    $mControls["Label"] = GUICtrlCreateLabel($Label, $Left, $Top + 3, 100, 17)
	GUICtrlSetTip(-1, $Tip)

    $mControls["Alert"]  = GUICtrlCreateLabel("", $Left + 74, $Top + 2, 4, 17)
    GUICtrlSetBkColor(-1, 0xFF0000)
    GUICtrlSetState(-1, $GUI_HIDE)
    GUICtrlSetTip(-1, $Tip)

    $mControls["Input"] = GUICtrlCreateInput($Value, $Left + 80, $Top, $Width, 21)
    GUICtrlSetTip(-1, $Tip)

    If $ButtonText <> "" Then
        $mControls["Button"] = GUICtrlCreateButton($ButtonText, 590, $Top - 2, 75, 25)
        GUICtrlSetTip(-1, $Tip)
    EndIf

    Return $mControls
EndFunc  ;==>_CreateInputRow

;===============================================================================
; Create Step Row (Checkbox with Go button)
;===============================================================================
Func _CreateStepRow($Label, $Left, $Top, $ChkWidth, $Checked = True, $Tip = "", $ButtonText = "Go", $ButtonWidth = 35)
    Local $mControls[]

    $mControls["Checkbox"] = GUICtrlCreateCheckbox($Label, $Left, $Top, $ChkWidth, 20)
    If $Checked Then GUICtrlSetState(-1, $GUI_CHECKED)
    GUICtrlSetTip(-1, $Tip)

    If $ButtonText <> "" Then
        $mControls["Button"] = GUICtrlCreateButton($ButtonText, $Left + $ChkWidth + 5, $Top - 2, $ButtonWidth, 22)
        GUICtrlSetTip(-1, $Tip)
    EndIf

    Return $mControls
EndFunc  ;==>_CreateStepRow

;===============================================================================
; Set All Checkboxes
;===============================================================================
Func _SetAllCheckboxes($State)
    _Log("Setting all checkboxes to state: " & $State)

	GUICtrlSetState($mExtractISO["Checkbox"], $State)
	GUICtrlSetState($mMountWIM["Checkbox"], $State)
	GUICtrlSetState($mCopyFiles["Checkbox"], $State)
	GUICtrlSetState($mAddPackages["Checkbox"], $State)
	GUICtrlSetState($mDisableDPI["Checkbox"], $State)
	GUICtrlSetState($mUnmountCommit["Checkbox"], $State)
	GUICtrlSetState($mTrimImages["Checkbox"], $State)
	If $State = $GUI_UNCHECKED Then GUICtrlSetState($mRemoveInstaller["Checkbox"], $State) ; Note: Remove Installer is excluded from select All
	GUICtrlSetState($mMakeISO["Checkbox"], $State)
EndFunc   ;==>_SetAllCheckboxes

;===============================================================================
; Auto-generate Output ISO name
;===============================================================================
Func _AutoGenerateOutputISO($sourceFile)
    _Log("Auto ISO name: " & $sourceFile)

	Local $dateTime = StringRight(@YEAR, 2) & @MON & @MDAY & "_" & @HOUR & @MIN
	Local $Output = StringTrimRight($sourceFile, 4) & "_WSHelper_" & $dateTime & ".iso"

	If IsDeclared("mOutputISOPath") Then GUICtrlSetData($mOutputISOPath["Input"], $Output)
    Return $Output
EndFunc   ;==>_AutoGenerateOutputISO

;===============================================================================
; Update adkPackages list from adk path
;===============================================================================
Func _UpdateADKPackages()
    _Log("Updating ADK Packages list from ADK path")

    ; Clear existing items
    _GUICtrlListView_DeleteAllItems($lvPackages)

    Local $adkOCsubpath = "\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs"
    Local $adkPath = GUICtrlRead($mADKPath["Input"]) & $adkOCsubpath
    If Not FileExists($adkPath) Then
        _Log("Error: ADK WinPE OCs path not found: " & $adkPath)
        Return
    EndIf

    ; Get CAB files from root (not language subfolders)
    Local $cabFiles = _FileListToArray($adkPath, "WinPE-*.cab", $FLTA_FILES, False)
    If @error Then
        _Log("Error: No CAB files found in ADK WinPE OCs path")
        Return
    EndIf

    ; Parse default packages into array for quick lookup
    Local $aDefaults = StringSplit($DefaultADKPackages, "|")

    ; Add items to ListView
    For $i = 1 To $cabFiles[0]
        Local $idx = _GUICtrlListView_AddItem($lvPackages, $cabFiles[$i])

        ; Check if this package is in the defaults list
        For $j = 1 To $aDefaults[0]
            If $cabFiles[$i] = $aDefaults[$j] Then
                _GUICtrlListView_SetItemChecked($lvPackages, $idx, True)
                ExitLoop
            EndIf
        Next
    Next

    $ADKPackagesPopulated = True
    _Log("Found " & $cabFiles[0] & " ADK packages")
EndFunc   ;==>_UpdateADKPackages

;===============================================================================
; Checks for GUI requirements (called by AdlibRegister)
;===============================================================================
Func _GUIChecks()
    ; Set or reset adlib timer to avoid overlapping calls
    AdlibRegister("_GUIChecks", $AdlibTimer)

    ; Get relevant environment states
    Local $Alert
    Global $WIMMountExists = FileExists($WIMMountPath)
    Global $BootWIMExists = FileExists($BootWIMPath)
    Global $SourceISOExists = FileExists($SourceISOPath)
    Global $ADKExists = FileExists($ADKPath & "\Assessment and Deployment Kit\Windows Preinstallation Environment")
    Global $TempExists = FileExists($ISOTempPath)
    Global $AddBootFilesExists = FileExists($AddBootFilesPath)
    Global $AddISOFilesExists = FileExists($AddISOFilesPath)

	Local $NewADKVersion = FileGetVersion($ADKPath & "\Assessment and Deployment Kit\Deployment Tools\amd64\DISM\Dism.exe")
	If $NewADKVersion <> $ADKVersion Then
		$ADKVersion = $NewADKVersion
		_Log("Detected ADK version: " & $ADKVersion)
		GUICtrlSetData($ADKVersionLabel, "ADK Version: " & $ADKVersion)
	EndIf

    ; SourceISOPath
    $Alert = BitAND(GUICtrlGetState($mSourceISOPath["Alert"]), $GUI_SHOW)
    If $SourceISOExists And $Alert Then
        GUICtrlSetState($mSourceISOPath["Alert"], $GUI_HIDE)
        GUICtrlSetState($mExtractISO["Button"], $GUI_ENABLE)
    ElseIf Not $SourceISOExists And Not $Alert Then
        GUICtrlSetState($mSourceISOPath["Alert"], $GUI_SHOW)
        GUICtrlSetState($mExtractISO["Button"], $GUI_DISABLE)
    EndIf

    ; TempPath
    $Alert = BitAND(GUICtrlGetState($mISOTempPath["Alert"]), $GUI_SHOW)
    If $TempExists And $Alert Then
        GUICtrlSetState($mISOTempPath["Alert"], $GUI_HIDE)
    ElseIf Not $TempExists And Not $Alert Then
        GUICtrlSetState($mISOTempPath["Alert"], $GUI_SHOW)
    EndIf

    ; BootWIMPath
    $Alert = BitAND(GUICtrlGetState($mBootWIMPath["Alert"]), $GUI_SHOW)
    If $BootWIMExists And $Alert Then
        GUICtrlSetState($mBootWIMPath["Alert"], $GUI_HIDE)
        GUICtrlSetState($ToolGetInfoButton, $GUI_ENABLE)
        GUICtrlSetState($mMountWIM["Button"], $GUI_ENABLE)
    ElseIf Not $BootWIMExists And Not $Alert Then
		GUICtrlSetBkColor($mBootWIMPath["Alert"], 0xFFFF00) ; Yellow
        GUICtrlSetState($mBootWIMPath["Alert"], $GUI_SHOW)
        GUICtrlSetState($ToolGetInfoButton, $GUI_DISABLE)
        GUICtrlSetState($mMountWIM["Button"], $GUI_DISABLE)
    EndIf

    ; Does ADK exist?
    $Alert = BitAND(GUICtrlGetState($mADKPath["Alert"]), $GUI_SHOW)
    If $ADKExists And $Alert Then
        GUICtrlSetState($mADKPath["Alert"], $GUI_HIDE)
        GUICtrlSetState($mAddPackages["Button"], $GUI_ENABLE)
        GUICtrlSetState($lvPackages, $GUI_ENABLE)
        If Not $ADKPackagesPopulated Then _UpdateADKPackages()
    ElseIf Not $ADKExists And Not $Alert Then
        $ADKPackagesPopulated = False
        GUICtrlSetState($mADKPath["Alert"], $GUI_SHOW)
        GUICtrlSetState($mAddPackages["Button"], $GUI_DISABLE)
        GUICtrlSetState($lvPackages, $GUI_DISABLE)
    EndIf

    ; AddBootFilesPath
    $Alert = BitAND(GUICtrlGetState($mAddBootFilesPath["Alert"]), $GUI_SHOW)
    If ($AddBootFilesExists Or $AddBootFilesPath = "") And $Alert Then
        GUICtrlSetState($mAddBootFilesPath["Alert"], $GUI_HIDE)
    ElseIf Not $AddBootFilesExists And $AddBootFilesPath <> "" And Not $Alert Then
        GUICtrlSetState($mAddBootFilesPath["Alert"], $GUI_SHOW)
    EndIf

    ; AddISOFilesPath
    $Alert = BitAND(GUICtrlGetState($mAddISOFilesPath["Alert"]), $GUI_SHOW)
    If ($AddISOFilesExists Or $AddISOFilesPath = "") And $Alert Then
        GUICtrlSetState($mAddISOFilesPath["Alert"], $GUI_HIDE)
    ElseIf Not $AddISOFilesExists And $AddISOFilesPath <> "" And Not $Alert Then
        GUICtrlSetState($mAddISOFilesPath["Alert"], $GUI_SHOW)
    EndIf

    ; Does mount path exist? (WIM probably mounted)
    $GUIEnabled = BitAND(GUICtrlGetState($ToolBrowseMountButton), $GUI_ENABLE)
    If $WIMMountExists And Not $GUIEnabled Then
        GUICtrlSetState($ToolBrowseMountButton, $GUI_ENABLE)
        GUICtrlSetState($ToolDiscardUnmountButton, $GUI_ENABLE)
        ;GUICtrlSetState($mMountWIM["Button"], $GUI_DISABLE)
    Elseif Not $WIMMountExists And $GUIEnabled Then
        GUICtrlSetState($ToolBrowseMountButton, $GUI_DISABLE)
        GUICtrlSetState($ToolDiscardUnmountButton, $GUI_DISABLE)
        ;GUICtrlSetState($mMountWIM["Button"], $GUI_ENABLE)
    EndIf

EndFunc   ;==>_GUIChecks

;===============================================================================
; Read GUI Inputs into Global Variables
;===============================================================================
Func _ReadGUI()
    ; Set or reset Adlib timer to avoid overlapping calls
    AdlibRegister("_ReadGUI", $AdlibTimer)

    ; Update global paths from GUI
	Global $SourceISOPath = GUICtrlRead($mSourceISOPath["Input"])
	Local $NewTempPath = GUICtrlRead($mISOTempPath["Input"])
	; Auto-sync boot.wim path when temp path changes (since boot.wim input is disabled)
	If $NewTempPath <> $ISOTempPath Then
		GUICtrlSetData($mBootWIMPath["Input"], $NewTempPath & "\sources\boot.wim")
	EndIf
	Global $ISOTempPath = $NewTempPath
	Global $BootWIMPath = GUICtrlRead($mBootWIMPath["Input"])
	Global $AddBootFilesPath = GUICtrlRead($mAddBootFilesPath["Input"])
	Global $AddISOFilesPath = GUICtrlRead($mAddISOFilesPath["Input"])
	Global $OutputISOPath = GUICtrlRead($mOutputISOPath["Input"])
	Global $WIMMountPath = GUICtrlRead($mWIMMountPath["Input"])
	Global $ADKPath = GUICtrlRead($mADKPath["Input"])
	Global $BootWIMIndex = GUICtrlRead($mBootWIMIndex["Input"])
EndFunc   ;==>_ReadGUI

;===============================================================================
; Load Settings
;===============================================================================
Func _LoadSettings()
	$SourceISOPath = IniRead($ConfigFile, "Paths", "SourceISOPath", $SourceISOPath)
	$ISOTempPath = IniRead($ConfigFile, "Paths", "ISOTempPath", $ISOTempPath)
    $BootWIMPath = IniRead($ConfigFile, "Paths", "BootWIMPath", $ISOTempPath & "\sources\boot.wim")
	$BootWIMIndex = IniRead($ConfigFile, "Paths", "BootWIMIndex", $BootWIMIndex)
	$WIMMountPath = IniRead($ConfigFile, "Paths", "WIMMountPath", $WIMMountPath)

	$ADKPath = IniRead($ConfigFile, "Paths", "ADKPath", $ADKPath)
    $AddBootFilesPath = IniRead($ConfigFile, "Paths", "AddBootFilesPath", "")

	$AddISOFilesPath = IniRead($ConfigFile, "Paths", "AddISOFilesPath", "")
	$OutputISOPath = IniRead($ConfigFile, "Paths", "OutputISOPath", $OutputISOPath)

EndFunc   ;==>_LoadSettings

;===============================================================================
; Save Settings to INI
;===============================================================================
Func _SaveSettings()
	IniWrite($ConfigFile, "Paths", "SourceISOPath", GUICtrlRead($mSourceISOPath["Input"]))
	IniWrite($ConfigFile, "Paths", "ISOTempPath", GUICtrlRead($mISOTempPath["Input"]))
	IniWrite($ConfigFile, "Paths", "BootWIMPath", GUICtrlRead($mBootWIMPath["Input"]))
	IniWrite($ConfigFile, "Paths", "BootWIMIndex", GUICtrlRead($mBootWIMIndex["Input"]))
	IniWrite($ConfigFile, "Paths", "WIMMountPath", GUICtrlRead($mWIMMountPath["Input"]))

	IniWrite($ConfigFile, "Paths", "AddBootFilesPath", GUICtrlRead($mAddBootFilesPath["Input"]))
	IniWrite($ConfigFile, "Paths", "AddISOFilesPath", GUICtrlRead($mAddISOFilesPath["Input"]))

	IniWrite($ConfigFile, "Paths", "ADKPath", GUICtrlRead($mADKPath["Input"]))
	IniWrite($ConfigFile, "Paths", "OutputISOPath", GUICtrlRead($mOutputISOPath["Input"]))
	
EndFunc   ;==>_SaveSettings

;===============================================================================
; Run Command and Log Output
;===============================================================================
Func _RunCmd($sCommand, $sDescription = "")
    If $sDescription <> "" Then
		_Log("=== " & $sDescription & " ===")
	EndIf

    ; Disable GUI interaction during command execution
    GUISetState(@SW_DISABLE, $GUIMain)
    WinSetTrans($GUIMain, "", 220)
    WinSetOnTop($hConsoleWnd, "", 1)
    WinSetOnTop($hConsoleWnd, "", 0)

    _Log("Command: " & $sCommand)
	Local $iPID = Run(@ComSpec & ' /c "' & $sCommand & '"', @ScriptDir, @SW_HIDE, $STDERR_MERGED)
	If @error Then
		_Log("Error: Failed to run command")
		Return SetError(1, 0, "")
	EndIf

	Local $sOutput = ""
	While 1
		$sOutput &= StdoutRead($iPID)
		If @error Then ExitLoop
		Sleep(50)
	WEnd

	; Wait for process to finish and get exit code
	ProcessWaitClose($iPID)
	Local $iExitCode = @extended

    ; Re-enable GUI interaction
    GUISetState(@SW_ENABLE, $GUIMain)
    WinSetTrans($GUIMain, "", 255)

	; Log output
	If StringStripWS($sOutput, 3) <> "" Then
		_Console_SetTextAttribute(-1, BitOR($FOREGROUND_BLUE, $FOREGROUND_GREEN, $FOREGROUND_INTENSITY)) ; Cyan
		_Log("Command Output:" & @CRLF & StringStripWS($sOutput, 3))
		_Console_SetTextAttribute(-1, BitOR($FOREGROUND_RED, $FOREGROUND_GREEN, $FOREGROUND_BLUE)) ; Reset to white
	EndIf

	If $iExitCode <> 0 Then
		_Log("Command failed with exit code: " & $iExitCode)
		Return SetError(1, $iExitCode, $sOutput)
	EndIf

	_Log("Command completed successfully.")
	Return SetError(0, 0, $sOutput)
EndFunc   ;==>_RunCmd

;===============================================================================
; Run Command Silently (no logging, for background checks)
;===============================================================================
Func _RunCmdSilent($sCommand)
	Local $iPID = Run(@ComSpec & ' /c "' & $sCommand & '"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	If @error Then Return ""

	Local $sOutput = ""
	While 1
		$sOutput &= StdoutRead($iPID)
		If @error Then ExitLoop
		Sleep(10)
	WEnd

	ProcessWaitClose($iPID)
	Return $sOutput
EndFunc   ;==>_RunCmdSilent

;===============================================================================
; Run Selected Steps
;===============================================================================
Func _RunSelectedSteps()
    _Log("Starting build process...")

	$IsRunning = True

    _ReadGUI()

	; Save settings
	_SaveSettings()

	; Validate paths
	If Not _ValidatePaths() Then
		$IsRunning = False
		Return
	EndIf

	Local $Success = True

	; Execute checked steps in order
	If BitAND(GUICtrlRead($mExtractISO["Checkbox"]), $GUI_CHECKED) And $Success Then
		$Success = _ExtractISO()
	EndIf

	If BitAND(GUICtrlRead($mMountWIM["Checkbox"]), $GUI_CHECKED) And $Success Then
		$Success = _MountBootWIM()
	EndIf

	If BitAND(GUICtrlRead($mCopyFiles["Checkbox"]), $GUI_CHECKED) And $Success Then
		$Success = _CopyFiles()
	EndIf

	If BitAND(GUICtrlRead($mAddPackages["Checkbox"]), $GUI_CHECKED) And $Success Then
		$Success = _AddPackages()
	EndIf

	If BitAND(GUICtrlRead($mDisableDPI["Checkbox"]), $GUI_CHECKED) And $Success Then
		$Success = _DisableDPIScaling()
	EndIf

	If BitAND(GUICtrlRead($mUnmountCommit["Checkbox"]), $GUI_CHECKED) And $Success Then
		$Success = _UnmountCommit()
	EndIf

	If BitAND(GUICtrlRead($mTrimImages["Checkbox"]), $GUI_CHECKED) And $Success Then
		$Success = _TrimBootWIM()
	EndIf

	If BitAND(GUICtrlRead($mRemoveInstaller["Checkbox"]), $GUI_CHECKED) And $Success Then
		$Success = _RemoveInstallWIM()
	EndIf

	If BitAND(GUICtrlRead($mMakeISO["Checkbox"]), $GUI_CHECKED) And $Success Then
		$Success = _MakeISO()
	EndIf


	If $Success Then
		_Log("=== Completed Selected Steps ===")
	Else
		_Log("=== One or More Steps Report Errors ===")
	EndIf

	$IsRunning = False
EndFunc   ;==>_RunSelectedSteps

;===============================================================================
; Construct WIM Index Parameter
;===============================================================================
Func _ConstructIndexParam($Index)
    If StringIsInt($Index) Then
        Return '/Index:' & $Index
    Else
        Return '/Name:"' & $Index & '"'
    EndIf
EndFunc   ;==>_ConstructIndexParam

;===============================================================================
; Validate Paths
;===============================================================================
Func _ValidatePaths()
	; Check Source ISO exists (if extract step is checked)
	If BitAND(GUICtrlRead($mExtractISO["Checkbox"]), $GUI_CHECKED) Then
		If Not FileExists($SourceISOPath) Then
			_Log("Error: Source ISO not found: " & $SourceISOPath)
			MsgBox(16, "Error", "Source ISO file not found:" & @CRLF & $SourceISOPath)
			Return False
		EndIf
	EndIf

	; Check ADK path exists (if add packages or make iso is checked)
	If BitAND(GUICtrlRead($mAddPackages["Checkbox"]), $GUI_CHECKED) Or BitAND(GUICtrlRead($mMakeISO["Checkbox"]), $GUI_CHECKED) Then
		If Not FileExists($ADKPath) Then
			_Log("Error: ADK path not found: " & $ADKPath)
			MsgBox(16, "Error", "Windows ADK not found at:" & @CRLF & $ADKPath & @CRLF & @CRLF & "Please install Windows ADK or correct the path.")
			Return False
		EndIf
	EndIf

	; Check Helper\Main.au3 exists
	If Not FileExists($HelperRepo & "\Helper\Main.au3") Then
		_Log("Error: Helper\Main.au3 not found in script directory")
		MsgBox(16, "Error", "Helper\Main.au3 not found." & @CRLF & "Make sure the script is in the correct location.")
		Return False
	EndIf

	Return True
EndFunc   ;==>_ValidatePaths

;===============================================================================
; Step 1: Extract ISO
;===============================================================================
Func _ExtractISO()
	_Log("Extracting ISO to Temp folder")

	; Remove existing temp folder
	If FileExists($ISOTempPath) Then
		_Log("Removing existing temp folder...")
		DirRemove($ISOTempPath, 1)
	EndIf

	; Create temp folder
	DirCreate($ISOTempPath)

	; Extract using 7-Zip
	Local $7zPath = $HelperRepo & "\Helper\Tools\7-Zip\7z.exe"
	If Not FileExists($7zPath) Then
		_Log("Error: 7-Zip not found at: " & $7zPath)
		Return False
	EndIf

	Local $cmd = '"' & $7zPath & '" x -y -o"' & $ISOTempPath & '" "' & $SourceISOPath & '"'
	_RunCmd($cmd, "Extracting ISO")

	If @error Then Return False
	Return True
EndFunc   ;==>_ExtractISO

;===============================================================================
; Step 2: Mount WIM
;===============================================================================
Func _MountBootWIM()
	_Log("Mounting boot.wim")

	; Create mount folder
	If FileExists($WIMMountPath) Then
		_Log("Warning user about existing mount folder")
		Local $Message = "Mount folder already exists: " & $WIMMountPath & @CRLF & @CRLF & "This may indicate a WIM is already mounted or that a previous mount was not properly cleaned up." & @CRLF & @CRLF & "The existing mount folder will try to be used if you continue."
		Local $Msg = MsgBox($MB_ICONWARNING + $MB_OKCANCEL, $Title, $Message, 0, $GUIMain)
		If $Msg = $IDOK Then
			_Log("Continue with existing mount folder")
			DirRemove($WIMMountPath, 1) ; Try to clean up existing mount folder before mounting
			If @error Then
				_Log("Error: Could not remove existing mount folder, this indicated it may be in use by an existing mounted WIM. Please investigate and try again.")
				Return SetError(1, 0, False)
			EndIf
		Else
			_Log("Canceled due to existing mount folder")
			Return SetError(1, 0, False)
		EndIf
	EndIf

	DirCreate($WIMMountPath)

	Local $cmd = 'Dism /Mount-image /ImageFile:"' & $BootWIMPath & '" ' & _ConstructIndexParam($BootWIMIndex) & ' /MountDir:"' & $WIMMountPath & '" /Optimize'
	_RunCmd($cmd, "Mounting WIM")

	If @error Then Return SetError(1, 0, False)
	$BootWIMMounted = True
	Return True
EndFunc   ;==>_MountBootWIM

;===============================================================================
; Step 3: Copy Files
;===============================================================================
Func _CopyFiles()
	_Log("Copying Helper files")

	; Copy Helper folder
	_Log("Copying Helper folder...")
	Local $helperDest = $WIMMountPath & "\Helper"
	If FileExists($helperDest) Then
		_Log("Removing existing Helper folder...")
		DirRemove($helperDest, 1)
	EndIf
	DirCreate($helperDest)
	Local $result = DirCopy($HelperRepo & "\Helper", $helperDest, 1)
	If Not $result Then
		_Log("Error: Error copying Helper folder")
		Return False
	EndIf

	; Copy Windows folder
	_Log("Copying Windows folder...")
	$result = DirCopy($HelperRepo & "\Windows", $WIMMountPath & "\Windows", 1)
	If Not $result Then
		_Log("Error: Warning: Could not copy Windows folder (may not exist)")
	EndIf

	; Copy extra files from AddBootFilesPath if specified, the trailing folder can use wildcards (_FileListToArray)
	If $AddBootFilesPath <> "" Then
		_Log("AddBootFilesPath: " & $AddBootFilesPath)
		Local $AddBootFilesPath_Parent = StringLeft($AddBootFilesPath, StringInStr($AddBootFilesPath, "\", 0, -1) - 1)
		Local $AddBootFilesPath_Folder = StringTrimLeft($AddBootFilesPath, StringInStr($AddBootFilesPath, "\", 0, -1))

		Local $aFolders = _FileListToArray($AddBootFilesPath_Parent, $AddBootFilesPath_Folder, $FLTA_FOLDERS)
		For $i = 1 To $aFolders[0]
			_Log("Copying extra boot files from: " & $AddBootFilesPath_Parent & "\" & $aFolders[$i])
			Local $Source = $AddBootFilesPath_Parent & "\" & $aFolders[$i]
			Local $Destination = $WIMMountPath
			$result = DirCopy($Source, $Destination, 1)
			If Not $result Then
				_Log("Error: Warning: Could not copy files from: " & $Source & " to: " & $Destination)
			EndIf
		Next
	EndIf

	; Clean up logs and temp files
	_Log("Removing log files and temp files...")
	FileDelete($WIMMountPath & "\Auto-saved*.xml")
	FileDelete($WIMMountPath & "\*.log")

	_Log("Step completed successfully.")
	Return True
EndFunc   ;==>_CopyFiles

;===============================================================================
; Step 4: Add Packages
;===============================================================================
Func _AddPackages()
	_Log("Adding packages to image")

	Local $adkPackages = $ADKPath & "\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs"
    Local $langFolder = $adkPackages & "\en-us"

	If Not FileExists($adkPackages) Then
		_Log("Error: ADK packages not found at: " & $adkPackages)
		Return False
	EndIf

	_Log("Packages from: " & $adkPackages)

	; Run DISM commands for each checked item
    Local $aSelectedPackages = _GUICtrlListView_GetCheckedText($lvPackages)
    For $i = 1 To $aSelectedPackages[0]
        Local $pkgPath = $adkPackages & "\" & $aSelectedPackages[$i]
        Local $Command = 'Dism /Image:"' & $WIMMountPath & '" /Add-Package /PackagePath:"' & $pkgPath & '"'
        ; If $adkPackages\en-us folder exists, add language-specific package, file name are suffixed with language code
        Local $langPkgPath = $langFolder & "\" & StringRegExpReplace($aSelectedPackages[$i], "\.cab$", "_en-us.cab")
        If FileExists($langPkgPath) Then
            $Command &= ' /PackagePath:"' & $langPkgPath & '"'
        EndIf

        _RunCmd($Command, "Adding package: " & $aSelectedPackages[$i])
        If @error Then
            _Log("Error: Error adding package: " & $aSelectedPackages[$i])
        EndIf
    Next

	Return True
EndFunc   ;==>_AddPackages

;===============================================================================
; Step 5: Disable DPI Scaling
;===============================================================================
Func _DisableDPIScaling()
	_Log("Disabling DPI scaling (registry)")

	Local $regHive = $WIMMountPath & "\Windows\System32\config\default"
	Local $mountPath = "HKLM\_WinPE_Default"

	; Load registry hive
	_RunCmd('reg load ' & $mountPath & ' "' & $regHive & '"', "Loading registry hive")

	; Test if registry hive loaded successfully by querying a known key
	RegRead($mountPath & "\Environment", "Path")
	If @error Then
		_Log("Error: Failed to load registry hive")
		Return False
	EndIf

	; Add registry keys
	RegWrite ( $mountPath & "\Control Panel\Desktop", "LogPixels", "REG_DWORD", 96)
	RegWrite ( $mountPath & "\Control Panel\Desktop", "Win8DpiScaling", "REG_DWORD", 1)
	RegWrite ( $mountPath & "\Control Panel\Desktop", "DpiScalingVer", "REG_DWORD", 4120)

	; Unload registry hive
	_RunCmd('reg unload ' & $mountPath, "Unloading registry hive")

	_Log("Step completed successfully.")
	Return True
EndFunc   ;==>_DisableDPIScaling

;===============================================================================
; Step 6: Unmount and Commit
;===============================================================================
Func _UnmountCommit()
	_Log("Unmounting and committing changes")
	_Log("NOTE: Make sure no files are open in the mount path!")

	Local $cmd = 'Dism /Unmount-Image /MountDir:"' & $WIMMountPath & '" /commit'
	_RunCmd($cmd, "Unmounting WIM")

	If @error Then Return False

	; Remove mount folder
	DirRemove($WIMMountPath, 1)
	$BootWIMMounted = False

	Return True
EndFunc   ;==>_UnmountCommit

;===============================================================================
; Step 7: Trim Images
;===============================================================================
Func _TrimBootWIM()
	_Log("Trimming boot.wim")

	; Construct source index parameter
	Local $SourceIndexParam = StringReplace(_ConstructIndexParam($BootWIMIndex), "/Index:", "/SourceIndex:")
	$SourceIndexParam = StringReplace($SourceIndexParam, "/Name:", "/SourceName:")

	Local $cmd = 'Dism /Export-Image /SourceImageFile:"' & $BootWIMPath & '" ' & $SourceIndexParam & ' /DestinationImageFile:"' & $ISOTempPath & '\sources\boot2.wim" /Compress:Max'
	_RunCmd($cmd, "Exporting image")

	If @error Then Return False

	; Replace original with trimmed version
	_Log("Replacing boot.wim with trimmed version...")
	FileDelete($BootWIMPath)
	FileMove($ISOTempPath & "\sources\boot2.wim", $BootWIMPath)

	_Log("Step completed successfully.")
	Return True
EndFunc   ;==>_TrimImages

;===============================================================================
; Step 8: Remove Installer
;===============================================================================
Func _RemoveInstallWIM()
	_Log("Removing Install.wim and related")

	; Remove support folder
	_Log("Removing support folder...")
	DirRemove($ISOTempPath & "\support", 1)

	; Remove setup files
	_Log("Removing setup.exe and autorun.inf...")
	FileDelete($ISOTempPath & "\setup.exe")
	FileDelete($ISOTempPath & "\autorun.inf")

	; Move boot.wim, remove sources, recreate and move back
	_Log("Cleaning sources folder...")
	FileMove($ISOTempPath & "\sources\boot.wim", $ISOTempPath & "\boot.wim", 1)
	DirRemove($ISOTempPath & "\sources", 1)
	DirCreate($ISOTempPath & "\sources")
	FileMove($ISOTempPath & "\boot.wim", $ISOTempPath & "\sources\boot.wim", 1)

	_Log("Step completed successfully.")
	Return True
EndFunc   ;==>_RemoveInstallWIM

;===============================================================================
; Step 9: Make ISO
;===============================================================================
Func _MakeISO()
	_Log("Creating ISO")

	Local $oscdimgPath = $ADKPath & "\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"

	If Not FileExists($oscdimgPath) Then
		_Log("Error: oscdimg.exe not found at: " & $oscdimgPath)
		Return False
	EndIf

	_Log("Input: " & $ISOTempPath)
	_Log("Output: " & $OutputISOPath)

	; Determine boot data based on BIOS boot files
	Local $bootData
	If FileExists($ISOTempPath & "\boot\etfsboot.com") Then
		; BIOS + UEFI
		$bootData = '2#p0,e,b"' & $ISOTempPath & '\boot\etfsboot.com"#pEF,e,b"' & $ISOTempPath & '\efi\microsoft\boot\efisys.bin"'
	Else
		; UEFI only
		$bootData = '1#pEF,e,b"' & $ISOTempPath & '\efi\microsoft\boot\efisys.bin"'
	EndIf

	; Delete existing output ISO
	If FileExists($OutputISOPath) Then
		_Log("Removing existing output ISO...")
		FileDelete($OutputISOPath)
	EndIf

	Local $cmd = '"' & $oscdimgPath & '" -bootdata:' & $bootData & ' -u1 -udfver102 "' & $ISOTempPath & '" "' & $OutputISOPath & '"'
	_RunCmd($cmd, "Creating ISO")

	If @error Then Return False
	Return True
EndFunc   ;==>_MakeISO

;===============================================================================
; Tool: Browse Mount Folder
;===============================================================================
Func _BrowseMountFolder()
	If FileExists($WIMMountPath) Then
		ShellExecute("explorer.exe", $WIMMountPath)
	Else
		MsgBox(48, "Browse Mount", "Mount folder does not exist:" & @CRLF & $WIMMountPath)
	EndIf
EndFunc   ;==>_BrowseMountFolder

;===============================================================================
; Tool: Unmount and Discard
;===============================================================================
Func _UnmountDiscard()
	_Log("Unmounting and discarding changes")

	_RunCmd('Dism /Unmount-Image /MountDir:"' & $WIMMountPath & '" /Discard', "Discarding changes")
	_RunCmd('Dism /Cleanup-Mountpoints', "Cleaning up mount points")
	; DISM /Cleanup-Wim

	DirRemove($WIMMountPath, 1)
	$BootWIMMounted = False

	_Log("Done.")
EndFunc   ;==>_UnmountDiscard

;===============================================================================
; Tool: Get Image Info
;===============================================================================
Func _GetImageInfo()
	_ReadGUI() ; Ensure $BootWIMPath and other globals are synchronized with the GUI

	_RunCmd('Dism /Get-MountedImageInfo', "Mounted Images")

	If FileExists($BootWIMPath) Then
		_RunCmd('Dism /Get-ImageInfo /imagefile:"' & $BootWIMPath & '"', "boot.wim Info")
	Else
		_Log("boot.wim not found at: " & $BootWIMPath)
	EndIf
EndFunc   ;==>_GetImageInfo

;===============================================================================
; Exit Handler
;===============================================================================
Func _Exit()
    ; If a WIM is mounted
    If $BootWIMMounted Then
        _Log("Exiting: WIM still mounted?")
    EndIf

    _Log("Script exited.")
EndFunc   ;==>_Exit
