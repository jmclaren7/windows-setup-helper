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
Global $MountPath = @TempDir & "\WSHelper-WIMMount"
Global $HelperRepo = @ScriptDir
Global $DefaultADKPackages = "WinPE-WMI.cab|WinPE-NetFx.cab|WinPE-Scripting.cab|WinPE-PowerShell.cab|WinPE-StorageWMI.cab|WinPE-SecureBootCmdlets.cab|WinPE-SecureStartup.cab|WinPE-DismCmdlets.cab|WinPE-EnhancedStorage.cab|WinPE-Dot3Svc.cab|WinPE-FMAPI.cab|WinPE-FontSupport-WinRE.cab|WinPE-PlatformId.cab|WinPE-WDS-Tools.cab|WinPE-HTA.cab|WinPE-WinReCfg.cab"
Global $GUIMain
Global $IsRunning = False
Global $ProgramActive = True
Global $BootWIMMounted = False

;===============================================================================
; Start Program
;===============================================================================


; !!!!!!  This build script is experimental !!!!!!!!
MsgBox(48, "Warning - " & $TitleFull, "This build script is experimental and incomplete" & @CRLF & @CRLF & "Use with caution and verify all specified paths are safe to work with.")

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
Sleep(1000)
_CreateGUI()
GUISetState(@SW_SHOW, $GUIMain)

; Periodic GUI updates
Global $AdlibTimer = 1000
_ReadGUI()
_GUIChecks()


While 1
	Local $nMsg = GUIGetMsg()
    Switch $nMsg 
        Case 0, -11, -7, -9, -4
            ; Do nothing
        Case Else
            ;_Log("GUI Message: " & $nMsg)
            _ReadGUI()
            _GUIChecks()
    EndSwitch

	Switch $nMsg
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

		Case $mTempPath["Button"]
			Local $folder = FileSelectFolder("Select Temp Folder", "", 0, GUICtrlRead($mTempPath["Input"]), $GUIMain)
			If Not @error Then 
                GUICtrlSetData($mTempPath["Input"], $folder)
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
			_SetAllPackages(True)

		Case $btnPkgSelectNone
			_SetAllPackages(False)

        Case $btnPkgSelectDefault
            _UpdateADKPackages()

	EndSwitch

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
	Local $labelWidth = 80
	Local $inputWidth = 480
    Local $btnText = "Browse..."
	Local $btnWidth = 75

	; === Path Inputs Group ===
	GUICtrlCreateGroup("Paths", 10, $yPos, $guiWidth - 20, 260)
	$yPos += 20
    Global $mSourceISOPath = _CreateInputRow("Source ISO:", $xPos, $yPos, $inputWidth, $SourceISOPath, "Path to the Windows ISO file to customize", $btnText)
    $yPos += 30
    Global $mTempPath = _CreateInputRow("Temp Folder:", $xPos, $yPos, $inputWidth, $TempPath, "Temporary folder where the ISO will be extracted and modified", $btnText)
    $yPos += 30
    Global $mBootWIMPath = _CreateInputRow("Boot.wim:", $xPos, $yPos, $inputWidth, $BootWIMPath, "Path to boot.wim file to mount and modify (usually in Temp\sources\)", $btnText)
    $yPos += 30
    Global $mBootWIMIndex = _CreateInputRow("Boot.wim Index:", $xPos, $yPos, $inputWidth, $BootWIMIndex, "Image index or name to mount from boot.wim (e.g., 2 or 'Microsoft Windows Setup (amd64)')", "")
    $yPos += 30
    Global $mADKPath = _CreateInputRow("ADK Path:", $xPos, $yPos, $inputWidth, $ADKPath, "Path to Windows Assessment and Deployment Kit (ADK) installation", $btnText)
    $yPos += 30
    Global $mAddBootFilesPath = _CreateInputRow("Add to PE:", $xPos, $yPos, $inputWidth, $AddBootFilesPath, "Folder containing additional files to copy into the boot.wim image", $btnText)
    $yPos += 30
    Global $mAddISOFilesPath = _CreateInputRow("Add to ISO:", $xPos, $yPos, $inputWidth, $AddISOFilesPath, "Folder containing additional files to copy into the ISO root", $btnText)
    $yPos += 30
    Global $mOutputISOPath = _CreateInputRow("Output ISO:", $xPos, $yPos, $inputWidth, $OutputISOPath, "Path where the customized ISO will be saved", $btnText)
	GUICtrlCreateGroup("", -99, -99, 1, 1) ; Close group
    $yPos += 40

	; === Build Steps Group ===
	Local $stepsGroupHeight = 295
	GUICtrlCreateGroup("Build Steps", 10, $yPos, 520, $stepsGroupHeight)
	Local $stepsStartY = $yPos + 20

	; Checkboxes - Single column with Go buttons
	Local $chkX = 20, $chkWidth = 200, $goX = 220, $goWidth = 35
	Local $stepY = $stepsStartY
	Local $stepHeight = 25

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
    $stepY += $stepHeight + 10

	; Select All / Select None buttons
	Global $btnSelectAll = GUICtrlCreateButton("All", $chkX, $stepY, 40, 25)
	GUICtrlSetTip(-1, "Check all build steps")
	Global $btnSelectNone = GUICtrlCreateButton("None", $chkX + 45, $stepY, 50, 25)
	GUICtrlSetTip(-1, "Uncheck all build steps")

	; Run button
	Global $btnRun = GUICtrlCreateButton("Run Selected", $goX - 110 + 35, $stepY, 110, 25)
	GUICtrlSetTip(-1, "Execute all checked build steps in order")
	GUICtrlCreateGroup("", -99, -99, 1, 1) ; Close group

	; === ADK Packages Group (middle, between steps and tools) ===
	Local $pkgGroupX = 270
	Local $pkgGroupWidth = 260
	GUICtrlCreateGroup("WinPE Packages", $pkgGroupX, $yPos, $pkgGroupWidth, $stepsGroupHeight)

	; Create ListView with checkboxes
	Global $lvPackages = GUICtrlCreateListView("Package Name", $pkgGroupX + 10, $yPos + 20, $pkgGroupWidth - 20, $stepsGroupHeight - 70, BitOR($LVS_REPORT, $LVS_SINGLESEL, $LVS_SHOWSELALWAYS))
	_GUICtrlListView_SetExtendedListViewStyle($lvPackages, BitOR($LVS_EX_CHECKBOXES, $LVS_EX_FULLROWSELECT))
	_GUICtrlListView_SetColumnWidth($lvPackages, 0, $pkgGroupWidth - 50)
	GUICtrlSetTip($lvPackages, "Select WinPE optional components to add to the boot image (Step 4)")

	; Buttons for package selection
	Local $pkgBtnY = $yPos + $stepsGroupHeight - 40
	Global $btnPkgSelectAll = GUICtrlCreateButton("All", $pkgGroupX + 10, $pkgBtnY, 40, 25)
	GUICtrlSetTip(-1, "Select all available packages")
	Global $btnPkgSelectNone = GUICtrlCreateButton("None", $pkgGroupX + 55, $pkgBtnY, 50, 25)
	GUICtrlSetTip(-1, "Deselect all packages")
    Global $btnPkgSelectDefault = GUICtrlCreateButton("Default", $pkgGroupX + 110, $pkgBtnY, 60, 25)
	GUICtrlSetTip(-1, "Reset to default recommended packages")

	GUICtrlCreateGroup("", -99, -99, 1, 1) ; Close group

	; === Tools Group (right side) ===
	Local $toolsGroupHeight = 240
	GUICtrlCreateGroup("Tools", 540, $yPos, 140, $toolsGroupHeight)
	Local $toolX = 550
	Local $toolY = $yPos + 25
	Local $toolBtnWidth = 120
	Local $toolBtnHeight = 28

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

	GUICtrlCreateGroup("", -99, -99, 1, 1) ; Close group

	; === Save/Exit Buttons (below Tools) ===
	Local $btnY = $yPos + $toolsGroupHeight + 15
	Global $btnSave = GUICtrlCreateButton("Save", $toolX - 10, $btnY, 65, $toolBtnHeight)
	GUICtrlSetTip(-1, "Save current settings to Build.ini")
	Global $btnExit = GUICtrlCreateButton("Exit", $toolX + 65, $btnY, 65, $toolBtnHeight)
	GUICtrlSetTip(-1, "Exit the program")

	; Populate the packages ListView
	_UpdateADKPackages()

	$yPos += $stepsGroupHeight + 10

EndFunc   ;==>_CreateGUI

;===============================================================================
; Create Input Row (Label, Left, Top, Width, DefaultValue, Tip)
;===============================================================================
Func _CreateInputRow($Label, $Left, $Top, $Width, $Value = "", $Tip= "", $ButtonText = "")
    Local $mControls[]

    $mControls["Label"] = GUICtrlCreateLabel($Label, $Left, $Top + 3, 100, 17)

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
        $mControls["Button"] = GUICtrlCreateButton($ButtonText, $Left + $ChkWidth, $Top - 2, $ButtonWidth, 22)
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

    Local $adkOCsubpath = "\Windows Preinstallation Environment\amd64\WinPE_OCs"
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

    _Log("Found " & $cabFiles[0] & " ADK packages")
EndFunc   ;==>_UpdateADKPackages

;===============================================================================
; Get selected packages from ListView
;===============================================================================
Func _GetSelectedPackages()
    Local $aSelected[1] = [0]

    For $i = 0 To _GUICtrlListView_GetItemCount($lvPackages) - 1
        If _GUICtrlListView_GetItemChecked($lvPackages, $i) Then
            Local $pkgName = _GUICtrlListView_GetItemText($lvPackages, $i)
            _ArrayAdd($aSelected, $pkgName)
            $aSelected[0] += 1
        EndIf
    Next

    Return $aSelected
EndFunc   ;==>_GetSelectedPackages

;===============================================================================
; Select/Deselect all packages in ListView
;===============================================================================
Func _SetAllPackages($bChecked)
    Local $itemCount = _GUICtrlListView_GetItemCount($lvPackages)
    For $i = 0 To $itemCount - 1
        _GUICtrlListView_SetItemChecked($lvPackages, $i, $bChecked)
    Next
EndFunc   ;==>_SetAllPackages


;===============================================================================
; Checks for GUI requirements (called by AdlibRegister)
;===============================================================================
Func _GUIChecks()
    ; Set or reset adlib timer to avoid overlapping calls
    AdlibRegister("_GUIChecks", $AdlibTimer)

    ; Get relevant environment states
    Local $Alert
    Global $MountExists = FileExists($MountPath & "\Windows")
    Global $BootWIMExists = FileExists($BootWIMPath)
    Global $SourceISOExists = FileExists($SourceISOPath)
    Global $ADKExists = FileExists($ADKPath & "\Windows Preinstallation Environment")
    Global $TempExists = FileExists($TempPath)
    Global $AddBootFilesExists = FileExists($AddBootFilesPath)
    Global $AddISOFilesExists = FileExists($AddISOFilesPath)

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
    $Alert = BitAND(GUICtrlGetState($mTempPath["Alert"]), $GUI_SHOW)
    If $TempExists And $Alert Then
        GUICtrlSetState($mTempPath["Alert"], $GUI_HIDE)
    ElseIf Not $TempExists And Not $Alert Then
        GUICtrlSetState($mTempPath["Alert"], $GUI_SHOW)
    EndIf

    ; BootWIMPath
    $Alert = BitAND(GUICtrlGetState($mBootWIMPath["Alert"]), $GUI_SHOW)
    If $BootWIMExists And $Alert Then
        GUICtrlSetState($mBootWIMPath["Alert"], $GUI_HIDE)
        GUICtrlSetState($ToolGetInfoButton, $GUI_ENABLE)
        GUICtrlSetState($mMountWIM["Button"], $GUI_ENABLE)
    ElseIf Not $BootWIMExists And Not $Alert Then
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
        _UpdateADKPackages()
    ElseIf Not $ADKExists And Not $Alert Then
        GUICtrlSetState($mADKPath["Alert"], $GUI_SHOW)
        GUICtrlSetState($mAddPackages["Button"], $GUI_DISABLE)
        GUICtrlSetState($lvPackages, $GUI_DISABLE)
    EndIf

    ; AddBootFilesPath
    $Alert = BitAND(GUICtrlGetState($mAddBootFilesPath["Alert"]), $GUI_SHOW)
    If $AddBootFilesExists And $Alert Then
        GUICtrlSetState($mAddBootFilesPath["Alert"], $GUI_HIDE)
    ElseIf Not $AddBootFilesExists And Not $Alert Then
        GUICtrlSetState($mAddBootFilesPath["Alert"], $GUI_SHOW)
    EndIf

    ; AddISOFilesPath
    $Alert = BitAND(GUICtrlGetState($mAddISOFilesPath["Alert"]), $GUI_SHOW)
    If $AddISOFilesExists And $Alert Then
        GUICtrlSetState($mAddISOFilesPath["Alert"], $GUI_HIDE)  
    ElseIf Not $AddISOFilesExists And Not $Alert Then
        GUICtrlSetState($mAddISOFilesPath["Alert"], $GUI_SHOW)
    EndIf

    ; Does mount path exist? (WIM probably mounted)
    $GUIEnabled = BitAND(GUICtrlGetState($ToolBrowseMountButton), $GUI_ENABLE)
    If $MountExists And Not $GUIEnabled Then
        GUICtrlSetState($ToolBrowseMountButton, $GUI_ENABLE)
        GUICtrlSetState($ToolDiscardUnmountButton, $GUI_ENABLE)
        ;GUICtrlSetState($mMountWIM["Button"], $GUI_DISABLE)
    Elseif Not $MountExists And $GUIEnabled Then
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
	Global $TempPath = GUICtrlRead($mTempPath["Input"])
	Global $BootWIMPath = GUICtrlRead($mBootWIMPath["Input"])
	Global $AddBootFilesPath = GUICtrlRead($mAddBootFilesPath["Input"])
	Global $AddISOFilesPath = GUICtrlRead($mAddISOFilesPath["Input"])
	Global $OutputISOPath = GUICtrlRead($mOutputISOPath["Input"])
	Global $ADKPath = GUICtrlRead($mADKPath["Input"])
	Global $BootWIMIndex = GUICtrlRead($mBootWIMIndex["Input"])
EndFunc   ;==>_ReadGUI

;===============================================================================
; Load Settings
;===============================================================================
Func _LoadSettings()
	Global $SourceISOPath = IniRead($ConfigFile, "Paths", "SourceISOPath", "Windows11.iso")
	Global $TempPath = IniRead($ConfigFile, "Paths", "TempPath", @ScriptDir & "\Temp")
    Global $BootWIMPath = IniRead($ConfigFile, "Paths", "BootWIMPath", $TempPath & "\sources\boot.wim")
	Global $BootWIMIndex = IniRead($ConfigFile, "Paths", "BootWIMIndex", "Microsoft Windows Setup (amd64)")

	Global $ADKPath = IniRead($ConfigFile, "Paths", "ADKPath", "")
    If $ADKPath = "" Then
        $ADKPath = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Kits\Installed Roots", "KitsRoot10")
        If Not @error Then
            $ADKPath = $ADKPath & "Assessment and Deployment Kit"
        Else
            $ADKPath = EnvGet("ProgramFiles(x86)") & "\Windows Kits\10\Assessment and Deployment Kit"
        EndIf
    EndIf

    Global $AddBootFilesPath = IniRead($ConfigFile, "Paths", "AddBootFilesPath", @ScriptDir & "\Additional Boot.wim Files")
	Global $AddISOFilesPath = IniRead($ConfigFile, "Paths", "AddISOFilesPath", @ScriptDir & "\Additional ISO Files")
	Global $OutputISOPath = IniRead($ConfigFile, "Paths", "OutputISOPath", "Windows11-Output.iso")

EndFunc   ;==>_LoadSettings

;===============================================================================
; Save Settings to INI
;===============================================================================
Func _SaveSettings()
	IniWrite($ConfigFile, "Paths", "SourceISOPath", GUICtrlRead($mSourceISOPath["Input"]))
	IniWrite($ConfigFile, "Paths", "TempPath", GUICtrlRead($mTempPath["Input"]))
	IniWrite($ConfigFile, "Paths", "AddBootFilesPath", GUICtrlRead($mAddBootFilesPath["Input"]))
	IniWrite($ConfigFile, "Paths", "AddISOFilesPath", GUICtrlRead($mAddISOFilesPath["Input"]))
	IniWrite($ConfigFile, "Paths", "OutputISOPath", GUICtrlRead($mOutputISOPath["Input"]))
	IniWrite($ConfigFile, "Paths", "ADKPath", GUICtrlRead($mADKPath["Input"]))
	IniWrite($ConfigFile, "Paths", "BootWIMIndex", GUICtrlRead($mBootWIMIndex["Input"]))
    IniWrite($ConfigFile, "Paths", "BootWIMPath", GUICtrlRead($mBootWIMPath["Input"]))
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
	Local $iPID = Run(@ComSpec & ' /c "' & $sCommand & '"', @ScriptDir, @SW_HIDE);, $STDERR_MERGED)
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

	ProcessWaitClose($iPID)
	Local $exitCode = @extended

    ; Re-enable GUI interaction
    GUISetState(@SW_ENABLE, $GUIMain)
    WinSetTrans($GUIMain, "", 255)

	; Log output
	If StringStripWS($sOutput, 3) <> "" Then
		_Log(StringStripWS($sOutput, 3))
	EndIf

	If $exitCode <> 0 Then
		_Log("Exit code: " & $exitCode)
		Return SetError($exitCode, $exitCode, $sOutput)
	EndIf

	_Log("Step completed successfully.")
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
		_Log("=== COMPLETED SUCCESSFULLY ===")
	Else
		_Log("=== FAILED ===")
	EndIf

	$IsRunning = False
EndFunc   ;==>_RunSelectedSteps

;===============================================================================
; Construct WIM Index Parameter
;===============================================================================
Func _ConstructIndexParam($Index)
    If IsNumber($Index) Then
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
	If FileExists($TempPath) Then
		_Log("Removing existing temp folder...")
		DirRemove($TempPath, 1)
	EndIf

	; Create temp folder
	DirCreate($TempPath)

	; Extract using 7-Zip
	Local $7zPath = $HelperRepo & "\Helper\Tools\7-Zip\7z.exe"
	If Not FileExists($7zPath) Then
		_Log("Error: 7-Zip not found at: " & $7zPath)
		Return False
	EndIf

	Local $cmd = '"' & $7zPath & '" x -y -o"' & $TempPath & '" "' & $SourceISOPath & '"'
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
	If Not FileExists($MountPath) Then
		DirCreate($MountPath)
	EndIf

	Local $cmd = 'Dism /Mount-image /ImageFile:"' & $BootWIMPath & '" ' & _ConstructIndexParam($BootWIMIndex) & ' /MountDir:"' & $MountPath & '" /Optimize'
	_RunCmd($cmd, "Mounting WIM")

	If @error Then Return False
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
	Local $helperDest = $MountPath & "\Helper"
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
	$result = DirCopy($HelperRepo & "\Windows", $MountPath & "\Windows", 1)
	If Not $result Then
		_Log("Error: Warning: Could not copy Windows folder (may not exist)")
	EndIf

	; Copy extra files from subdirectories
	;~ If FileExists($AddBootFilesPath) Then
	;~ 	_Log("Copying extra files from: " & $AddBootFilesPath)
	;~ 	Local $aFolders = _FileListToArray($AddBootFilesPath, "*", $FLTA_FOLDERS)
	;~ 	If Not @error Then
	;~ 		For $i = 1 To $aFolders[0]
	;~ 			Local $srcFolder = $AddBootFilesPath & "\" & $aFolders[$i]
	;~ 			_Log("  Copying: " & $aFolders[$i])
	;~ 			DirCopy($srcFolder, $MountPath, 1)
	;~ 		Next
	;~ 	EndIf
	;~ EndIf

	; Clean up logs and temp files
	_Log("Removing log files and temp files...")
	Local $aFiles = _FileListToArray($MountPath, "Auto-saved*.xml", $FLTA_FILES)
	If Not @error Then
		For $i = 1 To $aFiles[0]
			FileDelete($MountPath & "\" & $aFiles[$i])
		Next
	EndIf
	FileDelete($MountPath & "\*.log")
	FileDelete($MountPath & "\Helper\Logon\*.log")

	_Log("Step completed successfully.")
	Return True
EndFunc   ;==>_CopyFiles

;===============================================================================
; Step 4: Add Packages
;===============================================================================
Func _AddPackages()
	_Log("Adding packages to image")

	Local $adkPackages = $ADKPath & "\Windows Preinstallation Environment\amd64\WinPE_OCs"
    Local $langFolder = $adkPackages & "\en-us"

	If Not FileExists($adkPackages) Then
		_Log("Error: ADK packages not found at: " & $adkPackages)
		Return False
	EndIf

	_Log("Packages from: " & $adkPackages)

	; Build the DISM command with all packages
    Local $aSelectedPackages = _GetSelectedPackages()
    For $i = 1 To $aSelectedPackages[0]
        Local $pkgPath = $adkPackages & "\" & $aSelectedPackages[$i]
        Local $Command = 'Dism /Image:"' & $MountPath & '" /Add-Package /PackagePath:"' & $pkgPath & '"'
        ; If $adkPackages\en-us folder exists, add language-specific package, file name are suffixed with language code
        Local $langPkgPath = $langFolder & "\" & StringRegExpReplace($aSelectedPackages[$i], "\.cab$", "_en-us.cab")
        If FileExists($langPkgPath) Then
            $Command &= ' /PackagePath:"' & $langPkgPath & '"'
        EndIf

        _RunCmd($Command, "Adding package: " & $aSelectedPackages[$i])

        If @error Then
            _Log("Error: Error adding package: " & $aSelectedPackages[$i])
            Return False
        EndIf
    Next

	Return True
EndFunc   ;==>_AddPackages

;===============================================================================
; Step 5: Disable DPI Scaling
;===============================================================================
Func _DisableDPIScaling()
	_Log("Disabling DPI scaling (registry)")

	Local $regHive = $MountPath & "\Windows\System32\config\default"

	; Load registry hive
	_RunCmd('reg load HKLM\_WinPE_Default "' & $regHive & '"', "Loading registry hive")
	If @error Then Return False

	; Add registry keys
	_RunCmd('reg add "HKLM\_WinPE_Default\Control Panel\Desktop" /v LogPixels /t REG_DWORD /d 96 /f', "Setting LogPixels")
	_RunCmd('reg add "HKLM\_WinPE_Default\Control Panel\Desktop" /v Win8DpiScaling /t REG_DWORD /d 1 /f', "Setting Win8DpiScaling")
	_RunCmd('reg add "HKLM\_WinPE_Default\Control Panel\Desktop" /v DpiScalingVer /t REG_DWORD /d 4120 /f', "Setting DpiScalingVer")

	; Unload registry hive
	_RunCmd('reg unload HKLM\_WinPE_Default', "Unloading registry hive")

	_Log("Step completed successfully.")
	Return True
EndFunc   ;==>_DisableDPIScaling

;===============================================================================
; Step 6: Unmount and Commit
;===============================================================================
Func _UnmountCommit()
	_Log("Unmounting and committing changes")
	_Log("NOTE: Make sure no files are open in the mount path!")

	Local $cmd = 'Dism /Unmount-Image /MountDir:"' & $MountPath & '" /commit'
	_RunCmd($cmd, "Unmounting WIM")

	If @error Then Return False

	; Remove mount folder
	DirRemove($MountPath, 1)
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

	Local $cmd = 'Dism /Export-Image /SourceImageFile:"' & $BootWIMPath & '" ' & $SourceIndexParam & ' /DestinationImageFile:"' & $TempPath & '\sources\boot2.wim" /Compress:Max'
	_RunCmd($cmd, "Exporting image")

	If @error Then Return False

	; Replace original with trimmed version
	_Log("Replacing boot.wim with trimmed version...")
	FileDelete($BootWIMPath)
	FileMove($TempPath & "\sources\boot2.wim", $BootWIMPath)

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
	DirRemove($TempPath & "\support", 1)

	; Remove setup files
	_Log("Removing setup.exe and autorun.inf...")
	FileDelete($TempPath & "\setup.exe")
	FileDelete($TempPath & "\autorun.inf")

	; Move boot.wim, remove sources, recreate and move back
	_Log("Cleaning sources folder...")
	FileMove($TempPath & "\sources\boot.wim", $TempPath & "\boot.wim", 1)
	DirRemove($TempPath & "\sources", 1)
	DirCreate($TempPath & "\sources")
	FileMove($TempPath & "\boot.wim", $TempPath & "\sources\boot.wim", 1)

	_Log("Step completed successfully.")
	Return True
EndFunc   ;==>_RemoveInstallWIM

;===============================================================================
; Step 9: Make ISO
;===============================================================================
Func _MakeISO()
	_Log("Creating ISO")

	Local $oscdimgPath = $ADKPath & "\Deployment Tools\amd64\Oscdimg\oscdimg.exe"

	If Not FileExists($oscdimgPath) Then
		_Log("Error: oscdimg.exe not found at: " & $oscdimgPath)
		Return False
	EndIf

	_Log("Input: " & $TempPath)
	_Log("Output: " & $OutputISOPath)

	; Determine boot data based on BIOS boot files
	Local $bootData
	If FileExists($TempPath & "\boot\etfsboot.com") Then
		; BIOS + UEFI
		$bootData = '2#p0,e,b"' & $TempPath & '\boot\etfsboot.com"#pEF,e,b"' & $TempPath & '\efi\microsoft\boot\efisys.bin"'
	Else
		; UEFI only
		$bootData = '1#pEF,e,b"' & $TempPath & '\efi\microsoft\boot\efisys.bin"'
	EndIf

	; Delete existing output ISO
	If FileExists($OutputISOPath) Then
		_Log("Removing existing output ISO...")
		FileDelete($OutputISOPath)
	EndIf

	Local $cmd = '"' & $oscdimgPath & '" -bootdata:' & $bootData & ' -u1 -udfver102 "' & $TempPath & '" "' & $OutputISOPath & '"'
	_RunCmd($cmd, "Creating ISO")

	If @error Then Return False
	Return True
EndFunc   ;==>_MakeISO

;===============================================================================
; Tool: Browse Mount Folder
;===============================================================================
Func _BrowseMountFolder()
	If FileExists($MountPath) Then
		ShellExecute("explorer.exe", $MountPath)
	Else
		MsgBox(48, "Browse Mount", "Mount folder does not exist:" & @CRLF & $MountPath)
	EndIf
EndFunc   ;==>_BrowseMountFolder

;===============================================================================
; Tool: Unmount and Discard
;===============================================================================
Func _UnmountDiscard()
	_Log("Unmounting and discarding changes")

	_RunCmd('Dism /Unmount-Image /MountDir:"' & $MountPath & '" /Discard', "Discarding changes")
	_RunCmd('Dism /Cleanup-Mountpoints', "Cleaning up mount points")

	DirRemove($MountPath, 1)
	$BootWIMMounted = False

	_Log("Done.")
EndFunc   ;==>_UnmountDiscard

;===============================================================================
; Tool: Get Image Info
;===============================================================================
Func _GetImageInfo()
	_Log("Getting Image Information")

	$BootWIMPath = GUICtrlRead($mTempPath["Input"]) & "\sources\boot.wim"

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
    Exit
EndFunc   ;==>_Exit
