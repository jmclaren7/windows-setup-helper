; This script adds file associations for common file types so the can be easily opened

Global $IsPE = StringInStr(@SystemDir, "X:")
If Not $IsPE Then Exit

$HelperHome = StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 0, -1) - 1)

_FileRegister("au3", '"' & @AutoItExe & '" "%1" %*', "Run With AutoIt", True, @AutoItExe & ",-99", "AutoIt Script")
_FileRegister("au3", '"' & @WindowsDir & '\notepad.exe" "%1"', "Edit")
_FileRegister("a3x", '"' & @AutoItExe & '" "%1" %*', "Run With AutoIt", True, @AutoItExe & ",-99", "AutoIt Script")

_FileRegister("txt", '"' & @WindowsDir & '\notepad.exe" "%1"', "Open with Notepad", True, @WindowsDir & "\System32\shell32.dll,-152", "Text Document")
_FileRegister("log", '"' & @WindowsDir & '\notepad.exe" "%1"', "Open with Notepad", True, @WindowsDir & "\System32\shell32.dll,-152", "Text Document")
_FileRegister("ini", '"' & @WindowsDir & '\notepad.exe" "%1"', "Open with Notepad", True, @WindowsDir & "\System32\shell32.dll,-151", "Text Document")
_FileRegister("xml", '"' & @WindowsDir & '\notepad.exe" "%1"', "Open with Notepad", True, @WindowsDir & "\System32\shell32.dll,-243", "Text Document")

_FileRegister("bmp", '"' & $HelperHome & '\Tools\ReactOS Paint.exe" "%1"', "Open with ReactOS Paint", True, @WindowsDir & "\System32\shell32.dll,-16823", "Image File")
_FileRegister("jpg", '"' & $HelperHome & '\Tools\ReactOS Paint.exe" "%1"', "Open with ReactOS Paint", True, @WindowsDir & "\System32\shell32.dll,-16823", "Image File")
_FileRegister("png", '"' & $HelperHome & '\Tools\ReactOS Paint.exe" "%1"', "Open with ReactOS Paint", True, @WindowsDir & "\System32\shell32.dll,-16823", "Image File")
_FileRegister("gif", '"' & $HelperHome & '\Tools\ReactOS Paint.exe" "%1"', "Open with ReactOS Paint", True, @WindowsDir & "\System32\shell32.dll,-16823", "Image File")

_FileRegister("zip", '"' & $HelperHome & '\Tools\7-Zip\7zFM.exe" "%1"', "Open with 7-Zip", True, $HelperHome & "\Tools\7-Zip\7z.dll,-1", "7-Zip File")
_FileRegister("7z", '"' & $HelperHome & '\Tools\7-Zip\7zFM.exe" "%1"', "Open with 7-Zip", True, $HelperHome & "\Tools\7-Zip\7z.dll,0", "7-Zip File")
_FileRegister("iso", '"' & $HelperHome & '\Tools\7-Zip\7zFM.exe" "%1"', "Open with 7-Zip", True, $HelperHome & "\Tools\7-Zip\7z.dll,-8", "7-Zip File")
_FileRegister("wim", '"' & $HelperHome & '\Tools\7-Zip\7zFM.exe" "%1"', "Open with 7-Zip", True, $HelperHome & "\Tools\7-Zip\7z.dll,-15", "7-Zip File")
_FileRegister("esd", '"' & $HelperHome & '\Tools\7-Zip\7zFM.exe" "%1"', "Open with 7-Zip", True, $HelperHome & "\Tools\7-Zip\7z.dll,-15", "7-Zip File")
_FileRegister("vhd", '"' & $HelperHome & '\Tools\7-Zip\7zFM.exe" "%1"', "Open with 7-Zip", True, $HelperHome & "\Tools\7-Zip\7z.dll,-15", "7-Zip File")
_FileRegister("vhdx", '"' & $HelperHome & '\Tools\7-Zip\7zFM.exe" "%1"', "Open with 7-Zip", True, $HelperHome & "\Tools\7-Zip\7z.dll,-15", "7-Zip File")

;==============================================================================================
; Description:		_FileRegister($FileExt, $Command, $Verb[, $Default = Default[, $Icon = Default[, $Description = Default]]])
;					Registers a file type in Explorer
; Parameter(s):		$FileExt - 	File Extension without period eg. "zip"
;					$Command - 	Program path with arguments eg. '"C:\test\testprog.exe" "%1"'
;								(%1 is 1st argument, %2 is 2nd, etc.)
;								Setting $Command to an empty string "" will skip setting that key and return the value of an existing key
;					$Verb 	- 	Name of action to perform on file
;								eg. "Open with ProgramName" or "Extract Files"
;					$Default - 	(True/False) The verb will be the default for this filetype
;								If the file is not already associated, this will be the default.
;								Setting default to False will return the current default verb
;					$Icon - 	Default icon for filetype including resource # if needed
;								eg. "C:\test\testprog.exe,0" or "C:\test\filetype.ico"
;					$Description - File Description eg. "Zip File" or "ProgramName Document"
; Returns:  		Returns the new verb is setting that key was a success, the old verb if it was not
;					Sets @extended to the new command if setting that key was a success or the old command if not
; Notes:
; Author:
; Date/Version:   	10/15/2014  --  v2.0.4
;===============================================================================================
Func _FileRegister($FileExt, $Command, $Verb = Default, $Default = Default, $Icon = Default, $Description = Default)
	; FileExt is a key representing the file extention and specifying which FileType to use for that extention
	; FileType is a key containing various properties and actions (verbs) that can be used for files of this type

	If $Default = Default Then $Default = False ; Do not make the new verb the default
	If $Command = Default Or $Command = "" Then $Command = "" ; Allow use of Default to be treated as ""

	; Remove the "." if it was included
	If StringLeft($FileExt, 1) = "." Then $FileExt = StringTrimLeft($FileExt, 1)

	; Get the current FileType for the extention
	Local $FileType = RegRead("HKCR\." & $FileExt, "")
	If @error And $Verb = Default Then
		; If FileExt doesn't exist but a verb wasn't specified then we have nothing to do.
		Return SetError(1, 0, "")
	ElseIf @error Then ; The extention doesn't exist so create it
		RegWrite("HKCR\." & $FileExt, "", "REG_SZ", $FileExt & "file") ; Create a new FileType to use with it
		$FileType = $FileExt & "file" ; Make the new Verb default since it doesn't have one
		$Default = True
	EndIf

	; Verb keys can't have spaces, still use $Verb for a display name
	$VerbKey = StringReplace($Verb, " ", "")

	; Set the default verb to use
	Local $CurrentDefaultVerb = RegRead("HKCR\" & $FileType & "\shell", "")
	If $Default Then
		If Not @error Then RegWrite("HKCR\" & $FileType & "\shell", "oldverb", "REG_SZ", $CurrentDefaultVerb) ; Backup default verb
		RegWrite("HKCR\" & $FileType & "\shell", "", "REG_SZ", $VerbKey) ; Make new verb the default
		If Not @error Then $CurrentDefaultVerb = $VerbKey
	EndIf

	If $Verb <> Default Then
		; The display name for the verb
		Local $CurrentVerbName = RegRead("HKCR\" & $FileType & "\shell\" & $VerbKey, "")
		If Not @error Then RegWrite("HKCR\" & $FileType & "\shell\" & $VerbKey, "oldname", "REG_SZ", $CurrentVerbName) ; Backup verb name
		RegWrite("HKCR\" & $FileType & "\shell\" & $VerbKey, "", "REG_SZ", $Verb) ; Set the display name

		; Command is a subkey reprenseting what happens when a particular verb is selected
		Local $CurrentCommand = RegRead("HKCR\" & $FileType & "\shell\" & $VerbKey & "\command", "")
		If $Command <> Default Then
			If Not @error Then RegWrite("HKCR\" & $FileType & "\shell\" & $VerbKey & "\command", "oldcmd", "REG_SZ", $CurrentCommand) ; Backup command
			RegWrite("HKCR\" & $FileType & "\shell\" & $VerbKey & "\command", "", "REG_SZ", $Command) ; Set the command
			If Not @error Then $CurrentCommand = $Command
		EndIf
	EndIf

	; Specify the icon to be used for the FileType
	If $Icon <> Default Then
		Local $CurrentIcon = RegRead("HKCR\" & $FileType & "\DefaultIcon", "")
		If Not @error Then RegWrite("HKCR\" & $FileType & "\DefaultIcon", "oldicon", "REG_SZ", $CurrentIcon) ; Backup icon
		RegWrite("HKCR\" & $FileType & "\DefaultIcon", "", "REG_SZ", $Icon)
	EndIf

	; Set the description for the the file type
	If $Description <> Default Then
		Local $CurrentDescription = RegRead("HKCR\" & $FileType, "")
		If @error Then RegWrite("HKCR\" & $FileType, "olddesc", "REG_SZ", $CurrentDescription) ; Backup description
		RegWrite("HKCR\" & $FileType, "", "REG_SZ", $Description) ; Write the description
	EndIf

	Return SetError(0, $CurrentCommand, $CurrentDefaultVerb)
EndFunc   ;==>_FileRegister
