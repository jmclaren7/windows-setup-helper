; This script adds file associations for common file types so the can be easily opened

#Include "CommonFunctions.au3"

Global $IsPE = StringInStr(@SystemDir, "X:")
If Not $IsPE Then Exit

$HelperHome = StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 0, -1) - 1)

_FileRegister("au3", '"' & @AutoItExe & '" "%1" %*', "Run With AutoIt", True, @AutoItExe & ",-99", "AutoIt Script")
_FileRegister("au3", '"' & @WindowsDir & '\notepad.exe" "%1"', "Edit")
_FileRegister("a3x", '"' & @AutoItExe & '" "%1" %*', "Run With AutoIt", True, @AutoItExe & ",-99", "AutoIt Script")

_FileRegister("ps1", '"' & @WindowsDir & '\notepad.exe" "%1"', "Edit", True)

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
