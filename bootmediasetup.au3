#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <AutoItConstants.au3>

$Title = "IT Setup Helper - Boot Media Loader"

$Parameters = "/AutoIt3ExecuteScript Main.au3 bootmedia"

Sleep(2000)

For $i = 65 To 90
	$Path = Chr($i) & ":\Windows 10 Images\20H2\sources\$OEM$\$$\IT\AutoIt3.exe"
	If FileExists($Path) Then
		ConsoleWrite("Found: "&$Path & ' ' & $Parameters&@CRLF)
		$Return = Run(@ComSpec & " /c " & $Path & ' ' & $Parameters, "", @SW_HIDE)
		ConsoleWrite("Run: "&$Return&@CRLF)
		Sleep(2000)
	Endif
Next

;Run(@ComSpec & " /c " & 'x:\setup.exe '&$CmdLineRaw, "", @SW_HIDE, $STDIO_INHERIT_PARENT)
