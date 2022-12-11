#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

_Log("Start")
TCPStartup()

$Host = "10.7.55.58"
$Port = 5950

$Return = TCPConnect ( $Host, $Port)
_Log("Error: " & @error)
_Log("Return: " & $Return)

_Log("End")

Sleep(3000)

Func _Log($data)
	ConsoleWrite(@CRLF & $data)
EndFunc