#include <WinAPIProc.au3>

_ConsoleWrite("VNCServer")
For $i=1 to 6
	Ping ("8.8.8.8", 1000)
	If Not @error Then ExitLoop
	Sleep(1000)
Next

FileChangeDir(@ScriptDir)
_ConsoleWrite("@WorkingDir="&@WorkingDir)
_ConsoleWrite("@ScriptDir="&@ScriptDir)

Run(@ComSpec & " /c " & 'wpeutil.exe disablefirewall', "", @SW_HIDE)

Run(@ComSpec & " /c " & 'reg.exe import tight_settings.reg', "", @SW_HIDE)
$VNCPid = Run(@ComSpec & " /c " & 'tvnserver.exe', "", @SW_HIDE)

_ConsoleWrite($VNCPid & @CRLF)

If NOT StringInStr(@ScriptFullPath, "x:\") Then
	$Parent = _ProcessGetParent(@AutoItPID)
	_ConsoleWrite($Parent & @CRLF)

	While ProcessExists($Parent)

		Sleep(50)
	Wend

	ProcessClose($VNCPid)
	ProcessClose("tvnserver.exe")

Else
	While ProcessExists($VNCPid)

		Sleep(50)
	Wend

Endif

Func _ConsoleWrite($Data)
	ConsoleWrite(@ScriptName & ": " & $Data & @CRLF)
EndFunc

Func _ProcessGetParent($i_pid)
    Local Const $TH32CS_SNAPPROCESS = 0x00000002

    Local $a_tool_help = DllCall("Kernel32.dll", "long", "CreateToolhelp32Snapshot", "int", $TH32CS_SNAPPROCESS, "int", 0)
    If IsArray($a_tool_help) = 0 Or $a_tool_help[0] = -1 Then Return SetError(1, 0, $i_pid)

    Local $tagPROCESSENTRY32 = _
        DllStructCreate _
            ( _
                "dword dwsize;" & _
                "dword cntUsage;" & _
                "dword th32ProcessID;" & _
                "uint th32DefaultHeapID;" & _
                "dword th32ModuleID;" & _
                "dword cntThreads;" & _
                "dword th32ParentProcessID;" & _
                "long pcPriClassBase;" & _
                "dword dwFlags;" & _
                "char szExeFile[260]" _
            )
    DllStructSetData($tagPROCESSENTRY32, 1, DllStructGetSize($tagPROCESSENTRY32))

    Local $p_PROCESSENTRY32 = DllStructGetPtr($tagPROCESSENTRY32)

    Local $a_pfirst = DllCall("Kernel32.dll", "int", "Process32First", "long", $a_tool_help[0], "ptr", $p_PROCESSENTRY32)
    If IsArray($a_pfirst) = 0 Then Return SetError(2, 0, $i_pid)

    Local $a_pnext, $i_return = 0
    If DllStructGetData($tagPROCESSENTRY32, "th32ProcessID") = $i_pid Then
        $i_return = DllStructGetData($tagPROCESSENTRY32, "th32ParentProcessID")
        DllCall("Kernel32.dll", "int", "CloseHandle", "long", $a_tool_help[0])
        If $i_return Then Return $i_return
        Return $i_pid
    EndIf

    While 1
        $a_pnext = DLLCall("Kernel32.dll", "int", "Process32Next", "long", $a_tool_help[0], "ptr", $p_PROCESSENTRY32)
        If IsArray($a_pnext) And $a_pnext[0] = 0 Then ExitLoop
        If DllStructGetData($tagPROCESSENTRY32, "th32ProcessID") = $i_pid Then
            $i_return = DllStructGetData($tagPROCESSENTRY32, "th32ParentProcessID")
            If $i_return Then ExitLoop
            $i_return = $i_pid
            ExitLoop
        EndIf
    WEnd

    If $i_return = "" Then $i_return = $i_pid

    DllCall("Kernel32.dll", "int", "CloseHandle", "long", $a_tool_help[0])
    Return $i_return
EndFunc