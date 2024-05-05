; This script starts a VNC server and updates the status bar in Windows Setup Helper GUI

#include <WinAPIProc.au3>

Global $Title = "PEVNCServer"
Global $Settings = @ScriptDir & "\Settings.ini"
Global $VNCRegPath = "HKEY_CURRENT_USER\SOFTWARE\TightVNC\Server"
Global $VNCExe = "tvnserver.exe"
Global $IsPE = StringInStr(@WindowsDir, "X:")
If Not $IsPE Then Exit
FileChangeDir(@ScriptDir)

_Log($Title)
_Log("@WorkingDir="&@WorkingDir)
_Log("@ScriptDir="&@ScriptDir)

OnAutoItExitRegister("_Exit")

; Starting the VNC server without network connectivity can cause it to fail
For $i=1 to 6
	Ping ("8.8.8.8", 1000)
	If Not @error Then ExitLoop
	Sleep(1000)
Next

; Disable the PE firewall
Run(@ComSpec & " /c " & 'wpeutil.exe disablefirewall', "", @SW_HIDE)

; Import VNC settings to registry
Run(@ComSpec & " /c " & 'reg.exe import vnc_settings.reg', "", @SW_HIDE)

; Get Settings from INI
$SetPort = IniRead($Settings, "Settings", "port", "5900" )
Global $SetPass = IniRead($Settings, "Settings", "password", "vncwatch" )

; Convert password to VNC hex and write to registry
$Run = Run(@ComSpec & " /c " & 'vncpassword\vncpassword.exe ' & $SetPass, "", @SW_SHOW, $STDERR_CHILD + $STDOUT_CHILD)
_Log("Run: " & $Run & " " & @error)

Local $Read
While 1
    $Read &= StdoutRead($Run)
    If @error Then ExitLoop
Wend
$Read = StringStripWS($Read, 1 + 2)
If StringLen($Read) <> 18 Or StringLeft($Read, 2) <> "0x" Then
	_Log("Error: Expected hex")
	Exit
EndIf

RegWrite($VNCRegPath, "Password", "REG_BINARY", $Read)

; Start the VNC server
$VNCPid = Run(@ComSpec & " /c " & $VNCExe & ' -run', "", @SW_HIDE)
_Log("$VNCPid=" & $VNCPid)

; Setup statusbar updates
_UpdateStatusBar()
AdlibRegister("_UpdateStatusBar", 3000)

; Get PID of parent process
$ParentPid = _ProcessGetParent(@AutoItPID)
_Log("$ParentPid=" & $ParentPid)

; Loop until vnc process closes
While ProcessExists($VNCPid)
	; Exit if not in PE and parent process exits
	If Not $IsPE And Not ProcessExists($ParentPid) Then
		_Log("Parent closed")
		Exit
	EndIf

	Sleep(50)
Wend

Exit

;=========== =========== =========== =========== =========== =========== =========== ===========
;=========== =========== =========== =========== =========== =========== =========== ===========

; Close any instances of VNC server running from script path
Func _Exit()
	If Not $IsPE Then
		$aProcList = ProcessList($VNCExe)
		For $i = 1 To Ubound($aProcList) - 1
			If StringInStr(_WinAPI_GetProcessFileName($aProcList[$i][1]), @ScriptDir) Then ProcessClose($aProcList[$i][1])
		Next
	EndIf
EndFunc

Func _UpdateStatusBar()
	Local $Path = @TempDir & "\Helper_Status_" & $Title & ".txt"
	Local $Port = RegRead($VNCRegPath, "rfbport")

	$hFile = FileOpen ($Path, 2)
	FileWrite($hFile, "VNC Running" & @CRLF & "VNC Port: " & $Port & " VNC Password: " & $SetPass)

	FileClose($hFile)
EndFunc

Func _Log($Data)
	ConsoleWrite($Title & ": " & $Data & @CRLF)
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