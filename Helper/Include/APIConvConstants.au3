#include-once

; #INDEX# =======================================================================================================================
; Title .........: WinAPICom Constants UDF Library for AutoIt3
; AutoIt Version : 3.3.18.0
; Language ......: English
; Description ...: Constants that can be used with UDF library
; Author(s) .....: Yashied, Jpm
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================

; code page identifiers
; _WinAPI_MultiByteToWideChar(), _WinAPI_MultiByteToWideCharEx() and _WinAPI_WideCharToMultiByte()
Global Const $CP_ACP = 0
Global Const $CP_OEMCP = 1
Global Const $CP_MACCP = 2
Global Const $CP_THREAD_ACP = 3
Global Const $CP_SYMBOL = 42
Global Const $CP_SHIFT_JIS = 932
Global Const $CP_UTF16 = 1200
Global Const $CP_UNICODE = $CP_UTF16
Global Const $CP_UTF7 = 65000
Global Const $CP_UTF8 = 65001

; conversion type
; _WinAPI_MultiByteToWideChar() and _WinAPI_MultiByteToWideCharEx()
Global Const $MB_PRECOMPOSED = 0x01
Global Const $MB_COMPOSITE = 0x02
Global Const $MB_USEGLYPHCHARS = 0x04
Global Const $MB_ERR_INVALID_CHARS = 0x08
; ===============================================================================================================================
