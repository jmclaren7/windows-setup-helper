
; For those who would fear the license - don't. I tried to license it as liberal as possible.
; It really means you can do what ever you want with this.
; Donations are wellcome And will be accepted via PayPal address: trancexx at yahoo dot com
; Thank you for the shiny stuff :kiss:

#comments-start
	Copyright 2013 Dragana R. <trancexx at yahoo dot com>

	Licensed under the Apache License, Version 2.0 (the "License");
	you may not use this file except in compliance with the License.
	You may obtain a copy of the License at

	http://www.apache.org/licenses/LICENSE-2.0

	Unless required by applicable law or agreed to in writing, software
	distributed under the License is distributed on an "AS IS" BASIS,
	WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
	See the License for the specific language governing permissions and
	limitations under the License.
#comments-end

#include-once
#include "WinHttpConstants.au3"

; #INDEX# ===================================================================================
; Title ...............: WinHttp
; File Name............: WinHttp.au3
; File Version.........: 1.6.4.1
; Min. AutoIt Version..: v3.3.7.20
; Description .........: AutoIt wrapper for WinHTTP functions
; Author... ...........: trancexx, ProgAndy
; Dll .................: winhttp.dll, kernel32.dll
; ===========================================================================================

; #CONSTANTS# ===============================================================================
Global Const $hWINHTTPDLL__WINHTTP = DllOpen("winhttp.dll")
DllOpen("winhttp.dll") ; making sure reference count never reaches 0
;============================================================================================

; #CURRENT# =================================================================================
;_WinHttpAddRequestHeaders
;_WinHttpCheckPlatform
;_WinHttpCloseHandle
;_WinHttpConnect
;_WinHttpCrackUrl
;_WinHttpCreateUrl
;_WinHttpDetectAutoProxyConfigUrl
;_WinHttpGetDefaultProxyConfiguration
;_WinHttpGetIEProxyConfigForCurrentUser
;_WinHttpOpen
;_WinHttpOpenRequest
;_WinHttpQueryAuthSchemes
;_WinHttpQueryDataAvailable
;_WinHttpQueryHeaders
;_WinHttpQueryOption
;_WinHttpReadData
;_WinHttpReceiveResponse
;_WinHttpSendRequest
;_WinHttpSetCredentials
;_WinHttpSetDefaultProxyConfiguration
;_WinHttpSetOption
;_WinHttpSetStatusCallback
;_WinHttpSetTimeouts
;_WinHttpSimpleBinaryConcat
;_WinHttpSimpleFormFill
;_WinHttpSimpleFormFill_SetUploadCallback
;_WinHttpSimpleReadData
;_WinHttpSimpleReadDataAsync
;_WinHttpSimpleRequest
;_WinHttpSimpleSendRequest
;_WinHttpSimpleSendSSLRequest
;_WinHttpSimpleSSLRequest
;_WinHttpTimeFromSystemTime
;_WinHttpTimeToSystemTime
;_WinHttpWriteData
; ===========================================================================================

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpAddRequestHeaders
; Description ...: Adds one or more HTTP request headers to the HTTP request handle.
; Syntax.........: _WinHttpAddRequestHeaders ($hRequest, $sHeaders [, $iModifiers = Default ])
; Parameters ....: $hRequest - Handle returned by _WinHttpOpenRequest function.
;                  $sHeader - [optional] Header(s) to append to the request.
;                  $iModifier - [optional] Contains the flags used to modify the semantics of this function. Default is $WINHTTP_ADDREQ_FLAG_ADD_IF_NEW.
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Remarks .......: In case of multiple additions at once use [[@CRLF]] to separate each [[$hRequest]] and responded [[$sHeaders]] and [[$iModifiers]].
; Related .......: _WinHttpOpenRequest, _WinHttpQueryHeaders
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384087.aspx
;============================================================================================
Func _WinHttpAddRequestHeaders($hRequest, $sHeader, $iModifier = Default)
	__WinHttpDefault($iModifier, $WINHTTP_ADDREQ_FLAG_ADD_IF_NEW)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpAddRequestHeaders", _
			"handle", $hRequest, _
			"wstr", $sHeader, _
			"dword", -1, _
			"dword", $iModifier)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpCheckPlatform
; Description ...: Determines whether the current platform is supported by this version of Microsoft Windows HTTP Services (WinHTTP).
; Syntax.........: _WinHttpCheckPlatform()
; Parameters ....: None
; Return values .: Success - Returns 1 if current platform is supported
;                          - Returns 0 if current platform is not supported
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384089.aspx
;============================================================================================
Func _WinHttpCheckPlatform()
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpCheckPlatform")
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpCloseHandle
; Description ...: Closes a single handle.
; Syntax.........: _WinHttpCloseHandle($hInternet)
; Parameters ....: $hInternet - Valid handle to be closed.
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Related .......: _WinHttpConnect, _WinHttpOpen, _WinHttpOpenRequest
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384090.aspx
;============================================================================================
Func _WinHttpCloseHandle($hInternet)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpCloseHandle", "handle", $hInternet)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpConnect
; Description ...: Specifies the initial target server of an HTTP request and returns connection handle to an HTTP session for that initial target.
; Syntax.........: _WinHttpConnect($hSession, $sServerName [, $iServerPort = Default ])
; Parameters ....: $hSession - Valid WinHttp session handle returned by a previous call to _WinHttpOpen().
;                  $sServerName - Host name of an HTTP server. In case URI scheme (http://, https://, ...) is specified $iServerPort is ignored.
;                  $iServerPort - [optional] TCP/IP port on the server to which a connection is made (default is $INTERNET_DEFAULT_PORT)
; Return values .: Success - Returns a valid connection handle to the HTTP session
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Remarks .......: [[$iServerPort]] can be defined via global constants [[$INTERNET_DEFAULT_PORT]], [[$INTERNET_DEFAULT_HTTP_PORT]] or [[$INTERNET_DEFAULT_HTTPS_PORT]]
; Related .......: _WinHttpOpen
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384091.aspx
;============================================================================================
Func _WinHttpConnect($hSession, $sServerName, $iServerPort = Default)
	Local $aURL = _WinHttpCrackUrl($sServerName), $iScheme = 0
	If @error Then
		__WinHttpDefault($iServerPort, $INTERNET_DEFAULT_PORT)
	Else
		$sServerName = $aURL[2]
		$iServerPort = $aURL[3]
		$iScheme = $aURL[1]
	EndIf
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "handle", "WinHttpConnect", _
			"handle", $hSession, _
			"wstr", $sServerName, _
			"dword", $iServerPort, _
			"dword", 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	_WinHttpSetOption($aCall[0], $WINHTTP_OPTION_CONTEXT_VALUE, $iScheme)
	Return $aCall[0]
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpCrackUrl
; Description ...: Separates a URL into its component parts such as host name and path.
; Syntax.........: _WinHttpCrackUrl($sURL [, $iFlag = Default ])
; Parameters ....: $sURL - String. Canonical URL to separate.
;                  $iFlag - [optional] Flag that control the operation. Default is $ICU_ESCAPE
; Return values .: Success - Returns array with 8 elements:
;                  |$array[0] - scheme name
;                  |$array[1] - internet protocol scheme
;                  |$array[2] - host name
;                  |$array[3] - port number
;                  |$array[4] - user name
;                  |$array[5] - password
;                  |$array[6] - URL path
;                  |$array[7] - extra information
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: ProgAndy
; Modified.......: trancexx
; Remarks .......: [[$iFlag]] is defined in WinHttpConstants.au3 and can be:
;                  |[[$ICU_DECODE]] - Converts characters that are "escape encoded" (%xx) to their non-escaped form.
;                  |[[$ICU_ESCAPE]] - Escapes certain characters to their escape sequences (%xx).
; Related .......: _WinHttpCreateUrl
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384092.aspx
;============================================================================================
Func _WinHttpCrackUrl($sURL, $iFlag = Default)
	__WinHttpDefault($iFlag, $ICU_ESCAPE)
	Local $tURL_COMPONENTS = DllStructCreate("dword StructSize;" & _
			"ptr SchemeName;" & _
			"dword SchemeNameLength;" & _
			"int Scheme;" & _
			"ptr HostName;" & _
			"dword HostNameLength;" & _
			"word Port;" & _
			"ptr UserName;" & _
			"dword UserNameLength;" & _
			"ptr Password;" & _
			"dword PasswordLength;" & _
			"ptr UrlPath;" & _
			"dword UrlPathLength;" & _
			"ptr ExtraInfo;" & _
			"dword ExtraInfoLength")
	DllStructSetData($tURL_COMPONENTS, 1, DllStructGetSize($tURL_COMPONENTS))
	Local $tBuffers[6]
	Local $iURLLen = StringLen($sURL)
	For $i = 0 To 5
		$tBuffers[$i] = DllStructCreate("wchar[" & $iURLLen + 1 & "]")
	Next
	DllStructSetData($tURL_COMPONENTS, "SchemeNameLength", $iURLLen)
	DllStructSetData($tURL_COMPONENTS, "SchemeName", DllStructGetPtr($tBuffers[0]))
	DllStructSetData($tURL_COMPONENTS, "HostNameLength", $iURLLen)
	DllStructSetData($tURL_COMPONENTS, "HostName", DllStructGetPtr($tBuffers[1]))
	DllStructSetData($tURL_COMPONENTS, "UserNameLength", $iURLLen)
	DllStructSetData($tURL_COMPONENTS, "UserName", DllStructGetPtr($tBuffers[2]))
	DllStructSetData($tURL_COMPONENTS, "PasswordLength", $iURLLen)
	DllStructSetData($tURL_COMPONENTS, "Password", DllStructGetPtr($tBuffers[3]))
	DllStructSetData($tURL_COMPONENTS, "UrlPathLength", $iURLLen)
	DllStructSetData($tURL_COMPONENTS, "UrlPath", DllStructGetPtr($tBuffers[4]))
	DllStructSetData($tURL_COMPONENTS, "ExtraInfoLength", $iURLLen)
	DllStructSetData($tURL_COMPONENTS, "ExtraInfo", DllStructGetPtr($tBuffers[5]))
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpCrackUrl", _
			"wstr", $sURL, _
			"dword", $iURLLen, _
			"dword", $iFlag, _
			"struct*", $tURL_COMPONENTS)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Local $aRet[8] = [DllStructGetData($tBuffers[0], 1), _
			DllStructGetData($tURL_COMPONENTS, "Scheme"), _
			DllStructGetData($tBuffers[1], 1), _
			DllStructGetData($tURL_COMPONENTS, "Port"), _
			DllStructGetData($tBuffers[2], 1), _
			DllStructGetData($tBuffers[3], 1), _
			DllStructGetData($tBuffers[4], 1), _
			DllStructGetData($tBuffers[5], 1)]
	Return $aRet
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpCreateUrl
; Description ...: Creates a URL from array of components such as the host name and path.
; Syntax.........: _WinHttpCreateUrl($aURLArray)
; Parameters ....: $aURLArray - Array of URL data.
; Return values .: Success - Returns created URL
;                  Failure - Returns empty string and sets @error:
;                  |1 - Invalid input.
;                  |2 - Initial DllCall failed
;                  |3 - Main DllCall failed
; Author ........: ProgAndy
; Modified.......: trancexx
; Remarks .......: Input is one dimensional 8 elements in size array:
;                  |- first element [0] scheme name
;                  |- second element [1] internet protocol scheme
;                  |- third element [2] host name
;                  |- fourth element [3] port number
;                  |- fifth element [4] user name
;                  |- sixth element [5] password
;                  |- seventh element [6] URL path
;                  |- eighth element [7] extra information
; Related .......: _WinHttpCrackUrl
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384093.aspx
;============================================================================================
Func _WinHttpCreateUrl($aURLArray)
	If UBound($aURLArray) - 8 Then Return SetError(1, 0, "")
	Local $tURL_COMPONENTS = DllStructCreate("dword StructSize;" & _
			"ptr SchemeName;" & _
			"dword SchemeNameLength;" & _
			"int Scheme;" & _
			"ptr HostName;" & _
			"dword HostNameLength;" & _
			"word Port;" & _
			"ptr UserName;" & _
			"dword UserNameLength;" & _
			"ptr Password;" & _
			"dword PasswordLength;" & _
			"ptr UrlPath;" & _
			"dword UrlPathLength;" & _
			"ptr ExtraInfo;" & _
			"dword ExtraInfoLength;")
	DllStructSetData($tURL_COMPONENTS, 1, DllStructGetSize($tURL_COMPONENTS))
	Local $tBuffers[6][2]
	$tBuffers[0][1] = StringLen($aURLArray[0])
	If $tBuffers[0][1] Then
		$tBuffers[0][0] = DllStructCreate("wchar[" & $tBuffers[0][1] + 1 & "]")
		DllStructSetData($tBuffers[0][0], 1, $aURLArray[0])
	EndIf
	$tBuffers[1][1] = StringLen($aURLArray[2])
	If $tBuffers[1][1] Then
		$tBuffers[1][0] = DllStructCreate("wchar[" & $tBuffers[1][1] + 1 & "]")
		DllStructSetData($tBuffers[1][0], 1, $aURLArray[2])
	EndIf
	$tBuffers[2][1] = StringLen($aURLArray[4])
	If $tBuffers[2][1] Then
		$tBuffers[2][0] = DllStructCreate("wchar[" & $tBuffers[2][1] + 1 & "]")
		DllStructSetData($tBuffers[2][0], 1, $aURLArray[4])
	EndIf
	$tBuffers[3][1] = StringLen($aURLArray[5])
	If $tBuffers[3][1] Then
		$tBuffers[3][0] = DllStructCreate("wchar[" & $tBuffers[3][1] + 1 & "]")
		DllStructSetData($tBuffers[3][0], 1, $aURLArray[5])
	EndIf
	$tBuffers[4][1] = StringLen($aURLArray[6])
	If $tBuffers[4][1] Then
		$tBuffers[4][0] = DllStructCreate("wchar[" & $tBuffers[4][1] + 1 & "]")
		DllStructSetData($tBuffers[4][0], 1, $aURLArray[6])
	EndIf
	$tBuffers[5][1] = StringLen($aURLArray[7])
	If $tBuffers[5][1] Then
		$tBuffers[5][0] = DllStructCreate("wchar[" & $tBuffers[5][1] + 1 & "]")
		DllStructSetData($tBuffers[5][0], 1, $aURLArray[7])
	EndIf
	DllStructSetData($tURL_COMPONENTS, "SchemeNameLength", $tBuffers[0][1])
	DllStructSetData($tURL_COMPONENTS, "SchemeName", DllStructGetPtr($tBuffers[0][0]))
	DllStructSetData($tURL_COMPONENTS, "HostNameLength", $tBuffers[1][1])
	DllStructSetData($tURL_COMPONENTS, "HostName", DllStructGetPtr($tBuffers[1][0]))
	DllStructSetData($tURL_COMPONENTS, "UserNameLength", $tBuffers[2][1])
	DllStructSetData($tURL_COMPONENTS, "UserName", DllStructGetPtr($tBuffers[2][0]))
	DllStructSetData($tURL_COMPONENTS, "PasswordLength", $tBuffers[3][1])
	DllStructSetData($tURL_COMPONENTS, "Password", DllStructGetPtr($tBuffers[3][0]))
	DllStructSetData($tURL_COMPONENTS, "UrlPathLength", $tBuffers[4][1])
	DllStructSetData($tURL_COMPONENTS, "UrlPath", DllStructGetPtr($tBuffers[4][0]))
	DllStructSetData($tURL_COMPONENTS, "ExtraInfoLength", $tBuffers[5][1])
	DllStructSetData($tURL_COMPONENTS, "ExtraInfo", DllStructGetPtr($tBuffers[5][0]))
	DllStructSetData($tURL_COMPONENTS, "Scheme", $aURLArray[1])
	DllStructSetData($tURL_COMPONENTS, "Port", $aURLArray[3])
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpCreateUrl", _
			"struct*", $tURL_COMPONENTS, _
			"dword", $ICU_ESCAPE, _
			"ptr", 0, _
			"dword*", 0)
	If @error Then Return SetError(2, 0, "")
	Local $iURLLen = $aCall[4]
	Local $URLBuffer = DllStructCreate("wchar[" & ($iURLLen + 1) & "]")
	$aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpCreateUrl", _
			"struct*", $tURL_COMPONENTS, _
			"dword", $ICU_ESCAPE, _
			"struct*", $URLBuffer, _
			"dword*", $iURLLen)
	If @error Or Not $aCall[0] Then Return SetError(3, 0, "")
	Return DllStructGetData($URLBuffer, 1)
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpDetectAutoProxyConfigUrl
; Description ...: Finds the URL for the Proxy Auto-Configuration (PAC) file.
; Syntax.........: _WinHttpDetectAutoProxyConfigUrl($iAutoDetectFlags)
; Parameters ....: $iAutoDetectFlags - Specifies what protocols to use to locate the PAC file.
; Return values .: Success - Returns URL for the PAC file.
;                  Failure - Returns empty string and sets @error:
;                  |1 - DllCall failed
;                  |2 - Internal failure.
; Author ........: trancexx
; Remarks .......: [[$iAutoDetectFlags]] values are defined in WinHttpconstants.au3
; Related .......: _WinHttpGetDefaultProxyConfiguration, _WinHttpGetIEProxyConfigForCurrentUser, _WinHttpSetDefaultProxyConfiguration
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384094.aspx
;============================================================================================
Func _WinHttpDetectAutoProxyConfigUrl($iAutoDetectFlags)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpDetectAutoProxyConfigUrl", "dword", $iAutoDetectFlags, "ptr*", 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, "")
	Local $pStr = $aCall[2]
	If $pStr Then
		Local $iLen = __WinHttpPtrStringLenW($pStr)
		If @error Then Return SetError(2, 0, "")
		Local $tString = DllStructCreate("wchar[" & $iLen + 1 & "]", $pStr)
		Local $sString = DllStructGetData($tString, 1)
		__WinHttpMemGlobalFree($pStr)
		Return $sString
	EndIf
	Return ""
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpGetDefaultProxyConfiguration
; Description ...: Retrieves the default WinHttp proxy configuration.
; Syntax.........: _WinHttpGetDefaultProxyConfiguration()
; Parameters ....: None.
; Return values .: Success - Returns array with 3 elements:
;                  |$array[0] - Integer. Access type.
;                  |$array[1] - String. Proxy server list.
;                  |$array[2] - String. Proxy bypass list.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Remarks .......: Access types are defined in WinHttpconstants.au3:
;                  |[[$WINHTTP_ACCESS_TYPE_DEFAULT_PROXY = 0]]
;                  |[[$WINHTTP_ACCESS_TYPE_NO_PROXY = 1]]
;                  |[[$WINHTTP_ACCESS_TYPE_NAMED_PROXY = 3]]
; Related .......: _WinHttpDetectAutoProxyConfigUrl, _WinHttpGetIEProxyConfigForCurrentUser, _WinHttpSetDefaultProxyConfiguration
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384095.aspx
;============================================================================================
Func _WinHttpGetDefaultProxyConfiguration()
	Local $tWINHTTP_PROXY_INFO = DllStructCreate("dword AccessType;" & _
			"ptr Proxy;" & _
			"ptr ProxyBypass")
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpGetDefaultProxyConfiguration", "struct*", $tWINHTTP_PROXY_INFO)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Local $iAccessType = DllStructGetData($tWINHTTP_PROXY_INFO, "AccessType")
	Local $pProxy = DllStructGetData($tWINHTTP_PROXY_INFO, "Proxy")
	Local $pProxyBypass = DllStructGetData($tWINHTTP_PROXY_INFO, "ProxyBypass")
	Local $sProxy
	If $pProxy Then
		Local $iProxyLen = __WinHttpPtrStringLenW($pProxy)
		If Not @error Then
			Local $tProxy = DllStructCreate("wchar[" & $iProxyLen + 1 & "]", $pProxy)
			$sProxy = DllStructGetData($tProxy, 1)
			__WinHttpMemGlobalFree($pProxy)
		EndIf
	EndIf
	Local $sProxyBypass
	If $pProxyBypass Then
		Local $iProxyBypassLen = __WinHttpPtrStringLenW($pProxyBypass)
		If Not @error Then
			Local $tProxyBypass = DllStructCreate("wchar[" & $iProxyBypassLen + 1 & "]", $pProxyBypass)
			$sProxyBypass = DllStructGetData($tProxyBypass, 1)
			__WinHttpMemGlobalFree($pProxyBypass)
		EndIf
	EndIf
	Local $aRet[3] = [$iAccessType, $sProxy, $sProxyBypass]
	Return $aRet
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpGetIEProxyConfigForCurrentUser
; Description ...: Retrieves the Internet Explorer proxy configuration for the current user.
; Syntax.........: _WinHttpGetIEProxyConfigForCurrentUser()
; Parameters ....: None.
; Return values .: Success - Returns array with 4 elements:
;                  |$array[0] - if 1 indicates that the IE proxy configuration for the current user specifies "automatically detect settings",
;                  |$array[1] - auto-configuration URL if the IE proxy configuration for the current user specifies "Use automatic proxy configuration",
;                  |$array[2] - proxy URL if the IE proxy configuration for the current user specifies "use a proxy server",
;                  |$array[3] - optional proxy by-pass server list.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
;                  |2 - Internal failure.
; Author ........: trancexx
; Related .......: _WinHttpDetectAutoProxyConfigUrl, _WinHttpGetDefaultProxyConfiguration, _WinHttpSetDefaultProxyConfiguration
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384096.aspx
;============================================================================================
Func _WinHttpGetIEProxyConfigForCurrentUser()
	Local $tWINHTTP_CURRENT_USER_IE_PROXY_CONFIG = DllStructCreate("int AutoDetect;" & _
			"ptr AutoConfigUrl;" & _
			"ptr Proxy;" & _
			"ptr ProxyBypass;")
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpGetIEProxyConfigForCurrentUser", "struct*", $tWINHTTP_CURRENT_USER_IE_PROXY_CONFIG)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Local $iAutoDetect = DllStructGetData($tWINHTTP_CURRENT_USER_IE_PROXY_CONFIG, "AutoDetect")
	Local $pAutoConfigUrl = DllStructGetData($tWINHTTP_CURRENT_USER_IE_PROXY_CONFIG, "AutoConfigUrl")
	Local $pProxy = DllStructGetData($tWINHTTP_CURRENT_USER_IE_PROXY_CONFIG, "Proxy")
	Local $pProxyBypass = DllStructGetData($tWINHTTP_CURRENT_USER_IE_PROXY_CONFIG, "ProxyBypass")
	Local $sAutoConfigUrl
	If $pAutoConfigUrl Then
		Local $iAutoConfigUrlLen = __WinHttpPtrStringLenW($pAutoConfigUrl)
		If Not @error Then
			Local $tAutoConfigUrl = DllStructCreate("wchar[" & $iAutoConfigUrlLen + 1 & "]", $pAutoConfigUrl)
			$sAutoConfigUrl = DllStructGetData($tAutoConfigUrl, 1)
			__WinHttpMemGlobalFree($pAutoConfigUrl)
		EndIf
	EndIf
	Local $sProxy
	If $pProxy Then
		Local $iProxyLen = __WinHttpPtrStringLenW($pProxy)
		If Not @error Then
			Local $tProxy = DllStructCreate("wchar[" & $iProxyLen + 1 & "]", $pProxy)
			$sProxy = DllStructGetData($tProxy, 1)
			__WinHttpMemGlobalFree($pProxy)
		EndIf
	EndIf
	Local $sProxyBypass
	If $pProxyBypass Then
		Local $iProxyBypassLen = __WinHttpPtrStringLenW($pProxyBypass)
		If Not @error Then
			Local $tProxyBypass = DllStructCreate("wchar[" & $iProxyBypassLen + 1 & "]", $pProxyBypass)
			$sProxyBypass = DllStructGetData($tProxyBypass, 1)
			__WinHttpMemGlobalFree($pProxyBypass)
		EndIf
	EndIf
	Local $aOutput[4] = [$iAutoDetect, $sAutoConfigUrl, $sProxy, $sProxyBypass]
	Return $aOutput
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpOpen
; Description ...: Initializes the use of WinHttp functions and returns a WinHttp-session handle.
; Syntax.........: _WinHttpOpen([$sUserAgent = Default [, $iAccessType = Default [, $sProxyName = Default [, $sProxyBypass = Default [, $iFlag = Default ]]]]])
; Parameters ....: $sUserAgent - [optional] The name of the application or entity calling the WinHttp functions.
;                  $iAccessType - [optional] Type of access required. Default is $WINHTTP_ACCESS_TYPE_NO_PROXY.
;                  $sProxyName - [optional] The name of the proxy server to use when proxy access is specified by setting $iAccessType to $WINHTTP_ACCESS_TYPE_NAMED_PROXY. Default is $WINHTTP_NO_PROXY_NAME.
;                  $sProxyBypass - [optional] An optional list of host names or IP addresses, or both, that should not be routed through the proxy when $iAccessType is set to $WINHTTP_ACCESS_TYPE_NAMED_PROXY. Default is $WINHTTP_NO_PROXY_BYPASS.
;                  $iFlag - [optional] Integer containing the flags that indicate various options affecting the behavior of this function. Default is 0.
; Return values .: Success - Returns valid session handle.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Remarks .......: <b>You are strongly discouraged to use WinHTTP in asynchronous mode with AutoIt. AutoIt's callback implementation can't handle reentrancy properly.</b>
;                  +For asynchronous mode set [[$iFlag]] to [[$WINHTTP_FLAG_ASYNC]]. In that case [[$WINHTTP_OPTION_CONTEXT_VALUE]] for the handle will inernally be set to [[$WINHTTP_FLAG_ASYNC]] also.
; Related .......: _WinHttpCloseHandle, _WinHttpConnect
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384098.aspx
;============================================================================================
Func _WinHttpOpen($sUserAgent = Default, $iAccessType = Default, $sProxyName = Default, $sProxyBypass = Default, $iFlag = Default)
	__WinHttpDefault($sUserAgent, __WinHttpUA())
	__WinHttpDefault($iAccessType, $WINHTTP_ACCESS_TYPE_NO_PROXY)
	__WinHttpDefault($sProxyName, $WINHTTP_NO_PROXY_NAME)
	__WinHttpDefault($sProxyBypass, $WINHTTP_NO_PROXY_BYPASS)
	__WinHttpDefault($iFlag, 0)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "handle", "WinHttpOpen", _
			"wstr", $sUserAgent, _
			"dword", $iAccessType, _
			"wstr", $sProxyName, _
			"wstr", $sProxyBypass, _
			"dword", $iFlag)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	If $iFlag = $WINHTTP_FLAG_ASYNC Then _WinHttpSetOption($aCall[0], $WINHTTP_OPTION_CONTEXT_VALUE, $WINHTTP_FLAG_ASYNC)
	Return $aCall[0]
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpOpenRequest
; Description ...: Creates an HTTP request handle.
; Syntax.........: _WinHttpOpenRequest($hConnect [, $sVerb = Default [, $sObjectName = Default [, $sVersion = Default [, $sReferrer = Default [, $sAcceptTypes = Default [, $iFlags = Default ]]]]]])
; Parameters ....: $hConnect - Handle to an HTTP session returned by _WinHttpConnect().
;                  $sVerb - [optional] HTTP verb to use in the request. Default is "GET".
;                  $sObjectName - [optional] The name of the target resource of the specified HTTP verb.
;                  $sVersion - [optional] HTTP version. Default is "HTTP/1.1"
;                  $sReferrer - [optional] URL of the document from which the URL in the request $sObjectName was obtained. Default is $WINHTTP_NO_REFERER.
;                  $sAcceptTypes - [optional] Media types accepted by the client. Default is $WINHTTP_DEFAULT_ACCEPT_TYPES
;                  $iFlags - [optional] Integer specifying the Internet flag values. Default is $WINHTTP_FLAG_ESCAPE_DISABLE
; Return values .: Success - Returns valid session handle.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Related .......: _WinHttpCloseHandle, _WinHttpConnect, _WinHttpSendRequest
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384099.aspx
;============================================================================================
Func _WinHttpOpenRequest($hConnect, $sVerb = Default, $sObjectName = Default, $sVersion = Default, $sReferrer = Default, $sAcceptTypes = Default, $iFlags = Default)
	__WinHttpDefault($sVerb, "GET")
	__WinHttpDefault($sObjectName, "")
	__WinHttpDefault($sVersion, "HTTP/1.1")
	__WinHttpDefault($sReferrer, $WINHTTP_NO_REFERER)
	__WinHttpDefault($iFlags, $WINHTTP_FLAG_ESCAPE_DISABLE)
	Local $pAcceptTypes
	If $sAcceptTypes = Default Or Number($sAcceptTypes) = -1 Then
		$pAcceptTypes = $WINHTTP_DEFAULT_ACCEPT_TYPES
	Else
		Local $aTypes = StringSplit($sAcceptTypes, ",", 2)
		Local $tAcceptTypes = DllStructCreate("ptr[" & UBound($aTypes) + 1 & "]")
		Local $tType[UBound($aTypes)]
		For $i = 0 To UBound($aTypes) - 1
			$tType[$i] = DllStructCreate("wchar[" & StringLen($aTypes[$i]) + 1 & "]")
			DllStructSetData($tType[$i], 1, $aTypes[$i])
			DllStructSetData($tAcceptTypes, 1, DllStructGetPtr($tType[$i]), $i + 1)
		Next
		$pAcceptTypes = DllStructGetPtr($tAcceptTypes)
	EndIf
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "handle", "WinHttpOpenRequest", _
			"handle", $hConnect, _
			"wstr", StringUpper($sVerb), _
			"wstr", $sObjectName, _
			"wstr", StringUpper($sVersion), _
			"wstr", $sReferrer, _
			"ptr", $pAcceptTypes, _
			"dword", $iFlags)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinHttpQueryAuthSchemes
; Description ...: Returns the authorization schemes that are supported by the server.
; Syntax ........: _WinHttpQueryAuthSchemes($hRequest, Byref $iSupportedSchemes, Byref $iFirstScheme, Byref $iAuthTarget)
; Parameters ....: $hRequest - Valid handle returned by _WinHttpSendRequest().
;                  $iSupportedSchemes - [out] Supported authentication schemes. See remarks.
;                  $iFirstScheme - [out] First authentication scheme listed by the server. See remarks.
;                  $iAuthTarget - [out] A flag that contains the authentication target. See remarks.
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Remarks .......: _WinHttpQueryAuthSchemes() is called after _WinHttpQueryHeaders().
;                  +Arguments are accepted ByRef.
;                  +Both [[$iSupportedSchemes]] and [[$iFirstScheme]] is set to combination of any of [[$WINHTTP_AUTH_SCHEME_]] flags.
;                  +[[$iAuthTarget]] parameter is set to one or more [[$WINHTTP_AUTH_TARGET_]] constants values.
; Related .......: _WinHttpSetCredentials, _WinHttpQueryHeaders, _WinHttpOpenRequest
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384100.aspx
; ===============================================================================================================================
Func _WinHttpQueryAuthSchemes($hRequest, ByRef $iSupportedSchemes, ByRef $iFirstScheme, ByRef $iAuthTarget)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpQueryAuthSchemes", _
			"handle", $hRequest, _
			"dword*", 0, _
			"dword*", 0, _
			"dword*", 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	$iSupportedSchemes = $aCall[2]
	$iFirstScheme = $aCall[3]
	$iAuthTarget = $aCall[4]
	Return 1
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpQueryDataAvailable
; Description ...: Returns the availability to be read with _WinHttpReadData().
; Syntax.........: _WinHttpQueryDataAvailable($hRequest)
; Parameters ....: $hRequest - handle returned by _WinHttpOpenRequest().
; Return values .: Success - Returns 1 if data is available.
;                          - Returns 0 if no data is available.
;                          - @extended receives the number of available bytes.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Remarks .......: _WinHttpReceiveResponse must have been called for this handle and completed before _WinHttpQueryDataAvailable is called.
; Related .......: _WinHttpOpenRequest, _WinHttpReadData, _WinHttpReceiveResponse
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384101.aspx
;============================================================================================
Func _WinHttpQueryDataAvailable($hRequest)
	Local $sReadType = "dword*"
	If BitAND(_WinHttpQueryOption(_WinHttpQueryOption(_WinHttpQueryOption($hRequest, $WINHTTP_OPTION_PARENT_HANDLE), $WINHTTP_OPTION_PARENT_HANDLE), $WINHTTP_OPTION_CONTEXT_VALUE), $WINHTTP_FLAG_ASYNC) Then $sReadType = "ptr"
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpQueryDataAvailable", "handle", $hRequest, $sReadType, 0)
	If @error Then Return SetError(1, 0, 0)
	Return SetExtended($aCall[2], $aCall[0])
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpQueryHeaders
; Description ...: Retrieves header information associated with an HTTP request.
; Syntax.........: _WinHttpQueryHeaders($hRequest [, $iInfoLevel = Default [, $sName = Default [, $iIndex = Default ]]])
; Parameters ....: $hRequest - Handle returned by _WinHttpOpenRequest().
;                  $iInfoLevel - [optional] A combination of attribute and modifier flags. Default is $WINHTTP_QUERY_RAW_HEADERS_CRLF.
;                  $sName - [optional] Header name string. Default is $WINHTTP_HEADER_NAME_BY_INDEX.
;                  $iIndex - [optional] Index used to enumerate multiple headers with the same name
; Return values .: Success - Returns string that contains header.
;                          - @extended is set to the index of the next header
;                  Failure - Returns empty string and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Related .......: _WinHttpAddRequestHeaders, _WinHttpOpenRequest
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384102.aspx
;============================================================================================
Func _WinHttpQueryHeaders($hRequest, $iInfoLevel = Default, $sName = Default, $iIndex = Default)
	__WinHttpDefault($iInfoLevel, $WINHTTP_QUERY_RAW_HEADERS_CRLF)
	__WinHttpDefault($sName, $WINHTTP_HEADER_NAME_BY_INDEX)
	__WinHttpDefault($iIndex, $WINHTTP_NO_HEADER_INDEX)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpQueryHeaders", _
			"handle", $hRequest, _
			"dword", $iInfoLevel, _
			"wstr", $sName, _
			"wstr", "", _
			"dword*", 65536, _
			"dword*", $iIndex)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, "")
	Return SetExtended($aCall[6], $aCall[4])
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpQueryOption
; Description ...: Queries an Internet option on the specified handle.
; Syntax.........: _WinHttpQueryOption($hInternet, $iOption)
; Parameters ....: $hInternet - Handle on which to query information.
;                  $iOption - Internet option to query.
; Return values .: Success - Returns data containing requested information.
;                  Failure - Returns empty string and sets @error:
;                  |1 - Initial DllCall failed
;                  |2 - Main DllCall failed
; Author ........: trancexx
; Remarks .......: Type of the returned data varies on request.
; Related .......: _WinHttpSetOption
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384103.aspx
;============================================================================================
Func _WinHttpQueryOption($hInternet, $iOption)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpQueryOption", _
			"handle", $hInternet, _
			"dword", $iOption, _
			"ptr", 0, _
			"dword*", 0)
	If @error Or $aCall[0] Then Return SetError(1, 0, "")
	Local $iSize = $aCall[4]
	Local $tBuffer
	Switch $iOption
		Case $WINHTTP_OPTION_CONNECTION_INFO, $WINHTTP_OPTION_PASSWORD, $WINHTTP_OPTION_PROXY_PASSWORD, $WINHTTP_OPTION_PROXY_USERNAME, $WINHTTP_OPTION_URL, $WINHTTP_OPTION_USERNAME, $WINHTTP_OPTION_USER_AGENT, _
				$WINHTTP_OPTION_PASSPORT_COBRANDING_TEXT, $WINHTTP_OPTION_PASSPORT_COBRANDING_URL
			$tBuffer = DllStructCreate("wchar[" & $iSize + 1 & "]")
		Case $WINHTTP_OPTION_PARENT_HANDLE, $WINHTTP_OPTION_CALLBACK, $WINHTTP_OPTION_SERVER_CERT_CONTEXT
			$tBuffer = DllStructCreate("ptr")
		Case $WINHTTP_OPTION_CONNECT_TIMEOUT, $WINHTTP_AUTOLOGON_SECURITY_LEVEL_HIGH, $WINHTTP_AUTOLOGON_SECURITY_LEVEL_LOW, $WINHTTP_AUTOLOGON_SECURITY_LEVEL_MEDIUM, _
				$WINHTTP_OPTION_CONFIGURE_PASSPORT_AUTH, $WINHTTP_OPTION_CONNECT_RETRIES, $WINHTTP_OPTION_EXTENDED_ERROR, $WINHTTP_OPTION_HANDLE_TYPE, $WINHTTP_OPTION_MAX_CONNS_PER_1_0_SERVER, _
				$WINHTTP_OPTION_MAX_CONNS_PER_SERVER, $WINHTTP_OPTION_MAX_HTTP_AUTOMATIC_REDIRECTS, $WINHTTP_OPTION_RECEIVE_RESPONSE_TIMEOUT, $WINHTTP_OPTION_RECEIVE_TIMEOUT, _
				$WINHTTP_OPTION_RESOLVE_TIMEOUT, $WINHTTP_OPTION_SECURITY_FLAGS, $WINHTTP_OPTION_SECURITY_KEY_BITNESS, $WINHTTP_OPTION_SEND_TIMEOUT
			$tBuffer = DllStructCreate("int")
		Case $WINHTTP_OPTION_CONTEXT_VALUE
			$tBuffer = DllStructCreate("dword_ptr")
		Case Else
			$tBuffer = DllStructCreate("byte[" & $iSize & "]")
	EndSwitch
	$aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpQueryOption", _
			"handle", $hInternet, _
			"dword", $iOption, _
			"struct*", $tBuffer, _
			"dword*", $iSize)
	If @error Or Not $aCall[0] Then Return SetError(2, 0, "")
	Return DllStructGetData($tBuffer, 1)
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpReadData
; Description ...: Reads data from a handle opened by the _WinHttpOpenRequest() function.
; Syntax.........: _WinHttpReadData($hRequest [, $iMode = Default [, $iNumberOfBytesToRead = Default ]])
; Parameters ....: $hRequest - Valid handle returned from a previous call to _WinHttpOpenRequest().
;                  $iMode - [optional] Integer representing reading mode. Default is 0 (charset is decoded as it is ANSI).
;                  $iNumberOfBytesToRead - [optional] The number of bytes to read. Default is 8192 bytes.
; Return values .: Success - Returns data read.
;                          - @extended receives the number of bytes read.
;                  Special: Sets @error to -1 if no more data to read (end reached).
;                  Failure - Returns empty string and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx, ProgAndy
; Remarks .......: [[$iMode]] can have these values:
;                  |[[0]] - ANSI
;                  |[[1]] - UTF8
;                  |[[2]] - Binary
; Related .......: _WinHttpOpenRequest, _WinHttpWriteData
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384104.aspx
;============================================================================================
Func _WinHttpReadData($hRequest, $iMode = Default, $iNumberOfBytesToRead = Default, $pBuffer = Default)
	__WinHttpDefault($iMode, 0)
	__WinHttpDefault($iNumberOfBytesToRead, 8192)
	Local $tBuffer, $vOutOnError = ""
	If $iMode = 2 Then $vOutOnError = Binary($vOutOnError)
	Switch $iMode
		Case 1, 2
			If $pBuffer And $pBuffer <> Default Then
				$tBuffer = DllStructCreate("byte[" & $iNumberOfBytesToRead & "]", $pBuffer)
			Else
				$tBuffer = DllStructCreate("byte[" & $iNumberOfBytesToRead & "]")
			EndIf
		Case Else
			$iMode = 0
			If $pBuffer And $pBuffer <> Default Then
				$tBuffer = DllStructCreate("char[" & $iNumberOfBytesToRead & "]", $pBuffer)
			Else
				$tBuffer = DllStructCreate("char[" & $iNumberOfBytesToRead & "]")
			EndIf
	EndSwitch
	Local $sReadType = "dword*"
	If BitAND(_WinHttpQueryOption(_WinHttpQueryOption(_WinHttpQueryOption($hRequest, $WINHTTP_OPTION_PARENT_HANDLE), $WINHTTP_OPTION_PARENT_HANDLE), $WINHTTP_OPTION_CONTEXT_VALUE), $WINHTTP_FLAG_ASYNC) Then $sReadType = "ptr"
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpReadData", _
			"handle", $hRequest, _
			"struct*", $tBuffer, _
			"dword", $iNumberOfBytesToRead, _
			$sReadType, 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, "")
	If Not $aCall[4] Then Return SetError(-1, 0, $vOutOnError)
	If $aCall[4] < $iNumberOfBytesToRead Then
		Switch $iMode
			Case 0
				Return SetExtended($aCall[4], StringLeft(DllStructGetData($tBuffer, 1), $aCall[4]))
			Case 1
				Return SetExtended($aCall[4], BinaryToString(BinaryMid(DllStructGetData($tBuffer, 1), 1, $aCall[4]), 4))
			Case 2
				Return SetExtended($aCall[4], BinaryMid(DllStructGetData($tBuffer, 1), 1, $aCall[4]))
		EndSwitch
	Else
		Switch $iMode
			Case 0, 2
				Return SetExtended($aCall[4], DllStructGetData($tBuffer, 1))
			Case 1
				Return SetExtended($aCall[4], BinaryToString(DllStructGetData($tBuffer, 1), 4))
		EndSwitch
	EndIf
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpReceiveResponse
; Description ...: Waits to receive the response to an HTTP request initiated by WinHttpSendRequest().
; Syntax.........: _WinHttpReceiveResponse($hRequest)
; Parameters ....: $hRequest - Handle returned by _WinHttpOpenRequest() and sent by _WinHttpSendRequest().
; Return values .: Success - Returns 1.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Remarks .......: Call to _WinHttpReceiveResponse() must be done before _WinHttpQueryDataAvailable() and _WinHttpReadData().
; Related .......: _WinHttpOpenRequest, _WinHttpSetTimeouts
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384105.aspx
;============================================================================================
Func _WinHttpReceiveResponse($hRequest)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpReceiveResponse", "handle", $hRequest, "ptr", 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpSendRequest
; Description ...: Sends the specified request to the HTTP server.
; Syntax.........: _WinHttpSendRequest($hRequest [, $sHeaders = Default [, $sOptional = Default [, $iTotalLength = Default [, $iContext = Default ]]]])
; Parameters ....: $hRequest - Handle returned by _WinHttpOpenRequest().
;                  $sHeaders - [optional] Additional headers to append to the request. Default is $WINHTTP_NO_ADDITIONAL_HEADERS.
;                  $vOptional - [optional] Optional data to send immediately after the request headers. Default is $WINHTTP_NO_REQUEST_DATA.
;                  $iTotalLength - [optional] Length, in bytes, of the total optional data sent. Default is 0.
;                  $iContext - [optional] Application-defined value that is passed, with the request handle, to any callback functions. Default is 0.
; Return values .: Success - Returns 1.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Remarks .......: Specifying optional data [[$vOptional]] will cause [[$iTotalLength]] to receive the size of that data if left default value.
; Related .......: _WinHttpOpenRequest
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384110.aspx
;============================================================================================
Func _WinHttpSendRequest($hRequest, $sHeaders = Default, $vOptional = Default, $iTotalLength = Default, $iContext = Default)
	__WinHttpDefault($sHeaders, $WINHTTP_NO_ADDITIONAL_HEADERS)
	__WinHttpDefault($vOptional, $WINHTTP_NO_REQUEST_DATA)
	__WinHttpDefault($iTotalLength, 0)
	__WinHttpDefault($iContext, 0)
	Local $pOptional = 0, $iOptionalLength = 0
	If @NumParams > 2 Then
		Local $tOptional
		$iOptionalLength = BinaryLen($vOptional)
		$tOptional = DllStructCreate("byte[" & $iOptionalLength & "]")
		If $iOptionalLength Then $pOptional = DllStructGetPtr($tOptional)
		DllStructSetData($tOptional, 1, $vOptional)
	EndIf
	If Not $iTotalLength Or $iTotalLength < $iOptionalLength Then $iTotalLength += $iOptionalLength
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpSendRequest", _
			"handle", $hRequest, _
			"wstr", $sHeaders, _
			"dword", 0, _
			"ptr", $pOptional, _
			"dword", $iOptionalLength, _
			"dword", $iTotalLength, _
			"dword_ptr", $iContext)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpSetCredentials
; Description ...: Passes the required authorization credentials to the server.
; Syntax.........: _WinHttpSetCredentials($hRequest, $iAuthTargets, $iAuthScheme, $sUserName, $sPassword)
; Parameters ....: $hRequest - Valid handle returned by _WinHttpOpenRequest().
;                  $iAuthTargets - Authentication target.
;                  $iAuthScheme - Authentication scheme.
;                  $sUserName - Valid user name.
;                  $sPassword - Valid password.
; Return values .: Success - Returns 1.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Related .......: _WinHttpOpenRequest
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384112.aspx
;============================================================================================
Func _WinHttpSetCredentials($hRequest, $iAuthTargets, $iAuthScheme, $sUserName, $sPassword)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpSetCredentials", _
			"handle", $hRequest, _
			"dword", $iAuthTargets, _
			"dword", $iAuthScheme, _
			"wstr", $sUserName, _
			"wstr", $sPassword, _
			"ptr", 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpSetDefaultProxyConfiguration
; Description ...: Sets the default WinHttp proxy configuration.
; Syntax.........: _WinHttpSetDefaultProxyConfiguration($iAccessType [, $sProxy = "" [, $sProxyBypass = ""])
; Parameters ....: $iAccessType - Integer. Access type.
;                  $sProxy - [optional] String. Proxy server list.
;                  $sProxyBypass - [optional] String. Proxy bypass list.
; Return values .: Success - Returns 1.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Related .......: _WinHttpDetectAutoProxyConfigUrl, _WinHttpGetDefaultProxyConfiguration, _WinHttpGetIEProxyConfigForCurrentUser
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384113.aspx
;============================================================================================
Func _WinHttpSetDefaultProxyConfiguration($iAccessType, $sProxy = "", $sProxyBypass = "")
	Local $tProxy = DllStructCreate("wchar[" & StringLen($sProxy) + 1 & "]")
	DllStructSetData($tProxy, 1, $sProxy)
	Local $tProxyBypass = DllStructCreate("wchar[" & StringLen($sProxyBypass) + 1 & "]")
	DllStructSetData($tProxyBypass, 1, $sProxyBypass)
	Local $tWINHTTP_PROXY_INFO = DllStructCreate("dword AccessType;" & _
			"ptr Proxy;" & _
			"ptr ProxyBypass")
	DllStructSetData($tWINHTTP_PROXY_INFO, "AccessType", $iAccessType)
	If $iAccessType <> $WINHTTP_ACCESS_TYPE_NO_PROXY Then
		DllStructSetData($tWINHTTP_PROXY_INFO, "Proxy", DllStructGetPtr($tProxy))
		DllStructSetData($tWINHTTP_PROXY_INFO, "ProxyBypass", DllStructGetPtr($tProxyBypass))
	EndIf
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpSetDefaultProxyConfiguration", "struct*", $tWINHTTP_PROXY_INFO)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpSetOption
; Description ...: Sets an Internet option.
; Syntax.........: _WinHttpSetOption($hInternet, $iOption, $vSetting [, $iSize = Default ])
; Parameters ....: $hInternet - Handle on which to set data.
;                  $iOption - Integer value that contains the Internet option to set.
;                  $vSetting - Value of setting
;                  $iSize    - [optional] Size of $vSetting, required if $vSetting is pointer to memory block
; Return values .: Success - Returns 1.
;                  Failure - Returns 0 and sets @error:
;                  |1 - Invalid Internet option
;                  |2 - Size required
;                  |3 - Datatype of value does not fit to option
;                  |4 - DllCall failed
; Author ........: ProgAndy, trancexx
; Related .......: _WinHttpQueryOption
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384114.aspx
;============================================================================================
Func _WinHttpSetOption($hInternet, $iOption, $vSetting, $iSize = Default)
	If $iSize = Default Then $iSize = -1
	If IsBinary($vSetting) Then
		$iSize = DllStructCreate("byte[" & BinaryLen($vSetting) & "]")
		DllStructSetData($iSize, 1, $vSetting)
		$vSetting = $iSize
		$iSize = DllStructGetSize($vSetting)
	EndIf
	Local $sType
	Switch $iOption
		Case $WINHTTP_OPTION_AUTOLOGON_POLICY, $WINHTTP_OPTION_CODEPAGE, $WINHTTP_OPTION_CONFIGURE_PASSPORT_AUTH, $WINHTTP_OPTION_CONNECT_RETRIES, _
				$WINHTTP_OPTION_CONNECT_TIMEOUT, $WINHTTP_OPTION_DISABLE_FEATURE, $WINHTTP_OPTION_ENABLE_FEATURE, $WINHTTP_OPTION_ENABLETRACING, _
				$WINHTTP_OPTION_MAX_CONNS_PER_1_0_SERVER, $WINHTTP_OPTION_MAX_CONNS_PER_SERVER, $WINHTTP_OPTION_MAX_HTTP_AUTOMATIC_REDIRECTS, _
				$WINHTTP_OPTION_MAX_HTTP_STATUS_CONTINUE, $WINHTTP_OPTION_MAX_RESPONSE_DRAIN_SIZE, $WINHTTP_OPTION_MAX_RESPONSE_HEADER_SIZE, _
				$WINHTTP_OPTION_READ_BUFFER_SIZE, $WINHTTP_OPTION_RECEIVE_TIMEOUT, _
				$WINHTTP_OPTION_RECEIVE_RESPONSE_TIMEOUT, $WINHTTP_OPTION_REDIRECT_POLICY, $WINHTTP_OPTION_REJECT_USERPWD_IN_URL, _
				$WINHTTP_OPTION_REQUEST_PRIORITY, $WINHTTP_OPTION_RESOLVE_TIMEOUT, $WINHTTP_OPTION_SECURE_PROTOCOLS, $WINHTTP_OPTION_SECURITY_FLAGS, _
				$WINHTTP_OPTION_SECURITY_KEY_BITNESS, $WINHTTP_OPTION_SEND_TIMEOUT, $WINHTTP_OPTION_SPN, $WINHTTP_OPTION_USE_GLOBAL_SERVER_CREDENTIALS, _
				$WINHTTP_OPTION_WORKER_THREAD_COUNT, $WINHTTP_OPTION_WRITE_BUFFER_SIZE, $WINHTTP_OPTION_DECOMPRESSION, $WINHTTP_OPTION_UNSAFE_HEADER_PARSING
			$sType = "dword*"
			$iSize = 4
		Case $WINHTTP_OPTION_CALLBACK, $WINHTTP_OPTION_PASSPORT_SIGN_OUT
			$sType = "ptr*"
			$iSize = 4
			If @AutoItX64 Then $iSize = 8
			If Not IsPtr($vSetting) Then Return SetError(3, 0, 0)
		Case $WINHTTP_OPTION_CONTEXT_VALUE
			$sType = "dword_ptr*"
			$iSize = 4
			If @AutoItX64 Then $iSize = 8
		Case $WINHTTP_OPTION_PASSWORD, $WINHTTP_OPTION_PROXY_PASSWORD, $WINHTTP_OPTION_PROXY_USERNAME, $WINHTTP_OPTION_USER_AGENT, $WINHTTP_OPTION_USERNAME
			$sType = "wstr"
			If (IsDllStruct($vSetting) Or IsPtr($vSetting)) Then Return SetError(3, 0, 0)
			If $iSize < 1 Then $iSize = StringLen($vSetting)
		Case $WINHTTP_OPTION_CLIENT_CERT_CONTEXT, $WINHTTP_OPTION_GLOBAL_PROXY_CREDS, $WINHTTP_OPTION_GLOBAL_SERVER_CREDS, $WINHTTP_OPTION_HTTP_VERSION, _
				$WINHTTP_OPTION_PROXY
			$sType = "ptr"
			If Not (IsDllStruct($vSetting) Or IsPtr($vSetting)) Then Return SetError(3, 0, 0)
		Case Else
			Return SetError(1, 0, 0)
	EndSwitch
	If $iSize < 1 Then
		If IsDllStruct($vSetting) Then
			$iSize = DllStructGetSize($vSetting)
		Else
			Return SetError(2, 0, 0)
		EndIf
	EndIf
	Local $aCall
	If IsDllStruct($vSetting) Then
		$aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpSetOption", "handle", $hInternet, "dword", $iOption, $sType, DllStructGetPtr($vSetting), "dword", $iSize)
	Else
		$aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpSetOption", "handle", $hInternet, "dword", $iOption, $sType, $vSetting, "dword", $iSize)
	EndIf
	If @error Or Not $aCall[0] Then Return SetError(4, 0, 0)
	Return 1
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpSetStatusCallback
; Description ...: Sets up a callback function that WinHttp can call as progress is made during an operation.
; Syntax.........: _WinHttpSetStatusCallback($hInternet, $hInternetCallback [, $iNotificationFlags = Default ])
; Parameters ....: $hInternet - Handle for which the callback is to be set.
;                  $hInternetCallback - Callback function to call when progress is made.
;                  $iNotificationFlags - [optional] Flags to indicate which events activate the callback function. Default is $WINHTTP_CALLBACK_FLAG_ALL_NOTIFICATIONS.
; Return values .: Success - Returns a pointer to the previously defined status callback function or NULL if there was no previously defined status callback function.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: ProgAndy
; Modified.......: trancexx
; Related .......: _WinHttpOpen
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384115.aspx
;============================================================================================
Func _WinHttpSetStatusCallback($hInternet, $hInternetCallback, $iNotificationFlags = Default)
	__WinHttpDefault($iNotificationFlags, $WINHTTP_CALLBACK_FLAG_ALL_NOTIFICATIONS)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "ptr", "WinHttpSetStatusCallback", _
			"handle", $hInternet, _
			"ptr", DllCallbackGetPtr($hInternetCallback), _
			"dword", $iNotificationFlags, _
			"ptr", 0)
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpSetTimeouts
; Description ...: Sets time-outs involved with HTTP transactions.
; Syntax.........: _WinHttpSetTimeouts($hInternet [, $iResolveTimeout = Default [, $iConnectTimeout = Default [, $iSendTimeout = Default [, $iReceiveTimeout = Default ]]]])
; Parameters ....: $hInternet - Handle returned by _WinHttpOpen() or _WinHttpOpenRequest().
;                  $iResolveTimeout - [optional] Time-out value, in milliseconds, to use for name resolution. Default is 0 ms.
;                  $iConnectTimeout - [optional] Time-out value, in milliseconds, to use for server connection requests. Default is 60000 ms.
;                  $iSendTimeout - [optional] Time-out value, in milliseconds, to use for sending requests. Default is 30000 ms.
;                  $iReceiveTimeout - [optional] Time-out value, in milliseconds, to receive a response to a request. Default is 30000 ms.
; Return values .: Success - Returns 1.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Remarks .......: Initial values are:
;                  |- $iResolveTimeout = 0
;                  |- $iConnectTimeout = 60000
;                  |- $iSendTimeout = 30000
;                  |- $iReceiveTimeout = 30000
; Related .......: _WinHttpReceiveResponse
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384116.aspx
;============================================================================================
Func _WinHttpSetTimeouts($hInternet, $iResolveTimeout = Default, $iConnectTimeout = Default, $iSendTimeout = Default, $iReceiveTimeout = Default)
	__WinHttpDefault($iResolveTimeout, 0)
	__WinHttpDefault($iConnectTimeout, 60000)
	__WinHttpDefault($iSendTimeout, 30000)
	__WinHttpDefault($iReceiveTimeout, 30000)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpSetTimeouts", _
			"handle", $hInternet, _
			"int", $iResolveTimeout, _
			"int", $iConnectTimeout, _
			"int", $iSendTimeout, _
			"int", $iReceiveTimeout)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpSimpleBinaryConcat
; Description ...: Concatenates two binary data returned by _WinHttpReadData() in binary mode.
; Syntax.........: _WinHttpSimpleBinaryConcat(ByRef $bBinary1, ByRef $bBinary2)
; Parameters ....: $bBinary1 - Binary data that is to be concatenated.
;                  $bBinary2 - Binary data to concatenate.
; Return values .: Success - Returns concatenated binary data.
;                  Failure - Returns empty binary and sets @error:
;                  |1 - Invalid input.
; Author ........: ProgAndy
; Modified.......: trancexx
; Related .......: _WinHttpReadData
;============================================================================================
Func _WinHttpSimpleBinaryConcat(ByRef $bBinary1, ByRef $bBinary2)
	Switch IsBinary($bBinary1) + 2 * IsBinary($bBinary2)
		Case 0
			Return SetError(1, 0, Binary(''))
		Case 1
			Return $bBinary1
		Case 2
			Return $bBinary2
	EndSwitch
	Local $tAuxiliary = DllStructCreate("byte[" & BinaryLen($bBinary1) & "];byte[" & BinaryLen($bBinary2) & "]")
	DllStructSetData($tAuxiliary, 1, $bBinary1)
	DllStructSetData($tAuxiliary, 2, $bBinary2)
	Local $tOutput = DllStructCreate("byte[" & DllStructGetSize($tAuxiliary) & "]", DllStructGetPtr($tAuxiliary))
	Return DllStructGetData($tOutput, 1)
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpSimpleFormFill
; Description ...: Fills web form.
; Syntax.........: _WinHttpSimpleFormFill(ByRef $hInternet [, $sActionPage = Default [, $sFormId = Default [, $sFldId1 = Default [, $sDta1 = Default [, (...) [, $sAdditionalData]]]]]])
; Parameters ....: $hInternet - Handle returned by _WinHttpConnect() or string variable with form.
;                  $sActionPage -  [optional] path to the page with form or session handle if $hInternet is string (default: "" - empty string; meaning 'default' page on the server in former).
;                  $sFormId - [optional] Id of the form. Can be name or zero-based index too (read Remarks section).
;                  $sFldId1 - [optional] Id of the input.
;                  $sDta1 - [optional] Data to set to coresponding field.
;                  (...) - [optional] Other pairs of Id/Data. Overall number of fields is 40.
;                  $sAdditionalData - [optional] Additional data (read Remarks section).
; Return values .: Success - Returns HTML source of the page returned by the server on submited form. @extended is set to HTTP status code.
;                  Failure - Returns empty string and sets @error:
;                  |1 - No forms on the page
;                  |2 - Invalid form
;                  |3 - No forms with specified attributes on the page
;                  |4 - Connection problems
;                  |5 - form's "action" is invalid
;                  |6 - invalid session handle passed
; Author ........: trancexx
; Remarks .......: In case form requires redirection and [[$hInternet]] is internet handle, this handle will be closed and replaced with new and required one.
;                  +When [[$hInternet]] is form string, form's "action" must specify URL and [[$sActionPage]] parameter must be session handle. On succesful call this variable will be changed to connection handle of the internally made connection.
;                  Don't forget to close this handle after the function returns and when no longer needed.
;                  +[[$sFormId]] specifies Id of the form same as [[.getElementById(FormId)]]. Aditionally you can use [["index:FormIndex"]] to
;                  identify form by its zero-based index number (in case of e.g. three forms on some page first one will have index=0, second index=1, third index=2).
;                  Use [["name:FormName"]] to identify form by its name like with [[.getElementsByName(FormName)]]. FormName will be taken to be what's right of colon mark.
;                  In that case first form with that name is filled.
;                  +As for fields, If [["name:FieldName"]] option is used all the fields except last with that name are removed from the form. Last one is filled with specified [[$sDta]] data.
;                  +This function can be used to fill forms with up to 40 fields at once.
;                  +"Submit" control you want to keep (click) set to True. If no such control is set then the first one found in the form is "clicked". You can also use [["type:submit", zero_based_index_of_the_submit]] to "click" if no id or name is available.
;                  +All other "submit" controls are removed from the submited form (including images).
;                  +If form is submitted by clicking an image then pass click coordinates [["name:image_name", "Xcoord,Ycoord"]] or [["image_id", "Xcoord,Ycoord"]]. If the image has no name or id then you can use its index of appearance [["type:image", "zero_based_index_of_the_image Xcoord,Ycoord"]].
;                  +"Checkbox" and "Button" input types are removed from the submitted form unless explicitly set. Same goes for "Radio" with exception that
;                  only one such control can be set, the rest are removed. These controls are set by their values. Wrong value makes them invalid and therefore not part of the submited data.
;                  +All other non-set fields are left default.
;                  +Last (superfluous) [[$sAdditionalData]] argument can be used to pass authorization credentials in form [["[CRED:username:password]"]], magic string to ignore certificate errors in form [["[IGNORE_CERT_ERRORS]"]], change output type to extended array with [["[RETURN_ARRAY]"]] moniker, and/or HTTP request header data to add.
;                  If array is returned then [[array[0]]] is the response header, [[array[1]]] is returned data and [[array[2]]] is URL of the final request.
;                  +
;                  +If this function is used to upload multiple files then there are two available ways. Default would be to submit the form following RFC2388 specification.
;                  In that case every file is represented as multipart/mixed part embedded within the multipart/form-data.
;                  +If you want to upload using alternative way (to avoid certain PHP bug that could exist on server side) then prefix the file string with [["PHP#50338:"]] string.
;                  +For example: [[..."name:files[]", "PHP#50338:" & $sFile1 & "|" & $sFile2 ...]]
;                  +Muliple files are always separated with vertical line ASCII character when filling the form.
; Related .......: _WinHttpConnect, _WinHttpSimpleFormFill_SetUploadCallback
;============================================================================================
Func _WinHttpSimpleFormFill(ByRef $hInternet, $sActionPage = Default, $sFormId = Default, $sFldId1 = Default, $sDta1 = Default, $sFldId2 = Default, $sDta2 = Default, $sFldId3 = Default, $sDta3 = Default, $sFldId4 = Default, $sDta4 = Default, $sFldId5 = Default, $sDta5 = Default, $sFldId6 = Default, $sDta6 = Default, $sFldId7 = Default, $sDta7 = Default, $sFldId8 = Default, $sDta8 = Default, $sFldId9 = Default, $sDta9 = Default, $sFldId10 = Default, $sDta10 = Default, _
		$sFldId11 = Default, $sDta11 = Default, $sFldId12 = Default, $sDta12 = Default, $sFldId13 = Default, $sDta13 = Default, $sFldId14 = Default, $sDta14 = Default, $sFldId15 = Default, $sDta15 = Default, $sFldId16 = Default, $sDta16 = Default, $sFldId17 = Default, $sDta17 = Default, $sFldId18 = Default, $sDta18 = Default, $sFldId19 = Default, $sDta19 = Default, $sFldId20 = Default, $sDta20 = Default, _
		$sFldId21 = Default, $sDta21 = Default, $sFldId22 = Default, $sDta22 = Default, $sFldId23 = Default, $sDta23 = Default, $sFldId24 = Default, $sDta24 = Default, $sFldId25 = Default, $sDta25 = Default, $sFldId26 = Default, $sDta26 = Default, $sFldId27 = Default, $sDta27 = Default, $sFldId28 = Default, $sDta28 = Default, $sFldId29 = Default, $sDta29 = Default, $sFldId30 = Default, $sDta30 = Default, _
		$sFldId31 = Default, $sDta31 = Default, $sFldId32 = Default, $sDta32 = Default, $sFldId33 = Default, $sDta33 = Default, $sFldId34 = Default, $sDta34 = Default, $sFldId35 = Default, $sDta35 = Default, $sFldId36 = Default, $sDta36 = Default, $sFldId37 = Default, $sDta37 = Default, $sFldId38 = Default, $sDta38 = Default, $sFldId39 = Default, $sDta39 = Default, $sFldId40 = Default, $sDta40 = Default)
	__WinHttpDefault($sActionPage, "")
	Local $iNumArgs = @NumParams, $sAdditionalHeaders, $sCredName, $sCredPass, $iIgnoreCertErr, $iRetArr
	Local $aDtas[41] = [0, $sDta1, $sDta2, $sDta3, $sDta4, $sDta5, $sDta6, $sDta7, $sDta8, $sDta9, $sDta10, $sDta11, $sDta12, $sDta13, $sDta14, $sDta15, $sDta16, $sDta17, $sDta18, $sDta19, $sDta20, $sDta21, $sDta22, $sDta23, $sDta24, $sDta25, $sDta26, $sDta27, $sDta28, $sDta29, $sDta30, $sDta31, $sDta32, $sDta33, $sDta34, $sDta35, $sDta36, $sDta37, $sDta38, $sDta39, $sDta40]
	Local $aFlds[41] = [0, $sFldId1, $sFldId2, $sFldId3, $sFldId4, $sFldId5, $sFldId6, $sFldId7, $sFldId8, $sFldId9, $sFldId10, $sFldId11, $sFldId12, $sFldId13, $sFldId14, $sFldId15, $sFldId16, $sFldId17, $sFldId18, $sFldId19, $sFldId20, $sFldId21, $sFldId22, $sFldId23, $sFldId24, $sFldId25, $sFldId26, $sFldId27, $sFldId28, $sFldId29, $sFldId30, $sFldId31, $sFldId32, $sFldId33, $sFldId34, $sFldId35, $sFldId36, $sFldId37, $sFldId38, $sFldId39, $sFldId40]
	If Not Mod($iNumArgs, 2) Then
		$sAdditionalHeaders = $aFlds[$iNumArgs / 2 - 1]
		$aFlds[$iNumArgs / 2 - 1] = 0
		$iIgnoreCertErr = StringInStr($sAdditionalHeaders, "[IGNORE_CERT_ERRORS]")
		If $iIgnoreCertErr Then $sAdditionalHeaders = StringReplace($sAdditionalHeaders, "[IGNORE_CERT_ERRORS]", "", 1)
		$iRetArr = StringInStr($sAdditionalHeaders, "[RETURN_ARRAY]")
		If $iRetArr Then $sAdditionalHeaders = StringReplace($sAdditionalHeaders, "[RETURN_ARRAY]", "", 1)
		Local $aCred = StringRegExp($sAdditionalHeaders, "\[CRED:(.*?)\]", 2)
		If Not @error Then
			Local $sCredDelim = ":"
			If Not StringInStr($aCred[1], $sCredDelim) Then $sCredDelim = ","
			Local $aStrSplit = StringSplit($aCred[1], $sCredDelim, 3)
			If Not @error Then
				$sCredName = $aStrSplit[0]
				$sCredPass = $aStrSplit[1]
			EndIf
			$sAdditionalHeaders = StringReplace($sAdditionalHeaders, $aCred[0], "", 1)
		EndIf
	EndIf
	; Get page source
	Local $hOpen, $aHTML, $sHTML, $sURL, $fVarForm, $iScheme = $INTERNET_SCHEME_HTTP
	If IsString($hInternet) Then ; $hInternet is page source
		$sHTML = $hInternet
		If _WinHttpQueryOption($sActionPage, $WINHTTP_OPTION_HANDLE_TYPE) <> $WINHTTP_HANDLE_TYPE_SESSION Then Return SetError(6, 0, "")
		$hOpen = $sActionPage
		$fVarForm = True
	Else
		$iScheme = _WinHttpQueryOption($hInternet, $WINHTTP_OPTION_CONTEXT_VALUE); read internet scheme from the connection handle
		Local $sAccpt = "Accept: text/html;q=0.9,text/plain;q=0.8,*/*;q=0.5"
		If $iScheme = $INTERNET_SCHEME_HTTPS Then
			$aHTML = _WinHttpSimpleSSLRequest($hInternet, Default, $sActionPage, Default, Default, $sAccpt, 1, Default, $sCredName, $sCredPass, $iIgnoreCertErr)
		ElseIf $iScheme = $INTERNET_SCHEME_HTTP Then
			$aHTML = _WinHttpSimpleRequest($hInternet, Default, $sActionPage, Default, Default, $sAccpt, 1, Default, $sCredName, $sCredPass)
		Else
			; Try both http and https scheme and deduct the right one besed on response
			$aHTML = _WinHttpSimpleRequest($hInternet, Default, $sActionPage, Default, Default, $sAccpt, 1, Default, $sCredName, $sCredPass)
			If @error Or @extended >= $HTTP_STATUS_BAD_REQUEST Then
				$aHTML = _WinHttpSimpleSSLRequest($hInternet, Default, $sActionPage, Default, Default, $sAccpt, 1, Default, $sCredName, $sCredPass, $iIgnoreCertErr)
				$iScheme = $INTERNET_SCHEME_HTTPS
			Else
				$iScheme = $INTERNET_SCHEME_HTTP
			EndIf
		EndIf
		If Not @error Then ; Error is checked after If...Then...Else block. Be careful before changing anything!
			$sHTML = $aHTML[1]
			$sURL = $aHTML[2]
		EndIf
	EndIf
	$sHTML = StringRegExpReplace($sHTML, "(?s)<!--.*?-->", "") ; removing comments
	$sHTML = StringRegExpReplace($sHTML, "(?s)<!\[CDATA\[.*?\]\]>", "") ; removing CDATA
	Local $fSend = False ; preset 'Sending flag'
	; Find all forms on page
	Local $aForm = StringRegExp($sHTML, "(?si)<\s*form(?:[^\w])\s*(.*?)(?:(?:<\s*/form\s*>)|\Z)", 3)
	If @error Then Return SetError(1, 0, "") ; There are no forms available
	; Process input
	Local $fGetFormByName, $sFormName, $fGetFormByIndex, $fGetFormById, $iFormIndex
	Local $aSplitForm = StringSplit($sFormId, ":", 2)
	If @error Then ; like .getElementById(FormId)
		$fGetFormById = True
	Else
		If $aSplitForm[0] = "name" Then ; like .getElementsByName(FormName)
			$sFormName = $aSplitForm[1]
			$fGetFormByName = True
		ElseIf $aSplitForm[0] = "index" Then
			$iFormIndex = Number($aSplitForm[1])
			$fGetFormByIndex = True
		ElseIf $aSplitForm[0] = "id" Then ; like .getElementById(FormId)
			$sFormId = $aSplitForm[1]
			$fGetFormById = True
		Else ; like .getElementById(FormId)
			$sFormId = $aSplitForm[0]
			$fGetFormById = True
		EndIf
	EndIf
	; Variables
	Local $sForm, $sAttributes, $aInput
	Local $iNumParams = Ceiling(($iNumArgs - 2) / 2) - 1
	Local $sAddData, $sNewURL
	; Loop thru all forms on the page and find one that was specified
	For $iFormOrdinal = 0 To UBound($aForm) - 1
		If $fGetFormByIndex And $iFormOrdinal <> $iFormIndex Then ContinueLoop
		$sForm = $aForm[$iFormOrdinal]
		; Extract form attributes
		$sAttributes = StringRegExp($sForm, "(?s)(.*?)>", 3)
		If Not @error Then $sAttributes = StringRegExpReplace($sAttributes[0], "\v", " ")
		Local $sAction = "", $sAccept = "", $sEnctype = "", $sMethod = "", $sName = "", $sId = ""
		; Check set attributes
		$sId = __WinHttpAttribVal($sAttributes, "id")
		If $fGetFormById And $sFormId <> Default And $sId <> $sFormId Then ContinueLoop
		$sName = __WinHttpAttribVal($sAttributes, "name")
		If $fGetFormByName And $sFormName <> $sName Then ContinueLoop
		$sAction = __WinHttpHTMLDecode(__WinHttpAttribVal($sAttributes, "action"))
		$sAccept = __WinHttpAttribVal($sAttributes, "accept")
		$sEnctype = __WinHttpAttribVal($sAttributes, "enctype")
		$sMethod = __WinHttpAttribVal($sAttributes, "method")
		; Requested form is found. Set $fSend flag to true
		$fSend = True
		Local $sSpr1 = Chr(27), $sSpr2 = Chr(26)
		__WinHttpNormalizeForm($sForm, $sSpr1, $sSpr2)
		$aInput = StringRegExp($sForm, "(?si)<\h*(?:input|textarea|label|fieldset|legend|select|optgroup|option|button)\h*(.*?)/*\h*>", 3)
		If @error Then Return SetError(2, 0, "") ; invalid form
		; HTML5 allows for "formaction", "formenctype", "formmethod" submit-control attributes to be set. If such control is "clicked" then form's attributes needs updating/correcting
		__WinHttpHTML5FormAttribs($aDtas, $aFlds, $iNumParams, $aInput, $sAction, $sEnctype, $sMethod)
		; Workout correct URL, scheme, etc...
		__WinHttpNormalizeActionURL($sActionPage, $sAction, $iScheme, $sNewURL, $sEnctype, $sMethod, $sURL)
		If $fVarForm And Not $sNewURL Then Return SetError(5, 0, "") ; "action" must have URL specified
		Local $aSplit, $sBoundary, $sPassedId, $sPassedData, $iNumRepl, $fMultiPart = False, $sSubmit, $sRadio, $sCheckBox, $sButton
		Local $sGrSep = Chr(29), $iErr
		Local $aInputIds[4][UBound($aInput)]
		Switch $sEnctype
			Case "", "application/x-www-form-urlencoded", "text/plain"
				For $i = 0 To UBound($aInput) - 1 ; for all input elements
					__WinHttpFormAttrib($aInputIds, $i, $aInput[$i])
					If $aInputIds[1][$i] Then ; if there is 'name' field then add it
						$aInputIds[1][$i] = __WinHttpURLEncode(StringReplace($aInputIds[1][$i], $sSpr1, " "), $sEnctype)
						$aInputIds[2][$i] = __WinHttpURLEncode(StringReplace(StringReplace($aInputIds[2][$i], $sSpr2, ">"), $sSpr1, " "), $sEnctype)
						$sAddData &= $aInputIds[1][$i] & "=" & $aInputIds[2][$i] & "&"
						If $aInputIds[3][$i] = "submit" Then $sSubmit &= $aInputIds[1][$i] & "=" & $aInputIds[2][$i] & $sGrSep ; add to overall "submit" string
						If $aInputIds[3][$i] = "radio" Then $sRadio &= $aInputIds[1][$i] & "=" & $aInputIds[2][$i] & $sGrSep ; add to overall "radio" string
						If $aInputIds[3][$i] = "checkbox" Then $sCheckBox &= $aInputIds[1][$i] & "=" & $aInputIds[2][$i] & $sGrSep ; add to overall "checkbox" string
						If $aInputIds[3][$i] = "button" Then $sButton &= $aInputIds[1][$i] & "=" & $aInputIds[2][$i] & $sGrSep ; add to overall "button" string
					EndIf
				Next
				$sSubmit = StringTrimRight($sSubmit, 1)
				$sRadio = StringTrimRight($sRadio, 1)
				$sCheckBox = StringTrimRight($sCheckBox, 1)
				$sButton = StringTrimRight($sButton, 1)
				$sAddData = StringTrimRight($sAddData, 1)
				For $k = 1 To $iNumParams
					$sPassedData = __WinHttpURLEncode($aDtas[$k], $sEnctype)
					$aDtas[$k] = 0
					$sPassedId = $aFlds[$k]
					$aFlds[$k] = 0
					$aSplit = StringSplit($sPassedId, ":", 2)
					$iErr = @error
					$aSplit[0] = __WinHttpURLEncode($aSplit[0], $sEnctype)
					If Not $iErr Then $aSplit[1] = __WinHttpURLEncode($aSplit[1], $sEnctype)
					If $iErr Or $aSplit[0] <> "name" Then ; like .getElementById
						If Not $iErr And $aSplit[0] = "id" Then $sPassedId = $aSplit[1]
						For $j = 0 To UBound($aInputIds, 2) - 1
							If $aInputIds[0][$j] = $sPassedId Then
								If $aInputIds[3][$j] = "submit" Then
									If $sPassedData = True Then ; if this "submit" is set to TRUE then
										If $sSubmit Then ; If not already processed; only the first is valid
											Local $fDelId = False
											For $sChunkSub In StringSplit($sSubmit, $sGrSep, 3) ; go tru all "submit" controls
												If $sChunkSub == $aInputIds[1][$j] & "=" & $aInputIds[2][$j] Then
													If $fDelId Then $sAddData = StringRegExpReplace($sAddData, "(?:&|\A)\Q" & $sChunkSub & "\E(?:&|\Z)", "&", 1)
													$fDelId = True
												Else
													$sAddData = StringRegExpReplace($sAddData, "(?:&|\A)\Q" & $sChunkSub & "\E(?:&|\Z)", "&") ; delete all but the TRUE one
												EndIf
												__WinHttpTrimBounds($sAddData, "&")
											Next
											$sSubmit = ""
										EndIf
									EndIf
								ElseIf $aInputIds[3][$j] = "radio" Then
									If $sPassedData = $aInputIds[2][$j] Then
										For $sChunkSub In StringSplit($sRadio, $sGrSep, 3) ; go tru all "radio" controls
											If $sChunkSub == $aInputIds[1][$j] & "=" & $sPassedData Then
												$sAddData = StringRegExpReplace(StringReplace($sAddData, "&", "&&"), "(?:&|\A)\Q" & $aInputIds[1][$j] & "\E(.*?)(?:&|\Z)", "&")
												$sAddData = StringReplace(StringReplace($sAddData, "&&", "&"), "&&", "&")
												If StringLeft($sAddData, 1) = "&" Then $sAddData = StringTrimLeft($sAddData, 1)
												$sAddData &= "&" & $sChunkSub
												$sRadio = StringRegExpReplace(StringReplace($sRadio, $sGrSep, $sGrSep & $sGrSep), "(?:" & $sGrSep & "|\A)\Q" & $aInputIds[1][$j] & "\E(.*?)(?:" & $sGrSep & "|\Z)", $sGrSep)
												$sRadio = StringReplace(StringReplace($sRadio, $sGrSep & $sGrSep, $sGrSep), $sGrSep & $sGrSep, $sGrSep)
											EndIf
										Next
									EndIf
								ElseIf $aInputIds[3][$j] = "checkbox" Then
									$sCheckBox = StringRegExpReplace($sCheckBox, "\Q" & $aInputIds[1][$j] & "=" & $sPassedData & "\E" & $sGrSep & "*", "")
									__WinHttpTrimBounds($sCheckBox, $sGrSep)
								ElseIf $aInputIds[3][$j] = "button" Then
									$sButton = StringRegExpReplace($sButton, "\Q" & $aInputIds[1][$j] & "=" & $sPassedData & "\E" & $sGrSep & "*", "")
									__WinHttpTrimBounds($sButton, $sGrSep)
								Else
									$sAddData = StringRegExpReplace(StringReplace($sAddData, "&", "&&"), "(?:&|\A)\Q" & $aInputIds[1][$j] & "=" & $aInputIds[2][$j] & "\E(?:&|\Z)", "&" & $aInputIds[1][$j] & "=" & $sPassedData & "&")
									$iNumRepl = @extended
									$sAddData = StringReplace($sAddData, "&&", "&")
									If $iNumRepl > 1 Then ; equalize ; TODO: remove duplicates
										$sAddData = StringRegExpReplace($sAddData, "(?:&|\A)\Q" & $aInputIds[1][$j] & "\E=.*?(?:&|\Z)", "&", $iNumRepl - 1)
									EndIf
									__WinHttpTrimBounds($sAddData, "&")
								EndIf
							EndIf
						Next
					Else ; like .getElementsByName
						For $j = 0 To UBound($aInputIds, 2) - 1
							If $aInputIds[3][$j] = "submit" Then
								If $sPassedData = True Then ; if this "submit" is set to TRUE then
									If $aInputIds[1][$j] == $aSplit[1] Then
										If $sSubmit Then ; If not already processed; only the first is valid
											Local $fDel = False
											For $sChunkSub In StringSplit($sSubmit, $sGrSep, 3) ; go tru all "submit" controls
												If $sChunkSub = $aInputIds[1][$j] & "=" & $aInputIds[2][$j] Then
													If $fDel Then $sAddData = StringRegExpReplace($sAddData, "(?:&|\A)\Q" & $sChunkSub & "\E(?:&|\Z)", "&", 1)
													$fDel = True
												Else
													$sAddData = StringRegExpReplace($sAddData, "(?:&|\A)\Q" & $sChunkSub & "\E(?:&|\Z)", "&") ; delete all but the TRUE one
												EndIf
												__WinHttpTrimBounds($sAddData, "&")
											Next
											$sSubmit = ""
										EndIf
										ContinueLoop 2 ; process next parameter
									EndIf
								Else ; False means do nothing
									ContinueLoop 2 ; process next parameter
								EndIf
							ElseIf $aInputIds[1][$j] == $aSplit[1] And $aInputIds[3][$j] = "radio" Then
								For $sChunkSub In StringSplit($sRadio, $sGrSep, 3) ; go tru all "radio" controls
									If $sChunkSub == $aInputIds[1][$j] & "=" & $sPassedData Then
										$sAddData = StringReplace(StringReplace(StringRegExpReplace(StringReplace($sAddData, "&", "&&"), "(?:&|\A)\Q" & $aInputIds[1][$j] & "\E(.*?)(?:&|\Z)", "&"), "&&", "&"), "&&", "&")
										If StringLeft($sAddData, 1) = "&" Then $sAddData = StringTrimLeft($sAddData, 1)
										$sAddData &= "&" & $sChunkSub
										$sRadio = StringRegExpReplace(StringReplace($sRadio, $sGrSep, $sGrSep & $sGrSep), "(?:" & $sGrSep & "|\A)\Q" & $aInputIds[1][$j] & "\E(.*?)(?:" & $sGrSep & "|\Z)", $sGrSep)
										$sRadio = StringReplace(StringReplace($sRadio, $sGrSep & $sGrSep, $sGrSep), $sGrSep & $sGrSep, $sGrSep)
									EndIf
								Next
								ContinueLoop 2 ; process next parameter
							ElseIf $aInputIds[1][$j] == $aSplit[1] And $aInputIds[3][$j] = "checkbox" Then
								$sCheckBox = StringRegExpReplace($sCheckBox, "\Q" & $aInputIds[1][$j] & "=" & $sPassedData & "\E" & $sGrSep & "*", "")
								__WinHttpTrimBounds($sCheckBox, $sGrSep)
								ContinueLoop 2 ; process next parameter
							ElseIf $aInputIds[1][$j] == $aSplit[1] And $aInputIds[3][$j] = "button" Then
								$sButton = StringRegExpReplace($sButton, "\Q" & $aInputIds[1][$j] & "=" & $sPassedData & "\E" & $sGrSep & "*", "")
								__WinHttpTrimBounds($sButton, $sGrSep)
								ContinueLoop 2 ; process next parameter
							EndIf
						Next
						$sAddData = StringRegExpReplace(StringReplace($sAddData, "&", "&&"), "(?:&|\A)\Q" & $aSplit[1] & "\E=.*?(?:&|\Z)", "&" & $aSplit[1] & "=" & $sPassedData & "&")
						$iNumRepl = @extended
						$sAddData = StringReplace($sAddData, "&&", "&")
						If $iNumRepl > 1 Then ; remove duplicates
							$sAddData = StringRegExpReplace($sAddData, "(?:&|\A)\Q" & $aSplit[1] & "\E=.*?(?:&|\Z)", "&", $iNumRepl - 1)
						EndIf
						__WinHttpTrimBounds($sAddData, "&")
					EndIf
				Next
				__WinHttpFinalizeCtrls($sSubmit, $sRadio, $sCheckBox, $sButton, $sAddData, $sGrSep, "&")
				If $sMethod = "GET" Then
					$sAction &= "?" & $sAddData
					$sAddData = "" ; not to send as addition to the request (this is GET)
				EndIf
			Case "multipart/form-data"
				If $sMethod = "POST" Then ; can't be GET
					$fMultiPart = True
					; Define boundary line
					$sBoundary = StringFormat("%s%.5f", "----WinHttpBoundaryLine_", Random(10000, 99999))
					Local $sCDisp = 'Content-Disposition: form-data; name="'
					For $i = 0 To UBound($aInput) - 1 ; for all input elements
						__WinHttpFormAttrib($aInputIds, $i, $aInput[$i])
						If $aInputIds[1][$i] Then ; if there is 'name' field
							$aInputIds[1][$i] = StringReplace($aInputIds[1][$i], $sSpr1, " ")
							$aInputIds[2][$i] = StringReplace(StringReplace($aInputIds[2][$i], $sSpr2, ">"), $sSpr1, " ")
							If $aInputIds[3][$i] = "file" Then
								$sAddData &= "--" & $sBoundary & @CRLF & _
										$sCDisp & $aInputIds[1][$i] & '"; filename=""' & @CRLF & @CRLF & _
										$aInputIds[2][$i] & @CRLF
							Else
								$sAddData &= "--" & $sBoundary & @CRLF & _
										$sCDisp & $aInputIds[1][$i] & '"' & @CRLF & @CRLF & _
										$aInputIds[2][$i] & @CRLF
							EndIf
							If $aInputIds[3][$i] = "submit" Then $sSubmit &= "--" & $sBoundary & @CRLF & _
									$sCDisp & $aInputIds[1][$i] & '"' & @CRLF & @CRLF & _
									$aInputIds[2][$i] & @CRLF & $sGrSep
							If $aInputIds[3][$i] = "radio" Then $sRadio &= "--" & $sBoundary & @CRLF & _
									$sCDisp & $aInputIds[1][$i] & '"' & @CRLF & @CRLF & _
									$aInputIds[2][$i] & @CRLF & $sGrSep
							If $aInputIds[3][$i] = "checkbox" Then $sCheckBox &= "--" & $sBoundary & @CRLF & _
									$sCDisp & $aInputIds[1][$i] & '"' & @CRLF & @CRLF & _
									$aInputIds[2][$i] & @CRLF & $sGrSep
							If $aInputIds[3][$i] = "button" Then $sButton &= "--" & $sBoundary & @CRLF & _
									$sCDisp & $aInputIds[1][$i] & '"' & @CRLF & @CRLF & _
									$aInputIds[2][$i] & @CRLF & $sGrSep
						EndIf
					Next
					$sSubmit = StringTrimRight($sSubmit, 1)
					$sRadio = StringTrimRight($sRadio, 1)
					$sCheckBox = StringTrimRight($sCheckBox, 1)
					$sButton = StringTrimRight($sButton, 1)
					$sAddData &= "--" & $sBoundary & "--" & @CRLF
					For $k = 1 To $iNumParams
						$sPassedData = $aDtas[$k]
						$aDtas[$k] = 0
						$sPassedId = $aFlds[$k]
						$aFlds[$k] = 0
						$aSplit = StringSplit($sPassedId, ":", 2)
						If @error Or $aSplit[0] <> "name" Then ; like getElementById
							If Not @error And $aSplit[0] = "id" Then $sPassedId = $aSplit[1]
							For $j = 0 To UBound($aInputIds, 2) - 1
								If $aInputIds[0][$j] = $sPassedId Then
									If $aInputIds[3][$j] = "file" Then
										$sAddData = StringReplace($sAddData, _
												$sCDisp & $aInputIds[1][$j] & '"; filename=""' & @CRLF & @CRLF & $aInputIds[2][$j] & @CRLF, _
												__WinHttpFileContent($sAccept, $aInputIds[1][$j], $sPassedData, $sBoundary), 0, 1)
									ElseIf $aInputIds[3][$j] = "submit" Then
										If $sPassedData = True Then ; if this "submit" is set to TRUE then
											If $sSubmit Then ; If not already processed; only the first is valid
												Local $fMDelId = False
												For $sChunkSub In StringSplit($sSubmit, $sGrSep, 3) ; go tru all "submit" controls
													If $sChunkSub = "--" & $sBoundary & @CRLF & _
															$sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & _
															$aInputIds[2][$j] & @CRLF Then
														If $fMDelId Then $sAddData = StringReplace($sAddData, $sChunkSub, "", 1, 1) ; Removing duplicates
														$fMDelId = True
													Else
														$sAddData = StringReplace($sAddData, $sChunkSub, "", 0, 1) ; delete all but the TRUE one
													EndIf
												Next
												$sSubmit = ""
											EndIf
										EndIf
									ElseIf $aInputIds[3][$j] = "radio" Then
										If $sPassedData = $aInputIds[2][$j] Then
											For $sChunkSub In StringSplit($sRadio, $sGrSep, 3) ; go tru all "radio" controls
												If StringInStr($sChunkSub, "--" & $sBoundary & @CRLF & $sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & $sPassedData & @CRLF, 1) Then
													$sAddData = StringRegExpReplace($sAddData, "\Q--" & $sBoundary & @CRLF & $sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & "\E" & "(.*?)" & @CRLF, "")
													$sAddData = StringReplace($sAddData, "--" & $sBoundary & "--" & @CRLF, "", 0, 1)
													$sAddData &= $sChunkSub & "--" & $sBoundary & "--" & @CRLF
													$sRadio = StringRegExpReplace($sRadio, "\Q--" & $sBoundary & @CRLF & $sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & "\E(.*?)" & @CRLF & $sGrSep & "?", "")
												EndIf
											Next
										EndIf
									ElseIf $aInputIds[3][$j] = "checkbox" Then
										$sCheckBox = StringRegExpReplace($sCheckBox, "\Q--" & $sBoundary & @CRLF & _
												$sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & _
												$sPassedData & @CRLF & "\E" & $sGrSep & "*", "")
										If StringRight($sCheckBox, 1) = $sGrSep Then $sCheckBox = StringTrimRight($sCheckBox, 1)
									ElseIf $aInputIds[3][$j] = "button" Then
										$sButton = StringRegExpReplace($sButton, "\Q--" & $sBoundary & @CRLF & _
												$sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & _
												$sPassedData & @CRLF & "\E" & $sGrSep & "*", "")
										If StringRight($sButton, 1) = $sGrSep Then $sButton = StringTrimRight($sButton, 1)
									Else
										$sAddData = StringReplace($sAddData, _
												$sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & $aInputIds[2][$j] & @CRLF, _
												$sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & $sPassedData & @CRLF, 0, 1)
										$iNumRepl = @extended
										If $iNumRepl > 1 Then ; equalize ; TODO: remove duplicates
											$sAddData = StringRegExpReplace($sAddData, '(?s)\Q--' & $sBoundary & @CRLF & $sCDisp & $aInputIds[1][$j] & '"' & '\E\r\n\r\n.*?\r\n', "", $iNumRepl - 1)
										EndIf
									EndIf
								EndIf
							Next
						Else ; like getElementsByName
							For $j = 0 To UBound($aInputIds, 2) - 1
								If $aInputIds[1][$j] == $aSplit[1] And $aInputIds[3][$j] = "file" Then
									$sAddData = StringReplace($sAddData, _
											$sCDisp & $aSplit[1] & '"; filename=""' & @CRLF & @CRLF & $aInputIds[2][$j] & @CRLF, _
											__WinHttpFileContent($sAccept, $aInputIds[1][$j], $sPassedData, $sBoundary), 0, 1)
								ElseIf $aInputIds[1][$j] == $aSplit[1] And $aInputIds[3][$j] = "submit" Then
									If $sPassedData = True Then ; if this "submit" is set to TRUE then
										If $sSubmit Then ; If not already processed; only the first is valid
											Local $fMDel = False
											For $sChunkSub In StringSplit($sSubmit, $sGrSep, 3) ; go tru all "submit" controls
												If $sChunkSub = "--" & $sBoundary & @CRLF & _
														$sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & _
														$aInputIds[2][$j] & @CRLF Then
													If $fMDel Then $sAddData = StringReplace($sAddData, $sChunkSub, "", 1, 1) ; Removing duplicates
													$fMDel = True
												Else
													$sAddData = StringReplace($sAddData, $sChunkSub, "", 0, 1) ; delete all but the TRUE one
												EndIf
											Next
											$sSubmit = ""
										EndIf
										ContinueLoop 2 ; process next parameter
									Else ; False means do nothing
										ContinueLoop 2 ; process next parameter
									EndIf
								ElseIf $aInputIds[1][$j] == $aSplit[1] And $aInputIds[3][$j] = "radio" Then
									For $sChunkSub In StringSplit($sRadio, $sGrSep, 3) ; go tru all "radio" controls
										If StringInStr($sChunkSub, "--" & $sBoundary & @CRLF & $sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & $sPassedData & @CRLF, 1) Then
											$sAddData = StringRegExpReplace($sAddData, "\Q--" & $sBoundary & @CRLF & $sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & "\E" & "(.*?)" & @CRLF, "")
											$sAddData = StringReplace($sAddData, "--" & $sBoundary & "--" & @CRLF, "", 0, 1)
											$sAddData &= $sChunkSub & "--" & $sBoundary & "--" & @CRLF
											$sRadio = StringRegExpReplace($sRadio, "\Q--" & $sBoundary & @CRLF & $sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & "\E(.*?)" & @CRLF & $sGrSep & "?", "")
										EndIf
									Next
									ContinueLoop 2 ; process next parameter
								ElseIf $aInputIds[1][$j] == $aSplit[1] And $aInputIds[3][$j] = "checkbox" Then
									$sCheckBox = StringRegExpReplace($sCheckBox, "\Q--" & $sBoundary & @CRLF & _
											$sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & _
											$sPassedData & @CRLF & "\E" & $sGrSep & "*", "")
									If StringRight($sCheckBox, 1) = $sGrSep Then $sCheckBox = StringTrimRight($sCheckBox, 1)
									ContinueLoop 2 ; process next parameter
								ElseIf $aInputIds[1][$j] == $aSplit[1] And $aInputIds[3][$j] = "button" Then
									$sButton = StringRegExpReplace($sButton, "\Q--" & $sBoundary & @CRLF & _
											$sCDisp & $aInputIds[1][$j] & '"' & @CRLF & @CRLF & _
											$sPassedData & @CRLF & "\E" & $sGrSep & "*", "")
									If StringRight($sButton, 1) = $sGrSep Then $sButton = StringTrimRight($sButton, 1)
									ContinueLoop 2 ; process next parameter
								EndIf
							Next
							$sAddData = StringRegExpReplace($sAddData, '(?s)\Q' & $sCDisp & $aSplit[1] & '"' & '\E\r\n\r\n.*?\r\n', _
									$sCDisp & $aSplit[1] & '"' & @CRLF & @CRLF & StringReplace($sPassedData, "\", "\\") & @CRLF)
							$iNumRepl = @extended
							If $iNumRepl > 1 Then ; remove duplicates
								$sAddData = StringRegExpReplace($sAddData, '(?s)\Q--' & $sBoundary & @CRLF & $sCDisp & $aSplit[1] & '"' & '\E\r\n\r\n.*?\r\n', "", $iNumRepl - 1)
							EndIf
						EndIf
					Next
				EndIf
				__WinHttpFinalizeCtrls($sSubmit, $sRadio, $sCheckBox, $sButton, $sAddData, $sGrSep)
		EndSwitch
		ExitLoop
	Next
	; Send
	If $fSend Then
		If $fVarForm Then
			$hInternet = _WinHttpConnect($hOpen, $sNewURL)
		Else
			If $sNewURL Then
				$hOpen = _WinHttpQueryOption($hInternet, $WINHTTP_OPTION_PARENT_HANDLE)
				_WinHttpCloseHandle($hInternet)
				$hInternet = _WinHttpConnect($hOpen, $sNewURL)
			EndIf
		EndIf
		Local $hRequest
		If $iScheme = $INTERNET_SCHEME_HTTPS Then
			$hRequest = __WinHttpFormSend($hInternet, $sMethod, $sAction, $fMultiPart, $sBoundary, $sAddData, True, $sAdditionalHeaders, $sCredName, $sCredPass, $iIgnoreCertErr)
		Else
			$hRequest = __WinHttpFormSend($hInternet, $sMethod, $sAction, $fMultiPart, $sBoundary, $sAddData, False, $sAdditionalHeaders, $sCredName, $sCredPass)
			If _WinHttpQueryHeaders($hRequest, $WINHTTP_QUERY_STATUS_CODE) >= $HTTP_STATUS_BAD_REQUEST Then
				_WinHttpCloseHandle($hRequest)
				$hRequest = __WinHttpFormSend($hInternet, $sMethod, $sAction, $fMultiPart, $sBoundary, $sAddData, True, $sAdditionalHeaders, $sCredName, $sCredPass, $iIgnoreCertErr) ; try adding $WINHTTP_FLAG_SECURE
			EndIf
		EndIf
		Local $vReturned = _WinHttpSimpleReadData($hRequest)
		If @error Then
			_WinHttpCloseHandle($hRequest)
			Return SetError(4, 0, "") ; either site is expiriencing problems or your connection
		EndIf
		Local $iSCode = _WinHttpQueryHeaders($hRequest, $WINHTTP_QUERY_STATUS_CODE)
		If $iRetArr Then
			Local $aReturn[3] = [_WinHttpQueryHeaders($hRequest), $vReturned, _WinHttpQueryOption($hRequest, $WINHTTP_OPTION_URL)]
			$vReturned = $aReturn
		EndIf
		_WinHttpCloseHandle($hRequest)
		Return SetExtended($iSCode, $vReturned)
	EndIf
	; If here then there is no form on the page with specified attributes (name, id or index)
	Return SetError(3, 0, "")
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WinHttpSimpleFormFill_SetUploadCallback
; Description ...: Sets user defined function as callback function for _WinHttpSimpleFormFill
; Syntax.........: _WinHttpSimpleFormFill_SetUploadCallback($vCallback [, $iNumChunks = 100 ])
; Parameters ....: $vCallback - UDF's name
;                  $iNumChunks - [optional] number of chunks to send during form submitting. Default is 100.
; Return values .: Undefined.
; Author ........: trancexx
; Remarks .......: Unregistering is done by passing [[0]] as first argument.
; Related .......: _WinHttpSimpleFormFill
; ===============================================================================================================================
Func _WinHttpSimpleFormFill_SetUploadCallback($vCallback = Default, $iNumChunks = 100)
	If $iNumChunks <= 0 Then $iNumChunks = 100
	Local Static $vFunc = Default, $iParts = $iNumChunks
	If $vCallback <> Default Then
		$vFunc = $vCallback
		$iParts = Ceiling($iNumChunks)
	ElseIf $vCallback = 0 Then
		$vFunc = Default
		$iParts = 1
	EndIf
	Local $aOut[2] = [$vFunc, $iParts]
	Return $aOut
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WinHttpSimpleReadData
; Description ...: Reads data from a request
; Syntax.........: _WinHttpSimpleReadData($hRequest [, $iMode = Default ])
; Parameters ....: $hRequest - request handle after _WinHttpReceiveResponse
;                  $iMode         - [optional] type of data returned
;                  |0 - ASCII-String
;                  |1 - UTF-8-String
;                  |2 - binary data
; Return values .: Success      - String or binary depending on $iMode
;                  Failure      - empty string or empty binary (mode 2) and set @error
;                  |1 - invalid mode
;                  |2 - no data available
; Author ........: ProgAndy
; Modified.......: trancexx
; Remarks .......: In case of default reading mode, if the server specifies utf-8 content type, function will force UTF-8 string.
; Related .......: _WinHttpReadData, _WinHttpSimpleRequest, _WinHttpSimpleSSLRequest
; ===============================================================================================================================
Func _WinHttpSimpleReadData($hRequest, $iMode = Default)
	If $iMode = Default Then
		$iMode = 0
		If __WinHttpCharSet(_WinHttpQueryHeaders($hRequest, $WINHTTP_QUERY_CONTENT_TYPE)) = 65001 Then $iMode = 1 ; header suggest utf-8
	Else
		__WinHttpDefault($iMode, 0)
	EndIf
	If $iMode > 2 Or $iMode < 0 Then Return SetError(1, 0, '')
	Local $vData = Binary('')
	If _WinHttpQueryDataAvailable($hRequest) Then
		Do
			$vData &= _WinHttpReadData($hRequest, 2)
		Until @error
		Switch $iMode
			Case 0
				Return BinaryToString($vData)
			Case 1
				Return BinaryToString($vData, 4)
			Case Else
				Return $vData
		EndSwitch
	EndIf
	Return SetError(2, 0, $vData)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WinHttpSimpleReadDataAsync
; Description ...: Reads data from a request in asynchronous mode
; Syntax.........: _WinHttpSimpleReadDataAsync($hInternet, Byref $pBuffer [, $iNumberOfBytesToRead = Default ])
; Parameters ....: $hInternet - Request handle (first parameter while in callback function).
;                  $pBuffer - Pointer to memory buffer to which to read.
;                  $iNumberOfBytesToRead - [optional] The number of bytes to read. Default is 8192 bytes.
;                  |0 - ASCII-String
;                  |1 - UTF-8-String
;                  |2 - binary data
; Return values .: Same as for _WinHttpReadData. Due to async nature here it has no meaning except in case of possible error.
; Author ........: trancexx
; Remarks .......: <b>You are strongly discouraged to use WinHTTP in asynchronous mode with AutoIt. AutoIt's callback implementation can't handle reentrancy properly.</b>
;                  +WinHttp is rentrant during asynchronous completion callback. Make sure you have only one callback running and only one request handled though it at time.
;                  +Also make sure memory buffer is at least 8192 bytes in size if [[$iNumberOfBytesToRead]] is left default.
; Related .......: _WinHttpSimpleReadData, _WinHttpReadData
; ===============================================================================================================================
Func _WinHttpSimpleReadDataAsync($hInternet, ByRef $pBuffer, $iNumberOfBytesToRead = Default)
	__WinHttpDefault($iNumberOfBytesToRead, 8192)
	Local $vOut = _WinHttpReadData($hInternet, 2, $iNumberOfBytesToRead, $pBuffer)
	Return SetError(@error, @extended, $vOut)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WinHttpSimpleRequest
; Description ...: A function to send a request in a simpler form
; Syntax.........: _WinHttpSimpleRequest($hConnect, $sType, $sPath [, $sReferrer = Default [, $sDta = Default [, $sHeader = Default [, $fGetHeaders = Default [, $iMode = Default ]]]]])
; Parameters ....: $hConnect  - Handle from _WinHttpConnect
;                  $sType       - [optional] GET or POST (default: GET)
;                  $sPath       - [optional] request path (default: "" - empty string; meaning 'default' page on the server)
;                  $sReferrer   - [optional] referrer (default: $WINHTTP_NO_REFERER)
;                  $sDta        - [optional] POST-Data (default: $WINHTTP_NO_REQUEST_DATA)
;                  $sHeader     - [optional] additional Headers (default: $WINHTTP_NO_ADDITIONAL_HEADERS)
;                  $fGetHeaders - [optional] return response headers (default: False)
;                  $iMode       - [optional] reading mode of result
;                  |0 - ASCII-text
;                  |1 - UTF-8 text
;                  |2 - binary data
; Return values .: Success      - response data if $fGetHeaders = False (default)
;                  |Array if $fGetHeaders = True
;                  | [0] - response headers
;                  | [1] - response data
;                  Failure      - 0 and set @error
;                  |1 - could not open request
;                  |2 - could not send request
;                  |3 - could not receive response
;                  |4 - $iMode is not valid
; Author ........: ProgAndy
; Modified.......: trancexx
; Related .......: _WinHttpSimpleSSLRequest, _WinHttpSimpleSendRequest, _WinHttpSimpleSendSSLRequest, _WinHttpQueryHeaders, _WinHttpSimpleReadData
; ===============================================================================================================================
Func _WinHttpSimpleRequest($hConnect, $sType = Default, $sPath = Default, $sReferrer = Default, $sDta = Default, $sHeader = Default, $fGetHeaders = Default, $iMode = Default, $sCredName = Default, $sCredPass = Default)
	; Author: ProgAndy
	__WinHttpDefault($sType, "GET")
	__WinHttpDefault($sPath, "")
	__WinHttpDefault($sReferrer, $WINHTTP_NO_REFERER)
	__WinHttpDefault($sDta, $WINHTTP_NO_REQUEST_DATA)
	__WinHttpDefault($sHeader, $WINHTTP_NO_ADDITIONAL_HEADERS)
	__WinHttpDefault($fGetHeaders, False)
	__WinHttpDefault($iMode, Default)
	__WinHttpDefault($sCredName, "")
	__WinHttpDefault($sCredPass, "")
	If $iMode > 2 Or $iMode < 0 Then Return SetError(4, 0, 0)
	Local $hRequest = _WinHttpSimpleSendRequest($hConnect, $sType, $sPath, $sReferrer, $sDta, $sHeader)
	If @error Then Return SetError(@error, 0, 0)
	__WinHttpSetCredentials($hRequest, $sHeader, $sDta, $sCredName, $sCredPass)
	Local $iExtended = _WinHttpQueryHeaders($hRequest, $WINHTTP_QUERY_STATUS_CODE)
	If $fGetHeaders Then
		Local $aData[3] = [_WinHttpQueryHeaders($hRequest), _WinHttpSimpleReadData($hRequest, $iMode), _WinHttpQueryOption($hRequest, $WINHTTP_OPTION_URL)]
		_WinHttpCloseHandle($hRequest)
		Return SetExtended($iExtended, $aData)
	EndIf
	Local $sOutData = _WinHttpSimpleReadData($hRequest, $iMode)
	_WinHttpCloseHandle($hRequest)
	Return SetExtended($iExtended, $sOutData)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WinHttpSimpleSendRequest
; Description ...: A function to send a request in a simpler form, but not read the data
; Syntax.........: _WinHttpSimpleSendRequest($hConnect, $sType, $sPath [, $sReferrer = Default [, $sDta = Default [, $sHeader = Default ]]])
; Parameters ....: $hConnect  - Handle from _WinHttpConnect
;                  $sType       - [optional] GET or POST (default: GET)
;                  $sPath       - [optional] request path (default: "" - empty string; meaning 'default' page on the server)
;                  $sReferrer   - [optional] referrer (default: $WINHTTP_NO_REFERER)
;                  $sDta        - [optional] POST-Data (default: $WINHTTP_NO_REQUEST_DATA)
;                  $sHeader     - [optional] additional Headers (default: $WINHTTP_NO_ADDITIONAL_HEADERS)
; Return values .: Success      - handle of request after _WinHttpReceiveResponse.
;                  Failure      - 0 and set @error
;                  |1 - could not open request
;                  |2 - could not send request
;                  |3 - could not receive response
; Author ........: ProgAndy
; Related .......: _WinHttpSimpleRequest, _WinHttpSimpleSendSSLRequest, _WinHttpSimpleReadData
; ===============================================================================================================================
Func _WinHttpSimpleSendRequest($hConnect, $sType = Default, $sPath = Default, $sReferrer = Default, $sDta = Default, $sHeader = Default)
	; Author: ProgAndy
	__WinHttpDefault($sType, "GET")
	__WinHttpDefault($sPath, "")
	__WinHttpDefault($sReferrer, $WINHTTP_NO_REFERER)
	__WinHttpDefault($sDta, $WINHTTP_NO_REQUEST_DATA)
	__WinHttpDefault($sHeader, $WINHTTP_NO_ADDITIONAL_HEADERS)
	Local $hRequest = _WinHttpOpenRequest($hConnect, $sType, $sPath, Default, $sReferrer)
	If Not $hRequest Then Return SetError(1, @error, 0)
	If $sType = "POST" And $sHeader = $WINHTTP_NO_ADDITIONAL_HEADERS Then $sHeader = "Content-Type: application/x-www-form-urlencoded" & @CRLF
	_WinHttpSetOption($hRequest, $WINHTTP_OPTION_DECOMPRESSION, $WINHTTP_DECOMPRESSION_FLAG_ALL)
	_WinHttpSetOption($hRequest, $WINHTTP_OPTION_UNSAFE_HEADER_PARSING, 1)
	_WinHttpSendRequest($hRequest, $sHeader, $sDta)
	If @error Then Return SetError(2, 0 * _WinHttpCloseHandle($hRequest), 0)
	_WinHttpReceiveResponse($hRequest)
	If @error Then Return SetError(3, 0 * _WinHttpCloseHandle($hRequest), 0)
	Return $hRequest
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WinHttpSimpleSendSSLRequest
; Description ...: A function to send a SSL request in a simpler form, but not read the data
; Syntax.........: _WinHttpSimpleSendSSLRequest($hConnect [, $sType [, $sPath [, $sReferrer = Default [, $sDta = Default [, $sHeader = Default ]]]]])
; Parameters ....: $hConnect  - Handle from _WinHttpConnect
;                  $sType       - [optional] GET or POST (default: GET)
;                  $sPath       - [optional] request path (default: "" - empty string; meaning 'default' page on the server)
;                  $sReferrer   - [optional] referrer (default: $WINHTTP_NO_REFERER)
;                  $sDta        - [optional] POST-Data (default: $WINHTTP_NO_REQUEST_DATA)
;                  $sHeader     - [optional] additional Headers (default: $WINHTTP_NO_ADDITIONAL_HEADERS)
; Return values .: Success      - handle of request after _WinHttpReceiveResponse.
;                  Failure      - 0 and set @error
;                  |1 - could not open request
;                  |2 - could not send request
;                  |3 - could not receive response
; Author ........: ProgAndy
; Related .......: _WinHttpSimpleSSLRequest, _WinHttpSimpleSendRequest, _WinHttpSimpleReadData
; ===============================================================================================================================
Func _WinHttpSimpleSendSSLRequest($hConnect, $sType = Default, $sPath = Default, $sReferrer = Default, $sDta = Default, $sHeader = Default, $iIgnoreAllCertErrors = 0)
	; Author: ProgAndy
	__WinHttpDefault($sType, "GET")
	__WinHttpDefault($sPath, "")
	__WinHttpDefault($sReferrer, $WINHTTP_NO_REFERER)
	__WinHttpDefault($sDta, $WINHTTP_NO_REQUEST_DATA)
	__WinHttpDefault($sHeader, $WINHTTP_NO_ADDITIONAL_HEADERS)
	Local $hRequest = _WinHttpOpenRequest($hConnect, $sType, $sPath, Default, $sReferrer, Default, BitOR($WINHTTP_FLAG_SECURE, $WINHTTP_FLAG_ESCAPE_DISABLE))
	If Not $hRequest Then Return SetError(1, @error, 0)
	If $iIgnoreAllCertErrors Then _WinHttpSetOption($hRequest, $WINHTTP_OPTION_SECURITY_FLAGS, BitOR($SECURITY_FLAG_IGNORE_UNKNOWN_CA, $SECURITY_FLAG_IGNORE_CERT_DATE_INVALID, $SECURITY_FLAG_IGNORE_CERT_CN_INVALID, $SECURITY_FLAG_IGNORE_CERT_WRONG_USAGE))
	If $sType = "POST" And $sHeader = $WINHTTP_NO_ADDITIONAL_HEADERS Then $sHeader = "Content-Type: application/x-www-form-urlencoded" & @CRLF
	_WinHttpSetOption($hRequest, $WINHTTP_OPTION_DECOMPRESSION, $WINHTTP_DECOMPRESSION_FLAG_ALL)
	_WinHttpSetOption($hRequest, $WINHTTP_OPTION_UNSAFE_HEADER_PARSING, 1)
	_WinHttpSendRequest($hRequest, $sHeader, $sDta)
	If @error Then Return SetError(2, 0 * _WinHttpCloseHandle($hRequest), 0)
	_WinHttpReceiveResponse($hRequest)
	If @error Then Return SetError(3, 0 * _WinHttpCloseHandle($hRequest), 0)
	Return $hRequest
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _WinHttpSimpleSSLRequest
; Description ...: A function to send a SSL request in a simpler form
; Syntax.........: _WinHttpSimpleSSLRequest($hConnect [, $sType [, $sPath [, $sReferrer = Default [, $sDta = Default [, $sHeader = Default [, $fGetHeaders = Default [, $iMode = Default ]]]]]]])
; Parameters ....: $hConnect  - Handle from _WinHttpConnect
;                  $sType       - [optional] GET or POST (default: GET)
;                  $sPath       - [optional] request path (default: "" - empty string; meaning 'default' page on the server)
;                  $sReferrer   - [optional] referrer (default: $WINHTTP_NO_REFERER)
;                  $sDta        - [optional] POST-Data (default: $WINHTTP_NO_REQUEST_DATA)
;                  $sHeader     - [optional] additional Headers (default: $WINHTTP_NO_ADDITIONAL_HEADERS)
;                  $fGetHeaders - [optional] return response headers (default: False)
;                  $iMode       - [optional] reading mode of result
;                  |0 - ASCII-text
;                  |1 - UTF-8 text
;                  |2 - binary data
; Return values .: Success      - response data if $fGetHeaders = False (default)
;                  |Array if $fGetHeaders = True
;                  | [0] - response headers
;                  | [1] - response data
;                  Failure      - 0 and set @error
;                  |1 - could not open request
;                  |2 - could not send request
;                  |3 - could not receive response
;                  |4 - $iMode is not valid
; Author ........: ProgAndy
; Modified.......: trancexx
; Related .......: _WinHttpSimpleRequest, _WinHttpSimpleSendSSLRequest, _WinHttpSimpleSendRequest, _WinHttpQueryHeaders, _WinHttpSimpleReadData
; ===============================================================================================================================
Func _WinHttpSimpleSSLRequest($hConnect, $sType = Default, $sPath = Default, $sReferrer = Default, $sDta = Default, $sHeader = Default, $fGetHeaders = Default, $iMode = Default, $sCredName = Default, $sCredPass = Default, $iIgnoreCertErrors = 0)
	; Author: ProgAndy
	__WinHttpDefault($sType, "GET")
	__WinHttpDefault($sPath, "")
	__WinHttpDefault($sReferrer, $WINHTTP_NO_REFERER)
	__WinHttpDefault($sDta, $WINHTTP_NO_REQUEST_DATA)
	__WinHttpDefault($sHeader, $WINHTTP_NO_ADDITIONAL_HEADERS)
	__WinHttpDefault($fGetHeaders, False)
	__WinHttpDefault($iMode, Default)
	__WinHttpDefault($sCredName, "")
	__WinHttpDefault($sCredPass, "")
	If $iMode > 2 Or $iMode < 0 Then Return SetError(4, 0, 0)
	Local $hRequest = _WinHttpSimpleSendSSLRequest($hConnect, $sType, $sPath, $sReferrer, $sDta, $sHeader, $iIgnoreCertErrors)
	If @error Then Return SetError(@error, 0, 0)
	__WinHttpSetCredentials($hRequest, $sHeader, $sDta, $sCredName, $sCredPass)
	If $fGetHeaders Then
		Local $aData[3] = [_WinHttpQueryHeaders($hRequest), _WinHttpSimpleReadData($hRequest, $iMode), _WinHttpQueryOption($hRequest, $WINHTTP_OPTION_URL)]
		_WinHttpCloseHandle($hRequest)
		Return $aData
	EndIf
	Local $sOutData = _WinHttpSimpleReadData($hRequest, $iMode)
	_WinHttpCloseHandle($hRequest)
	Return $sOutData
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpTimeFromSystemTime
; Description ...: Formats a system date and time according to the HTTP version 1.0 specification.
; Syntax.........: _WinHttpTimeFromSystemTime()
; Parameters ....: None.
; Return values .: Success - Returns time string.
;                  Failure - Returns empty string and sets @error:
;                  |1 - Initial DllCall failed
;                  |2 - Main DllCall failed
; Author ........: trancexx
; Related .......: _WinHttpTimeToSystemTime
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384117.aspx
;============================================================================================
Func _WinHttpTimeFromSystemTime()
	Local $SYSTEMTIME = DllStructCreate("word Year;" & _
			"word Month;" & _
			"word DayOfWeek;" & _
			"word Day;" & _
			"word Hour;" & _
			"word Minute;" & _
			"word Second;" & _
			"word Milliseconds")
	DllCall("kernel32.dll", "none", "GetSystemTime", "struct*", $SYSTEMTIME)
	If @error Then Return SetError(1, 0, "")
	Local $tTime = DllStructCreate("wchar[62]")
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpTimeFromSystemTime", "struct*", $SYSTEMTIME, "struct*", $tTime)
	If @error Or Not $aCall[0] Then Return SetError(2, 0, "")
	Return DllStructGetData($tTime, 1)
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpTimeToSystemTime
; Description ...: Takes an HTTP time/date string and converts it to array (SYSTEMTIME structure values).
; Syntax.........: _WinHttpTimeToSystemTime($sHttpTime)
; Parameters ....: $sHttpTime - Date/time string to convert.
; Return values .: Success - Returns array with 8 elements:
;                  |$array[0] - Year,
;                  |$array[1] - Month,
;                  |$array[2] - DayOfWeek,
;                  |$array[3] - Day,
;                  |$array[4] - Hour,
;                  |$array[5] - Minute,
;                  |$array[6] - Second.,
;                  |$array[7] - Milliseconds.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx
; Related .......: _WinHttpTimeFromSystemTime
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384118.aspx
;============================================================================================
Func _WinHttpTimeToSystemTime($sHttpTime)
	Local $SYSTEMTIME = DllStructCreate("word Year;" & _
			"word Month;" & _
			"word DayOfWeek;" & _
			"word Day;" & _
			"word Hour;" & _
			"word Minute;" & _
			"word Second;" & _
			"word Milliseconds")
	Local $tTime = DllStructCreate("wchar[62]")
	DllStructSetData($tTime, 1, $sHttpTime)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpTimeToSystemTime", "struct*", $tTime, "struct*", $SYSTEMTIME)
	If @error Or Not $aCall[0] Then Return SetError(2, 0, 0)
	Local $aRet[8] = [DllStructGetData($SYSTEMTIME, "Year"), _
			DllStructGetData($SYSTEMTIME, "Month"), _
			DllStructGetData($SYSTEMTIME, "DayOfWeek"), _
			DllStructGetData($SYSTEMTIME, "Day"), _
			DllStructGetData($SYSTEMTIME, "Hour"), _
			DllStructGetData($SYSTEMTIME, "Minute"), _
			DllStructGetData($SYSTEMTIME, "Second"), _
			DllStructGetData($SYSTEMTIME, "Milliseconds")]
	Return $aRet
EndFunc

; #FUNCTION# ;===============================================================================
; Name...........: _WinHttpWriteData
; Description ...: Writes request data to an HTTP server.
; Syntax.........: _WinHttpWriteData($hRequest, $vData [, $iMode = Default ])
; Parameters ....: $hRequest - Valid handle returned by _WinHttpSendRequest().
;                  $vData - Data to write.
;                  $iMode - [optional] Integer representing writing mode. Default is 0 - write ANSI string.
; Return values .: Success - Returns 1
;                          - @extended receives the number of bytes written.
;                  Failure - Returns 0 and sets @error:
;                  |1 - DllCall failed
; Author ........: trancexx, ProgAndy
; Remarks .......: [[$vData]] variable is either string or binary data to write.
;                  [[$iMode]] can have these values:
;                  |[[0]] - to write ANSI string
;                  |[[1]] - to write binary data
; Related .......: _WinHttpSendRequest, _WinHttpReadData
; Link ..........: http://msdn.microsoft.com/en-us/library/aa384120.aspx
;============================================================================================
Func _WinHttpWriteData($hRequest, $vData, $iMode = Default)
	__WinHttpDefault($iMode, 0)
	Local $iNumberOfBytesToWrite, $tData
	If $iMode = 1 Then
		$iNumberOfBytesToWrite = BinaryLen($vData)
		$tData = DllStructCreate("byte[" & $iNumberOfBytesToWrite & "]")
		DllStructSetData($tData, 1, $vData)
	ElseIf IsDllStruct($vData) Then
		$iNumberOfBytesToWrite = DllStructGetSize($vData)
		$tData = $vData
	Else
		$iNumberOfBytesToWrite = StringLen($vData)
		$tData = DllStructCreate("char[" & $iNumberOfBytesToWrite + 1 & "]")
		DllStructSetData($tData, 1, $vData)
	EndIf
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpWriteData", _
			"handle", $hRequest, _
			"struct*", $tData, _
			"dword", $iNumberOfBytesToWrite, _
			"dword*", 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return SetExtended($aCall[4], 1)
EndFunc


; #INTERNAL FUNCTIONS# ;=====================================================================
Func __WinHttpFileContent($sAccept, $sName, $sFileString, $sBoundaryMain = "")
	#forceref $sAccept ; FIXME: $sAccept is specified by the server (or left default). In case $sFileString is non-supported MIME type action should be aborted.
	Local $fNonStandard = False
	If StringLeft($sFileString, 10) = "PHP#50338:" Then
		$sFileString = StringTrimLeft($sFileString, 10)
		$fNonStandard = True
	EndIf
	Local $sOut = 'Content-Disposition: form-data; name="' & $sName & '"'
	If Not $sFileString Then Return $sOut & '; filename=""' & @CRLF & @CRLF & @CRLF
	; Check $sFileString string
	If StringRight($sFileString, 1) = "|" Then $sFileString = StringTrimRight($sFileString, 1)
	Local $aFiles = StringSplit($sFileString, "|", 2), $hFile
	If UBound($aFiles) = 1 Then
		$hFile = FileOpen($aFiles[0], 16)
		$sOut &= '; filename="' & StringRegExpReplace($aFiles[0], ".*\\", "") & '"' & @CRLF & _
				"Content-Type: " & __WinHttpMIMEType($aFiles[0]) & @CRLF & @CRLF & BinaryToString(FileRead($hFile), 1) & @CRLF
		FileClose($hFile)
		Return $sOut ; That's it
	EndIf
	; Multiple files specified, separated by "|". Support on server side required!
	If $fNonStandard Then
		; This way is forced by some browsers to avoid PHP's inability to parse multipart/mixed content
		$sOut = "" ; discharge
		Local $iFiles = UBound($aFiles)
		For $i = 0 To $iFiles - 1
			$hFile = FileOpen($aFiles[$i], 16)
			$sOut &= 'Content-Disposition: form-data; name="' & $sName & '"' & _
					'; filename="' & StringRegExpReplace($aFiles[$i], ".*\\", "") & '"' & @CRLF & _
					"Content-Type: " & __WinHttpMIMEType($aFiles[$i]) & @CRLF & @CRLF & BinaryToString(FileRead($hFile), 1) & @CRLF
			FileClose($hFile)
			If $i < $iFiles - 1 Then $sOut &= "--" & $sBoundaryMain & @CRLF
		Next
	Else
		; RFC2388 ( http://www.ietf.org/rfc/rfc2388.txt )
		Local $sBoundary = StringFormat("%s%.5f", "----WinHttpSubBoundaryLine_", Random(10000, 99999))
		$sOut &= @CRLF & "Content-Type: multipart/mixed; boundary=" & $sBoundary & @CRLF & @CRLF
		For $i = 0 To UBound($aFiles) - 1
			$hFile = FileOpen($aFiles[$i], 16)
			$sOut &= "--" & $sBoundary & @CRLF & _
					'Content-Disposition: file; filename="' & StringRegExpReplace($aFiles[$i], ".*\\", "") & '"' & @CRLF & _
					"Content-Type: " & __WinHttpMIMEType($aFiles[$i]) & @CRLF & @CRLF & BinaryToString(FileRead($hFile), 1) & @CRLF
			FileClose($hFile)
		Next
		$sOut &= "--" & $sBoundary & "--" & @CRLF
	EndIf
	Return $sOut
EndFunc

Func __WinHttpMIMEType($sFileName)
	Local $aArray = StringRegExp(__WinHttpMIMEAssocString(), "(?i)\Q;" & StringRegExpReplace($sFileName, ".*\.", "") & "\E\|(.*?);", 3)
	If @error Then Return "application/octet-stream"
	Return $aArray[0]
EndFunc

Func __WinHttpMIMEAssocString()
	Return ";ai|application/postscript;aif|audio/x-aiff;aifc|audio/x-aiff;aiff|audio/x-aiff;asc|text/plain;atom|application/atom+xml;au|audio/basic;avi|video/x-msvideo;bcpio|application/x-bcpio;bin|application/octet-stream;bmp|image/bmp;cdf|application/x-netcdf;cgm|image/cgm;class|application/octet-stream;cpio|application/x-cpio;cpt|application/mac-compactpro;csh|application/x-csh;css|text/css;dcr|application/x-director;dif|video/x-dv;dir|application/x-director;djv|image/vnd.djvu;djvu|image/vnd.djvu;dll|application/octet-stream;dmg|application/octet-stream;dms|application/octet-stream;doc|application/msword;dtd|application/xml-dtd;dv|video/x-dv;dvi|application/x-dvi;dxr|application/x-director;eps|application/postscript;etx|text/x-setext;exe|application/octet-stream;ez|application/andrew-inset;gif|image/gif;gram|application/srgs;grxml|application/srgs+xml;gtar|application/x-gtar;hdf|application/x-hdf;hqx|application/mac-binhex40;htm|text/html;html|text/html;ice|x-conference/x-cooltalk;ico|image/x-icon;ics|text/calendar;ief|image/ief;ifb|text/calendar;iges|model/iges;igs|model/iges;jnlp|application/x-java-jnlp-file;jp2|image/jp2;jpe|image/jpeg;jpeg|image/jpeg;jpg|image/jpeg;js|application/x-javascript;kar|audio/midi;latex|application/x-latex;lha|application/octet-stream;lzh|application/octet-stream;m3u|audio/x-mpegurl;m4a|audio/mp4a-latm;m4b|audio/mp4a-latm;m4p|audio/mp4a-latm;m4u|video/vnd.mpegurl;m4v|video/x-m4v;mac|image/x-macpaint;man|application/x-troff-man;mathml|application/mathml+xml;me|application/x-troff-me;mesh|model/mesh;mid|audio/midi;midi|audio/midi;mif|application/vnd.mif;mov|video/quicktime;movie|video/x-sgi-movie;mp2|audio/mpeg;mp3|audio/mpeg;mp4|video/mp4;mpe|video/mpeg;mpeg|video/mpeg;mpg|video/mpeg;mpga|audio/mpeg;ms|application/x-troff-ms;msh|model/mesh;mxu|video/vnd.mpegurl;nc|application/x-netcdf;oda|application/oda;ogg|application/ogg;pbm|image/x-portable-bitmap;pct|image/pict;pdb|chemical/x-pdb;pdf|application/pdf;pgm|image/x-portable-graymap;pgn|application/x-chess-pgn;pic|image/pict;pict|image/pict;png|image/png;pnm|image/x-portable-anymap;pnt|image/x-macpaint;pntg|image/x-macpaint;ppm|image/x-portable-pixmap;ppt|application/vnd.ms-powerpoint;ps|application/postscript;qt|video/quicktime;qti|image/x-quicktime;qtif|image/x-quicktime;ra|audio/x-pn-realaudio;ram|audio/x-pn-realaudio;ras|image/x-cmu-raster;rdf|application/rdf+xml;rgb|image/x-rgb;rm|application/vnd.rn-realmedia;roff|application/x-troff;rtf|text/rtf;rtx|text/richtext;sgm|text/sgml;sgml|text/sgml;sh|application/x-sh;shar|application/x-shar;silo|model/mesh;sit|application/x-stuffit;skd|application/x-koan;skm|application/x-koan;skp|application/x-koan;skt|application/x-koan;smi|application/smil;smil|application/smil;snd|audio/basic;so|application/octet-stream;spl|application/x-futuresplash;src|application/x-wais-source;sv4cpio|application/x-sv4cpio;sv4crc|application/x-sv4crc;svg|image/svg+xml;swf|application/x-shockwave-flash;t|application/x-troff;tar|application/x-tar;tcl|application/x-tcl;tex|application/x-tex;texi|application/x-texinfo;texinfo|application/x-texinfo;tif|image/tiff;tiff|image/tiff;tr|application/x-troff;tsv|text/tab-separated-values;txt|text/plain;ustar|application/x-ustar;vcd|application/x-cdlink;vrml|model/vrml;vxml|application/voicexml+xml;wav|audio/x-wav;wbmp|image/vnd.wap.wbmp;wbmxl|application/vnd.wap.wbxml;wml|text/vnd.wap.wml;wmlc|application/vnd.wap.wmlc;wmls|text/vnd.wap.wmlscript;wmlsc|application/vnd.wap.wmlscriptc;wrl|model/vrml;xbm|image/x-xbitmap;xht|application/xhtml+xml;xhtml|application/xhtml+xml;xls|application/vnd.ms-excel;xml|application/xml;xpm|image/x-xpixmap;xsl|application/xml;xslt|application/xslt+xml;xul|application/vnd.mozilla.xul+xml;xwd|image/x-xwindowdump;xyz|chemical/x-xyz;zip|application/zip;"
EndFunc

Func __WinHttpCharSet($sContentType)
	Local $aContentType = StringRegExp($sContentType, "(?i).*?\Qcharset=\E(?U)([^ ]+)(;| |\Z)", 2)
	If Not @error Then $sContentType = $aContentType[1]
	If StringLeft($sContentType, 2) = "cp" Then Return Int(StringTrimLeft($sContentType, 2))
	If $sContentType = "utf-8" Then Return 65001
EndFunc

Func __WinHttpURLEncode($vData, $sEncType = "")
	If IsBool($vData) Then Return $vData
	$vData = __WinHttpHTMLDecode($vData)
	If $sEnctype = "text/plain" Then Return StringReplace($vData, " ", "+")
	Local $aData = StringToASCIIArray($vData, Default, Default, 2)
	Local $sOut
	For $i = 0 To UBound($aData) - 1
		Switch $aData[$i]
			Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
				$sOut &= Chr($aData[$i])
			Case 32
				$sOut &= "+"
			Case Else
				$sOut &= "%" & Hex($aData[$i], 2)
		EndSwitch
	Next
	Return $sOut
EndFunc

Func __WinHttpHTMLDecode($vData)
	Return StringReplace(StringReplace(StringReplace(StringReplace($vData, "&amp;", "&"), "&lt;", "<"), "&gt;", ">"), "&quot;", '"')
EndFunc

Func __WinHttpNormalizeActionURL($sActionPage, ByRef $sAction, ByRef $iScheme, ByRef $sNewURL, ByRef $sEnctype, ByRef $sMethod, $sURL = "")
	Local $aCrackURL = _WinHttpCrackUrl($sAction)
	If @error Then
		If $sAction Then
			If StringLeft($sAction, 2) = "//" Then
				$aCrackURL = _WinHttpCrackUrl($sURL)
				If Not @error Then
					$aCrackURL = _WinHttpCrackUrl($aCrackURL[0] & ":" & $sAction)
					If Not @error Then
						$sNewURL = $aCrackURL[0] & "://" & $aCrackURL[2] & ":" & $aCrackURL[3]
						$iScheme = $aCrackURL[1]
						$sAction = $aCrackURL[6] & $aCrackURL[7]
						$sActionPage = ""
					EndIf
				EndIf
			ElseIf StringLeft($sAction, 1) = "?" Then
				$aCrackURL = _WinHttpCrackUrl($sURL)
				$sAction = $aCrackURL[6] & $sAction
			ElseIf StringLeft($sAction, 1) = "#" Then
				$sAction = StringReplace($sActionPage, StringRegExpReplace($sActionPage, "(.*?)(#.*?)", "$2"), $sAction)
			ElseIf StringLeft($sAction, 1) <> "/" Then
					Local $sCurrent
					Local $aURL = StringRegExp($sActionPage, '(.*)/', 3)
					If Not @error Then $sCurrent = $aURL[0]
					If $sCurrent Then $sAction = $sCurrent & "/" & $sAction
			EndIf
			If StringLeft($sAction, 1) = "?" Then $sAction = $sActionPage & $sAction
		EndIf
		If Not $sAction Then $sAction = $sActionPage
		$sAction = StringRegExpReplace($sAction, "\A(/*\.\./)*", "") ; /../
	Else
		$iScheme = $aCrackURL[1]
		$sNewURL = $aCrackURL[0] & "://" & $aCrackURL[2] & ":" & $aCrackURL[3]
		$sAction = $aCrackURL[6] & $aCrackURL[7]
	EndIf
	If Not $sMethod Then $sMethod = "GET"
	If $sMethod = "GET" Then $sEnctype = ""
EndFunc

Func __WinHttpHTML5FormAttribs(ByRef $aDtas, ByRef $aFlds, ByRef $iNumParams, ByRef $aInput, ByRef $sAction, ByRef $sEnctype, ByRef $sMethod)
	; Clicking "submit" is done using:
	; "type:submit", zero_based_index_of_the_submit_button
	; "name:whatever", True
	; "id:whatever", True
	; "whatever", True     ;<- same as "id:whatever"
	; Clicking "image" is done using:
	; "type:image", "zero_based_index_of_the_image_control Xcoord,Ycoord"
	; "name:whatever", "Xcoord,Ycoord"
	; "id:whatever", "Xcoord,Ycoord"
	; "whatever", "Xcoord,Ycoord"     ;<- same as "id:whatever"
	Local $aSpl, $iSubmitHTML5 = 0, $iInpSubm, $sImgAppx = "."
	For $k = 1 To $iNumParams
		$aSpl = StringSplit($aFlds[$k], ":", 2)
		If $aSpl[0] = "type" And ($aSpl[1] = "submit" Or $aSpl[1] = "image") Then
			Local $iSubmIndex = $aDtas[$k], $iSubmCur = 0, $iImgCur = 0, $sType, $sInpNme
			If $aSpl[1] = "image" Then
				$iSubmIndex = Int($aDtas[$k])
			EndIf
			For $i = 0 To UBound($aInput) - 1 ; for all input elements
				Switch __WinHttpAttribVal($aInput[$i], "type")
					Case "submit"
						If $iSubmCur = $iSubmIndex Then
							$iSubmitHTML5 = 1
							$iInpSubm = $i
							ExitLoop 2
						EndIf
						$iSubmCur += 1
					Case "image"
						If $iImgCur = $iSubmIndex Then
							$iSubmitHTML5 = 1
							$iInpSubm = $i
							$sInpNme = __WinHttpAttribVal($aInput[$iInpSubm], "name")
							If $sInpNme Then $sInpNme &= $sImgAppx
							$aInput[$iInpSubm] = 'type="image" formaction="' & __WinHttpAttribVal($aInput[$iInpSubm], "formaction") & '" formenctype="' & __WinHttpAttribVal($aInput[$iInpSubm], "formenctype") & '" formmethod="' & __WinHttpAttribVal($aInput[$iInpSubm], "formmethod") & '"'
							Local $iX = 0, $iY = 0
							$iX = Int(StringRegExpReplace($aDtas[$k], "(\d+)\h*(\d+),(\d+)", "$2", 1))
							$iY = Int(StringRegExpReplace($aDtas[$k], "(\d+)\h*(\d+),(\d+)", "$3", 1))
							ReDim $aInput[UBound($aInput) + 2]
							$aInput[UBound($aInput) - 2] = 'type="image" name="' & $sInpNme & 'x" value="' & $iX & '"'
							$aInput[UBound($aInput) - 1] = 'type="image" name="' & $sInpNme & 'y" value="' & $iY & '"'
							ExitLoop 2
						EndIf
						$iImgCur += 1
				EndSwitch
			Next
			ElseIf $aSpl[0] = "name" Then
			Local $sInpNme = $aSpl[1], $sType
			For $i = 0 To UBound($aInput) - 1 ; for all input elements
				$sType = __WinHttpAttribVal($aInput[$i], "type")
				If $sType = "submit" Then
					If __WinHttpAttribVal($aInput[$i], "name") = $sInpNme And $aDtas[$k] = True Then
						$iSubmitHTML5 = 1
						$iInpSubm = $i
						ExitLoop 2
					EndIf
				ElseIf $sType = "image" Then
					If __WinHttpAttribVal($aInput[$i], "name") = $sInpNme And $aDtas[$k] Then
						$iSubmitHTML5 = 1
						$iInpSubm = $i
						Local $aStrSplit = StringSplit($aDtas[$k], ",", 3), $iX = 0, $iY = 0
						If Not @error Then
							$iX = Int($aStrSplit[0])
							$iY = Int($aStrSplit[1])
						EndIf
						$aInput[$iInpSubm] = 'type="image" formaction="' & __WinHttpAttribVal($aInput[$iInpSubm], "formaction") & '" formenctype="' & __WinHttpAttribVal($aInput[$iInpSubm], "formenctype") & '" formmethod="' & __WinHttpAttribVal($aInput[$iInpSubm], "formmethod") & '"'
						$sInpNme &= $sImgAppx
						ReDim $aInput[UBound($aInput) + 2]
						$aInput[UBound($aInput) - 2] = 'type="image" name="' & $sInpNme & 'x" value="' & $iX & '"'
						$aInput[UBound($aInput) - 1] = 'type="image" name="' & $sInpNme & 'y" value="' & $iY & '"'
						ExitLoop 2
					EndIf
				EndIf
			Next
		Else ; id
			Local $sInpId, $sType
			If @error Then
				$sInpId = $aSpl[0]
			ElseIf $aSpl[0] = "id" Then
				$sInpId = $aSpl[1]
			EndIf
			For $i = 0 To UBound($aInput) - 1 ; for all input elements
				$sType = __WinHttpAttribVal($aInput[$i], "type")
				If $sType = "submit" Then
					If __WinHttpAttribVal($aInput[$i], "id") = $sInpId And $aDtas[$k] = True Then
						$iSubmitHTML5 = 1
						$iInpSubm = $i
						ExitLoop 2
					EndIf
				ElseIf $sType = "image" Then
					If __WinHttpAttribVal($aInput[$i], "id") = $sInpId And $aDtas[$k] Then
						$iSubmitHTML5 = 1
						$iInpSubm = $i
						Local $sInpNme = __WinHttpAttribVal($aInput[$iInpSubm], "name")
						If $sInpNme Then $sInpNme &= $sImgAppx
						Local $aStrSplit = StringSplit($aDtas[$k], ",", 3), $iX = 0, $iY = 0
						If Not @error Then
							$iX = Int($aStrSplit[0])
							$iY = Int($aStrSplit[1])
						EndIf
						$aInput[$iInpSubm] = 'type="image" formaction="' & __WinHttpAttribVal($aInput[$iInpSubm], "formaction") & '" formenctype="' & __WinHttpAttribVal($aInput[$iInpSubm], "formenctype") & '" formmethod="' & __WinHttpAttribVal($aInput[$iInpSubm], "formmethod") & '"'
						ReDim $aInput[UBound($aInput) + 2]
						$aInput[UBound($aInput) - 2] = 'type="image" name="' & $sInpNme & 'x" value="' & $iX & '"'
						$aInput[UBound($aInput) - 1] = 'type="image" name="' & $sInpNme & 'y" value="' & $iY & '"'
						ExitLoop 2
					EndIf
				EndIf
			Next
		EndIf
	Next
	If $iSubmitHTML5 Then
		Local $iUbound = UBound($aInput) - 1
		If __WinHttpAttribVal($aInput[$iInpSubm], "type") = "image" Then $iUbound -= 2 ; two form fields are added for "image"
		For $j = 0 To $iUbound ; for all other input elements
			If $j = $iInpSubm Then ContinueLoop
			Switch __WinHttpAttribVal($aInput[$j], "type")
				Case "submit", "image"
					$aInput[$j] = "" ; remove any other submit/image controls
			EndSwitch
		Next
		Local $sAttr = __WinHttpAttribVal($aInput[$iInpSubm], "formaction")
		If $sAttr Then $sAction = $sAttr
		$sAttr = __WinHttpAttribVal($aInput[$iInpSubm], "formenctype")
		If $sAttr Then $sEnctype = $sAttr
		$sAttr = __WinHttpAttribVal($aInput[$iInpSubm], "formmethod")
		If $sAttr Then $sMethod = $sAttr
		If __WinHttpAttribVal($aInput[$iInpSubm], "type") = "image" Then $aInput[$iInpSubm] = ""
	EndIf
EndFunc

Func __WinHttpNormalizeForm(ByRef $sForm, $sSpr1, $sSpr2)
	Local $aData = StringToASCIIArray($sForm, Default, Default, 2)
	Local $sOut, $bQuot = False, $bSQuot = False, $bOpTg = True
	For $i = 0 To UBound($aData) - 1
		Switch $aData[$i]
			Case 34
				If $bOpTg Then $bQuot = Not $bQuot
				$sOut &= Chr($aData[$i])
			Case 39
				If $bOpTg Then $bSQuot = Not $bSQuot
				$sOut &= Chr($aData[$i])
			Case 32 ; space
				If $bQuot Or $bSQuot Then
					$sOut &= $sSpr1
				Else
					$sOut &= Chr($aData[$i])
				EndIf
			Case 60 ; <
				If Not $bOpTg Then $bOpTg = True
				$sOut &= Chr($aData[$i])
			Case 62 ; >
				If $bQuot Or $bSQuot Then
					$sOut &= $sSpr2
				Else
					If $bOpTg Then $bOpTg = False
					$sOut &= Chr($aData[$i])
				EndIf
			Case Else
				$sOut &= Chr($aData[$i])
		EndSwitch
	Next
	$sForm = $sOut
EndFunc

Func __WinHttpFinalizeCtrls($sSubmit, $sRadio, $sCheckBox, $sButton, ByRef $sAddData, $sGrSep, $sBound = "")
	If $sSubmit Then ; If no submit is specified
		Local $aSubmit = StringSplit($sSubmit, $sGrSep, 3)
		For $m = 1 To UBound($aSubmit) - 1
			$sAddData = StringRegExpReplace($sAddData, "(?:\Q" & $sBound & "\E|\A)\Q" & $aSubmit[$m] & "\E(?:\Q" & $sBound & "\E|\z)", $sBound)
		Next
		__WinHttpTrimBounds($sAddData, $sBound)
	EndIf
	If $sRadio Then ; If no radio is specified
		If $sRadio <> $sGrSep Then
			For $sElem In StringSplit($sRadio, $sGrSep, 3)
				$sAddData = StringRegExpReplace($sAddData, "(?:\Q" & $sBound & "\E|\A)\Q" & $sElem & "\E(?:\Q" & $sBound & "\E|\z)", $sBound)
			Next
			__WinHttpTrimBounds($sAddData, $sBound)
		EndIf
	EndIf
	If $sCheckBox Then ; If no checkbox is specified
		For $sElem In StringSplit($sCheckBox, $sGrSep, 3)
			$sAddData = StringRegExpReplace($sAddData, "(?:\Q" & $sBound & "\E|\A)\Q" & $sElem & "\E(?:\Q" & $sBound & "\E|\z)", $sBound)
		Next
		__WinHttpTrimBounds($sAddData, $sBound)
	EndIf
	If $sButton Then ; If no button is specified
		For $sElem In StringSplit($sButton, $sGrSep, 3)
			$sAddData = StringRegExpReplace($sAddData, "(?:\Q" & $sBound & "\E|\A)\Q" & $sElem & "\E(?:\Q" & $sBound & "\E|\z)", $sBound)
		Next
		__WinHttpTrimBounds($sAddData, $sBound)
	EndIf
EndFunc

Func __WinHttpTrimBounds(ByRef $sDta, $sBound)
	Local $iBLen = StringLen($sBound)
	If StringRight($sDta, $iBLen) = $sBound Then $sDta = StringTrimRight($sDta, $iBLen)
	If StringLeft($sDta, $iBLen) = $sBound Then $sDta = StringTrimLeft($sDta, $iBLen)
EndFunc

Func __WinHttpFormAttrib(ByRef $aAttrib, $i, $sElement)
	$aAttrib[0][$i] = __WinHttpAttribVal($sElement, "id")
	$aAttrib[1][$i] = __WinHttpAttribVal($sElement, "name")
	$aAttrib[2][$i] = __WinHttpAttribVal($sElement, "value")
	$aAttrib[3][$i] = __WinHttpAttribVal($sElement, "type")
EndFunc

Func __WinHttpAttribVal($sIn, $sAttrib)
	Local $aArray = StringRegExp($sIn, '(?i).*?(\A| )\b' & $sAttrib & '\h*=(\h*"(.*?)"|' & "\h*'(.*?)'|" & '\h*(.*?)(?: |\Z))', 1) ; e.g. id="abc" or id='abc' or id=abc
	If @error Then Return ""
	Return $aArray[UBound($aArray) - 1]
EndFunc

Func __WinHttpFormSend($hInternet, $sMethod, $sAction, $fMultiPart, $sBoundary, $sAddData, $fSecure = False, $sAdditionalHeaders = "", $sCredName = "", $sCredPass = "", $iIgnoreAllCertErrors = 0)
	Local $hRequest
	If $fSecure Then
		$hRequest = _WinHttpOpenRequest($hInternet, $sMethod, $sAction, Default, Default, Default, $WINHTTP_FLAG_SECURE)
		If $iIgnoreAllCertErrors Then _WinHttpSetOption($hRequest, $WINHTTP_OPTION_SECURITY_FLAGS, BitOR($SECURITY_FLAG_IGNORE_UNKNOWN_CA, $SECURITY_FLAG_IGNORE_CERT_DATE_INVALID, $SECURITY_FLAG_IGNORE_CERT_CN_INVALID, $SECURITY_FLAG_IGNORE_CERT_WRONG_USAGE))
	Else
		$hRequest = _WinHttpOpenRequest($hInternet, $sMethod, $sAction)
	EndIf
	If $fMultiPart Then
		_WinHttpAddRequestHeaders($hRequest, "Content-Type: multipart/form-data; boundary=" & $sBoundary)
	Else
		If $sMethod = "POST" Then _WinHttpAddRequestHeaders($hRequest, "Content-Type: application/x-www-form-urlencoded")
	EndIf
	_WinHttpAddRequestHeaders($hRequest, "Accept: application/xml,application/xhtml+xml,text/html;q=0.9,text/plain;q=0.8,*/*;q=0.5")
	_WinHttpAddRequestHeaders($hRequest, "Accept-Charset: utf-8;q=0.7")
	If $sAdditionalHeaders Then _WinHttpAddRequestHeaders($hRequest, $sAdditionalHeaders, BitOR($WINHTTP_ADDREQ_FLAG_REPLACE, $WINHTTP_ADDREQ_FLAG_ADD))
	_WinHttpSetOption($hRequest, $WINHTTP_OPTION_DECOMPRESSION, $WINHTTP_DECOMPRESSION_FLAG_ALL)
	_WinHttpSetOption($hRequest, $WINHTTP_OPTION_UNSAFE_HEADER_PARSING, 1)
	__WinHttpFormUpload($hRequest, "", $sAddData)
	_WinHttpReceiveResponse($hRequest)
	__WinHttpSetCredentials($hRequest, "", $sAddData, $sCredName, $sCredPass, 1)
	Return $hRequest
EndFunc

Func __WinHttpSetCredentials($hRequest, $sHeaders = "", $sOptional = "", $sCredName = "", $sCredPass = "", $iFormFill = 0)
	If $sCredName And $sCredPass Then
		Local $iStatusCode = _WinHttpQueryHeaders($hRequest, $WINHTTP_QUERY_STATUS_CODE)
		; Check status code
		If $iStatusCode = $HTTP_STATUS_DENIED Or $iStatusCode = $HTTP_STATUS_PROXY_AUTH_REQ Then
			; Query Authorization scheme
			Local $iSupportedSchemes, $iFirstScheme, $iAuthTarget
			If _WinHttpQueryAuthSchemes($hRequest, $iSupportedSchemes, $iFirstScheme, $iAuthTarget) Then
				; Set passed credentials
				If $iFirstScheme = $WINHTTP_AUTH_SCHEME_PASSPORT And $iStatusCode = $HTTP_STATUS_PROXY_AUTH_REQ Then
					_WinHttpSetOption($hRequest, $WINHTTP_OPTION_CONFIGURE_PASSPORT_AUTH, $WINHTTP_ENABLE_PASSPORT_AUTH)
					_WinHttpSetOption($hRequest, $WINHTTP_OPTION_PROXY_USERNAME, $sCredName)
					_WinHttpSetOption($hRequest, $WINHTTP_OPTION_PROXY_PASSWORD, $sCredPass)
				Else
					_WinHttpSetCredentials($hRequest, $iAuthTarget, $iFirstScheme, $sCredName, $sCredPass)
				EndIf
				; Send request again now
				If $iFormFill Then
					__WinHttpFormUpload($hRequest, $sHeaders, $sOptional)
				Else
					_WinHttpSendRequest($hRequest, $sHeaders, $sOptional)
				EndIf
				; And wait for the response again
				_WinHttpReceiveResponse($hRequest)
			EndIf
		EndIf
	EndIf
EndFunc

Func __WinHttpFormUpload($hRequest, $sHeaders, $sData)
	Local $aClbk = _WinHttpSimpleFormFill_SetUploadCallback()
	If $aClbk[0] <> Default Then
		Local $iSize = StringLen($sData), $iChunk = Floor($iSize / $aClbk[1]), $iRest = $iSize - ($aClbk[1] - 1) * $iChunk, $iCurCh = $iChunk
		_WinHttpSendRequest($hRequest, Default, Default, $iSize)
		For $i = 1 To $aClbk[1]
			If $i = $aClbk[1] Then $iCurCh = $iRest
			_WinHttpWriteData($hRequest, StringMid($sData, 1 + $iChunk * ($i -1), $iCurCh))
			Call($aClbk[0], Floor($i * 100 / $aClbk[1]))
		Next
	Else
		_WinHttpSendRequest($hRequest, Default, $sData)
	EndIf
EndFunc

Func __WinHttpDefault(ByRef $vInput, $vOutput)
	If $vInput = Default Or Number($vInput) = -1 Then $vInput = $vOutput
EndFunc

Func __WinHttpMemGlobalFree($pMem)
	Local $aCall = DllCall("kernel32.dll", "ptr", "GlobalFree", "ptr", $pMem)
	If @error Or $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc

Func __WinHttpPtrStringLenW($pStr)
	Local $aCall = DllCall("kernel32.dll", "dword", "lstrlenW", "ptr", $pStr)
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc

Func __WinHttpUA()
	Local Static $sUA = "Mozilla/5.0 " & __WinHttpSysInfo() & " WinHttp/" & __WinHttpVer() & " (WinHTTP/5.1) like Gecko"
	Return $sUA
EndFunc

Func __WinHttpSysInfo()
	Local $sDta = FileGetVersion("kernel32.dll")
	$sDta = "(Windows NT " & StringLeft($sDta, StringInStr($sDta, ".", 1, 2) - 1)
	If StringInStr(@OSArch, "64") And Not @AutoItX64 Then $sDta &= "; WOW64"
	$sDta &= ")"
	Return $sDta
EndFunc

Func __WinHttpVer()
	Return "1.6.4.1"
EndFunc

Func _WinHttpBinaryConcat(ByRef $bBinary1, ByRef $bBinary2)
	Local $bOut = _WinHttpSimpleBinaryConcat($bBinary1, $bBinary2)
	Return SetError(@error, 0, $bOut)
EndFunc
;============================================================================================
