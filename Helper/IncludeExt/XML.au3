#include-once
;===============================================================================
; XML Helper Functions using Microsoft.XMLDOM
;===============================================================================
#include "CommonFunctions.au3"

; Load XML from file or string and return the DOM object
Func _XMLLoad($sXMLSource, $bIsFile = True)
	Local $oXML = ObjCreate("Microsoft.XMLDOM")
	If Not IsObj($oXML) Then
		_Log("_XMLLoad: Failed to create Microsoft.XMLDOM object")
		Return SetError(1, 0, 0)
	EndIf

	$oXML.async = False
	$oXML.preserveWhiteSpace = True
	$oXML.setProperty("SelectionLanguage", "XPath")

	Local $bResult
	If $bIsFile Then
		$bResult = $oXML.load($sXMLSource)
	Else
		$bResult = $oXML.loadXML($sXMLSource)
	EndIf

	If Not $bResult Then
		Local $oError = $oXML.parseError
		_Log("_XMLLoad: Parse error - " & $oError.reason & " Line: " & $oError.line)
		Return SetError(2, $oError.line, 0)
	EndIf

	Return $oXML
EndFunc   ;==>_XMLLoad

; Get the text value of an XML node by XPath, searches all namespaces
Func _XMLGetValue($oXML, $sXPath)
	If Not IsObj($oXML) Then Return SetError(1, 0, "")

	; Try direct XPath first
	Local $oNode = $oXML.selectSingleNode($sXPath)

	; If not found, try with local-name() for namespace-agnostic search
	If Not IsObj($oNode) Then
		; Convert simple path like "//ComputerName" to namespace-agnostic version
		Local $sLocalNameXPath = _XPathToLocalName($sXPath)
		$oNode = $oXML.selectSingleNode($sLocalNameXPath)
	EndIf

	If IsObj($oNode) Then
		Return $oNode.text
	EndIf

	Return SetError(2, 0, "")
EndFunc   ;==>_XMLGetValue

; Set the text value of XML nodes by XPath (updates all matching nodes)
Func _XMLSetValue($oXML, $sXPath, $sValue)
	If Not IsObj($oXML) Then Return SetError(1, 0, 0)

	; Try direct XPath first
	Local $oNodes = $oXML.selectNodes($sXPath)

	; If not found, try with local-name() for namespace-agnostic search
	If $oNodes.length = 0 Then
		Local $sLocalNameXPath = _XPathToLocalName($sXPath)
		$oNodes = $oXML.selectNodes($sLocalNameXPath)
	EndIf

	If $oNodes.length = 0 Then Return SetError(2, 0, 0)

	For $i = 0 To $oNodes.length - 1
		$oNodes.item($i).text = $sValue
		_Log("_XMLSetValue: Set " & $sXPath & " = " & $sValue)
	Next

	Return $oNodes.length
EndFunc   ;==>_XMLSetValue

; Convert simple XPath to namespace-agnostic XPath using local-name()
Func _XPathToLocalName($sXPath)
	; Handle paths like "//InstallFrom/Path" -> "//*[local-name()='InstallFrom']/*[local-name()='Path']"
	Local $sResult = $sXPath

	; Replace //ElementName with //*[local-name()='ElementName']
	$sResult = StringRegExpReplace($sResult, "//([A-Za-z][A-Za-z0-9_]*)", "//*[local-name()='$1']")

	; Replace /ElementName with /*[local-name()='ElementName']
	$sResult = StringRegExpReplace($sResult, "/([A-Za-z][A-Za-z0-9_]*)(?![^\[]*\])", "/*[local-name()='$1']")

	Return $sResult
EndFunc   ;==>_XPathToLocalName

; Remove XML nodes matching XPath
Func _XMLRemoveNodes($oXML, $sXPath)
	If Not IsObj($oXML) Then Return SetError(1, 0, 0)

	Local $oNodes = $oXML.selectNodes($sXPath)
	If $oNodes.length = 0 Then
		Local $sLocalNameXPath = _XPathToLocalName($sXPath)
		$oNodes = $oXML.selectNodes($sLocalNameXPath)
	EndIf

	Local $iCount = 0
	For $i = $oNodes.length - 1 To 0 Step -1
		Local $oNode = $oNodes.item($i)
		If IsObj($oNode) And IsObj($oNode.parentNode) Then
			$oNode.parentNode.removeChild($oNode)
			$iCount += 1
			_Log("_XMLRemoveNodes: Removed node matching " & $sXPath)
		EndIf
	Next

	Return $iCount
EndFunc   ;==>_XMLRemoveNodes

; Uncomment XML sections by removing comment markers around specific tags
; Replaces <!--TagName with empty and TagName--> with empty
Func _XMLUncommentSection($oXML, $sTag)
	If Not IsObj($oXML) Then Return SetError(1, 0, 0)

	Local $sXMLString = $oXML.xml
	$sXMLString = StringReplace($sXMLString, "<!--" & $sTag, "")
	$sXMLString = StringReplace($sXMLString, $sTag & "-->", "")

	; Reload the modified XML
	$oXML.loadXML($sXMLString)

	If $oXML.parseError.errorCode <> 0 Then
		_Log("_XMLUncommentSection: Error after uncommenting " & $sTag & " - " & $oXML.parseError.reason)
		Return SetError(2, 0, 0)
	EndIf

	_Log("_XMLUncommentSection: Uncommented section " & $sTag)
	Return 1
EndFunc   ;==>_XMLUncommentSection

; Save XML object to string
Func _XMLToString($oXML)
	If Not IsObj($oXML) Then Return SetError(1, 0, "")
	Return $oXML.xml
EndFunc   ;==>_XMLToString

; Save XML object to file
Func _XMLSave($oXML, $sFilePath)
	If Not IsObj($oXML) Then Return SetError(1, 0, 0)

	Local $hFile = FileOpen($sFilePath, $FO_OVERWRITE + $FO_UTF8_NOBOM)
	If $hFile = -1 Then
		_Log("_XMLSave: Failed to open file for writing: " & $sFilePath)
		Return SetError(2, 0, 0)
	EndIf

	FileWrite($hFile, $oXML.xml)
	FileClose($hFile)

	_Log("_XMLSave: Saved to " & $sFilePath)
	Return 1
EndFunc   ;==>_XMLSave