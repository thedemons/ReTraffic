;ver=0.3
#include-once

#include <IE.au3>
#include <File.au3>
#include <Misc.au3>
#include <GuiComboBoxEx.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>

#include "AutoitObject_Internal.au3"
#include "_HttpRequest.au3"
#include "ListURL.au3"
#include "Item.au3"
#include "JSON.au3"
#include "Sciter-UDF.au3"


; GUI
Global $GUI
Global $List, $btnAddNew, $Item, $prevSel = 0, $btnDone
Global $editRaw, $editText, $listRequest, $listPost, $listJSON, $GUI_JSON

; info
Global $isSubGui = False
Global $__version = 0.3

; gui
Global $GuiW = 1200, $GuiH = 800
Global $WebView

Global $aUserAgent[0]

__CreateUserAgent("Chrome", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.71 Safari/537.36")
__CreateUserAgent("Firefox", "Mozilla/5.0 (Windows NT 6.2; WOW64; rv:63.0) Gecko/20100101 Firefox/63.0")
__CreateUserAgent("Safari iOS 12", "Mozilla/5.0 (iPhone; CPU iPhone OS 12_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/12.0 Mobile/15E148 Safari/604.1")


Func __GetCookieFromHeader($Header)

	Local $cookie = StringRegExp($Header, "Cookie: (.*?)" & @CRLF, 1)

	If IsArray($cookie) = False Then Return ""

	Return $cookie[0]

EndFunc

Func __GetRefFromHeader($Header)

	Local $ref = StringRegExp($Header, "Referer: (.*?)" & @CRLF, 1)

	If IsArray($ref) = False Then Return ""

	Return $ref[0]

EndFunc

Func __GetUAFromHeader($Header)

	Local $UA = StringRegExp($Header, "User-Agent: (.*?)" & @CRLF, 1)

	If IsArray($UA) = False Then Return ""

	Return $UA[0]

EndFunc

Func __CreateUserAgent($name, $str)

	$index = UBound($aUserAgent)
	ReDim $aUserAgent[ $index + 1]

	$aUserAgent[$index] = IDispatch()
	$aUserAgent[$index].name = $name
	$aUserAgent[$index].str = $str

EndFunc

; copied
Func __GetComboSel($hCombo)

    Local $sSelectedText = GUICtrlRead($hCombo)
    Local $iEntryId = -1 ;; zero based.

    For $i = 0 To UBound($aUserAgent) - 1

        If Not ($aUserAgent[$i].name == $sSelectedText) Then ContinueLoop ;; case sensitve.

        $iEntryId = $i
        ExitLoop
    Next

    If $iEntryId < 0 Then SetError(1)
    Return $iEntryId
EndFunc

; utf 8
Func __Encode($str)
	$len = StringLen($str)

	For $i = $len To 1 Step - 1

		$chr = StringMid($str, $i, 1)
		$asc = Asc($chr)

		; if unicode
		If Chr($asc) = $chr Then ContinueLoop

		$str = StringReplace($str, $i, @LF)
		$str = StringReplace($str, @LF, "&#x" & Hex(AscW($chr), 4) & ";")
	Next
	Return $str
EndFunc
