
#include-once

#include <IE.au3>
#include <File.au3>
#include <Misc.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>

#include "AutoitObject_Internal.au3"
#include "_HttpRequest.au3"
#include "ListURL.au3"
#include "Item.au3"
#include "Sciter-UDF.au3"


; GUI
Global $GUI
Global $List, $btnAddNew, $Item
Global $editRaw, $editText, $listRequest, $listPost

; info
Global $isSubGui = False
Global $__version = 0.1

; gui
Global $GuiW = 1200, $GuiH = 800
Global $WebView

Func __GetCookieFromHeader($Header)

	Local $cookie = StringRegExp($Header, "Cookie: (.*?)" & @CRLF, 1)

	If IsArray($cookie) = False Then Return False

	Return $cookie[0]

EndFunc
