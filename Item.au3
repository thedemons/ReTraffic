#AutoIt3Wrapper_Run_AU3Check=n
#include-once
#include "header.au3"

Func Item()

	Local $Item = IDispatch()

	Local $itms[0]
	$Item.itms = $itms
	$Item.sel = False

	$Item.__defineGetter("add", ItemAdd)
	$Item.__defineGetter("load", ItemLoad)
	$Item.__defineGetter("go", ItemGo)
	$Item.__defineGetter("change", ItemChange)
	$Item.__defineGetter("copy", ItemCopy)

	Return $Item
EndFunc

Func ItemAdd($this)

	Local $Item = $this.parent
	Local $Data = $this.arguments.values[0]

	Local $itms = $Item.itms
	Local $UBound = UBound($itms)
	ReDim $itms[$UBound + 1]

	; split url
	Local $URL = StringRegExp($Data, " (.*?) ", 1)
	If IsArray($URL) = False Then Return False

	; split url & request
	Local $Request = StringSplit($URL[0], "?", 1)
	$URL = $Request[1]

	; if any request /?a=b&c=d
	$Request = __GetRequestData($Request)

	; if any post data
	Local $Post = __GetPostData($Data)


	$itms[$UBound] = IDispatch()
	$itms[$UBound].raw = $Data
	$itms[$UBound].url = $URL

	$itms[$UBound].request = $Request

	$itms[$UBound].post = $Post

	$Item.itms = $itms
EndFunc

Func ItemLoad($this)

	Local $Item = $this.parent
	Local $Index = $this.arguments.values[0]

	Local $itms = $Item.itms
	Local $item = $itms[$Index]
	Local $Request = $item.request
	Local $Post = $item.post

	GUICtrlSetData($editRaw, $item.raw)

	_GUICtrlListView_DeleteAllItems($listPost)
	_GUICtrlListView_DeleteAllItems($listRequest)

	If $Request <> False Then

		For $i = 0 To UBound($Request) - 1

			$indexItem = _GUICtrlListView_AddItem($listRequest, $Request[$i].key)
			_GUICtrlListView_AddSubItem($listRequest, $indexItem, $Request[$i].value, 1)
			_GUICtrlListView_SetItemChecked($listRequest, $indexItem, $Request[$i].check)
		Next
	EndIf

	If $Post <> False Then

		For $i = 0 To UBound($Post) - 1

			$indexItem = _GUICtrlListView_AddItem($listPost, $Post[$i].key)
			_GUICtrlListView_AddSubItem($listPost, $indexItem, $Post[$i].value, 1)
			_GUICtrlListView_SetItemChecked($listPost, $indexItem, $Post[$i].check)
		Next
	EndIf


EndFunc

Func ItemGo($this)

	Local $Item = $this.parent
	Local $Index = $item.sel

	Local $itms = $Item.itms
	Local $Item = $itms[$Index]
	Local $Request = $item.request
	Local $Post = $item.post

	Local $RequestCount = _GUICtrlListView_GetItemCount($listRequest)
	Local $PostCount = _GUICtrlListView_GetItemCount($listPost)

	Local $DataRequest = "?", $DataPost, $key, $value

	For $i = 0 To $RequestCount - 1

		If _GUICtrlListView_GetItemChecked($listRequest, $i) = False Then ContinueLoop

		$key = _GUICtrlListView_GetItemText($listRequest, $i)
		$value = _GUICtrlListView_GetItemText($listRequest, $i, 1)

		$DataRequest &= __URIEncode($key) & "=" & __URIEncode($value)

		If $i < $RequestCount - 1 Then $DataRequest &= "&"

	Next

	For $i = 0 To $PostCount - 1

		If _GUICtrlListView_GetItemChecked($listPost, $i) = False Then ContinueLoop

		$key = _GUICtrlListView_GetItemText($listPost, $i)
		$value = _GUICtrlListView_GetItemText($listPost, $i, 1)

		$DataPost &= __URIEncode($key) & "=" & __URIEncode($value)

		If $i < $PostCount - 1 Then $DataPost &= "&"

	Next

	_StLoadHtml($WebView, "Loading, please wait..")
	Local $cookie = __GetCookieFromHeader($item.raw)
	Local $Html = _HttpRequest(2, $item.url & $DataRequest, $DataPost, $cookie)

	GUICtrlSetData($editText, $Html)
	_StLoadHtml($WebView, $Html)

EndFunc

Func ItemChange($this)

	Local $Item = $this.parent
	Local $iLv = $this.arguments.values[0]
	Local $iRow = $this.arguments.values[1]
	Local $iCol = $this.arguments.values[2]
	Local $list = $iLv = 1 ? $listRequest : $listPost


	Local $Data = _GUICtrlListView_GetItemText($list, $iRow, $iCol)

	Local $GUI_Change = GUICreate("Change value", 300, 35)
	Local $Input = GUICtrlCreateInput($Data, 5, 5, 290, 25)

	GUISetState()

	While 1

		If _IsPressed("0D") Then

			_GUICtrlListView_SetItemText($list, $iRow, GUICtrlRead($Input), $iCol)
			GUIDelete($GUI_Change)

			Return True
		EndIf

		Switch GUIGetMsg()
			Case -3
				GUIDelete($GUI_Change)
				Return False
		EndSwitch

		Sleep(10)
	WEnd

EndFunc

Func ItemCopy($this)

	Local $Item = $this.parent
	Local $Index = $item.sel

	Local $itms = $Item.itms
	Local $Item = $itms[$Index]

	Local $RequestCount = _GUICtrlListView_GetItemCount($listRequest)
	Local $PostCount = _GUICtrlListView_GetItemCount($listPost)

	Local $cookie = __GetCookieFromHeader($item.raw)

	Local $DataRequest = "?", $DataPost, $key, $value

	Local $Data = 'Local $URL = "' & $Item.url & '"' & @CRLF

	$Data &= 'Local $Cookie = "' & $cookie & '"' & @CRLF

	$Data &= "Local $Request, $DataRequest, $DataPost" & @CRLF

	For $i = 0 To $RequestCount - 1

		If _GUICtrlListView_GetItemChecked($listRequest, $i) = False Then ContinueLoop

		$key = _GUICtrlListView_GetItemText($listRequest, $i)
		$value = _GUICtrlListView_GetItemText($listRequest, $i, 1)

		$Data &= '$DataRequest &= "' & __URIEncode($key) & "=" & __URIEncode($value)

		If $i < $RequestCount - 1 Then $Data &= "&"

		$Data &= '"' & @CRLF

	Next

	$Data &= @CRLF

	For $i = 0 To $PostCount - 1

		If _GUICtrlListView_GetItemChecked($listPost, $i) = False Then ContinueLoop

		$key = _GUICtrlListView_GetItemText($listPost, $i)
		$value = _GUICtrlListView_GetItemText($listPost, $i, 1)

		$Data &= '$DataPost &= "' & __URIEncode($key) & "=" & __URIEncode($value)

		If $i < $PostCount - 1 Then $Data &= "&"

		$Data &= '"' & @CRLF

	Next

	$Data &= @CRLF

	$Data &= "$Request = _HttpRequest(2, $URL & $DataRequest, $DataPost, $Cookie)"

	ClipPut($Data)

EndFunc

Func __GetPostData($Data)

	Local $Post = StringSplit($Data, @CRLF & @CRLF, 1)

	If $Post[0] < 2 Then $Post = False
	$Post = $Post <> False ? $Post[ $Post[0] ] : False
	If $Post Then $Post = __Key2Array($Post)

	Return $Post
EndFunc

Func __GetRequestData($Request)

	If $Request[0] < 2 Then $Request = False
	$Request = $Request <> False ? $Request[ $Request[0] ] : False

	If $Request Then $Request = __Key2Array($Request)

	Return $Request
EndFunc

Func __Key2Array($Data)

	Local $aData = StringSplit($Data, "&", 1), $SplitData

	Local $Ret[ $aData[0] ], $key, $value

	For $i = 1 To $aData[0]

		; if data not contain "="
		If StringInStr( $aData[$i], "=") = False Then

			$key = $aData[$i]
			$value = ""
		Else

			$SplitData = StringSplit($aData[$i], "=", 1)
			$key = $SplitData[1]
			$value = $SplitData[2]
		EndIf

		$Ret[$i - 1] = IDispatch()
		$Ret[$i - 1].key = __URIDecode( $key ) ; key
		$Ret[$i - 1].value = __URIDecode( $value ) ; value
		$Ret[$i - 1].check = True

	Next

	If UBound($Ret) = 0 Then Return False

	Return $Ret
EndFunc