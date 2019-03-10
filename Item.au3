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
	$Item.__defineGetter("delete", ItemDelete)
	$Item.__defineGetter("update", ItemUpdate)

	Return $Item
EndFunc

Func ItemDelete($this)

	Local $Item = $this.parent

	Local $itms[0]
	$Item.itms = $itms

	_GUICtrlListView_DeleteAllItems($listPost)
	_GUICtrlListView_DeleteAllItems($listRequest)
	_StLoadHtml($WebView, "")
	GUICtrlSetData($editRaw, $item.raw)

EndFunc

Func ItemAdd($this)

	Local $Item = $this.parent
	Local $Data = $this.arguments.values[0]

	Local $itms = $Item.itms
	Local $UBound = UBound($itms)
	ReDim $itms[$UBound + 1]

	$itms[$UBound] = IDispatch()
	$itms[$UBound].raw = $Data

	$Item.itms = $itms
	$Item.update($UBound)

EndFunc

Func ItemUpdate($this)

	Local $Item = $this.parent
	Local $Index = $this.arguments.values[0]

	Local $itms = $Item.itms
	Local $subItem = $itms[$Index]

	Local $Data = $this.arguments.length < 2 ? $subItem.raw : $this.arguments.values[1]

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


	$subItem.raw = $Data
	$subItem.url = $URL
	$subItem.request = $Request
	$subItem.post = $Post

	$itms[$Index] = $subItem
	$Item.itms = $itms

EndFunc

Func ItemLoad($this)

	Local $Item = $this.parent
	Local $Index = $item.sel
	Local $isReload = $this.arguments.length = 0 ? False : $this.arguments.values[0]

	; if reload data (edit raw)
	If $isReload Then $Item.update($Index, GUICtrlRead($editRaw))

	Local $itms = $Item.itms
	If $Index < 0 Or $Index >= UBound($itms) Then Return False

	Local $subItem = $itms[$Index]
	Local $Request = $subItem.request
	Local $Post = $subItem.post

	GUICtrlSetData($editRaw, $subItem.raw)
	_StLoadHtml($WebView, "")

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

	_StLoadHtml($WebView, $subItem.html)

EndFunc

Func ItemGo($this)

	Local $Item = $this.parent
	Local $Index = $item.sel

	$Item.update($Index, GUICtrlRead($editRaw))
	Local $itms = $Item.itms
	If $Index < 0 Or $Index >= UBound($itms) Then Return False

	Local $subItem = $itms[$Index]
	Local $Request = $subItem.request
	Local $Post = $subItem.post

	; get request and post
	Local $RequestCount = _GUICtrlListView_GetItemCount($listRequest)

	Local $DataRequest = "?" & __GetDataFromList($listRequest, $Request)
	Local $DataPost = __GetDataFromList($listPost, $Post)

	; header, cookie and ref
	Local $cookie = __GetCookieFromHeader($subItem.raw)
	Local $ref = __GetRefFromHeader($subItem.raw)

	Local $selUA = __GetComboSel($comboUA) ; get current selected user agent
	Local $UserAgent = $selUA >= 0 ? $aUserAgent[$selUA].str : ""


	_StLoadHtml($WebView, "Loading, please wait..")

	Local $Html = _HttpRequest(2, $subItem.url & $DataRequest, $DataPost, $cookie, $ref, $UserAgent)

	$subItem.html = __Encode($Html)

	GUICtrlSetData($editText, $Html)
	_StLoadHtml($WebView, $subItem.html)

	$subItem.request = $Request
	$subItem.post = $Post
	$itms[$Index] = $subItem
	$Item.itms = $itms

EndFunc

Func ItemChange($this)

	Local $Item = $this.parent
	Local $iLv = $this.arguments.values[0]
	Local $iRow = $this.arguments.values[1]
	Local $iCol = $this.arguments.values[2]
	Local $list = $iLv = 1 ? $listRequest : $listPost


	Local $Data = _GUICtrlListView_GetItemText($list, $iRow, $iCol)

	Local $w = 500
	Local $h = 35

	If StringLen($Data) > 200 Then
		$w = 800
		$h = 300
	EndIf

	Local $GUI_Change = GUICreate("Change value", $w, $h)
	Local $Input = GUICtrlCreateEdit($Data, 5, 5, $w - 10, $h - 10, 0x0004)

	GUISetState()

	ControlClick($GUI_Change, "", "", "left", 1, 0, 0)
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
	If $Index < 0 Or $Index >= UBound($itms) Then Return False

	Local $Item = $itms[$Index]

	Local $PostCount = _GUICtrlListView_GetItemCount($listPost)

	; header, cookie and ref
	Local $cookie = __GetCookieFromHeader($item.raw)
	Local $ref = __GetRefFromHeader($item.raw)

	Local $selUA = __GetComboSel($comboUA) ; get current selected user agent
	Local $UserAgent = $selUA >= 0 ? $aUserAgent[$selUA].str : ""

	Local $key, $value

	Local $Data = "; code generated by ReTraffic" & @CRLF
	$Data &= 'Local $URL = "' & $Item.url & '?"' & @CRLF
	$Data &= 'Local $Cookie = "' & $cookie & '"' & @CRLF

	If $ref Then $Data &= 'Local $Ref = "' & $ref & '"' & @CRLF
	If $UserAgent Then $Data &= 'Local $UserAgent = "' & $UserAgent & '"' & @CRLF


	$Data &= "Local $Request, $DataRequest, $DataPost" & @CRLF & @CRLF

	$Data &= __GetValue($listRequest, "$DataRequest")

	$Data &= __GetValue($listPost, "$DataPost")

	$Data &= "$Request = _HttpRequest(2, $URL & $DataRequest, $DataPost, $Cookie"

	If $ref Then $Data &= ", $Ref"

	If $UserAgent And $ref Then

		$Data &= ", $UserAgent"
	ElseIf $UserAgent Then

		$Data &= ', "", $UserAgent'
	EndIf

	$Data &= ")"

	ClipPut($Data)

EndFunc

Func __GetDataFromList($list, ByRef $Item)

	Local $listCount = _GUICtrlListView_GetItemCount($list), $Data, $key, $value

	For $i = 0 To $listCount - 1

		$check = _GUICtrlListView_GetItemChecked($list, $i)
		$Item[$i].check = $check

		If $check = False Then ContinueLoop

		$key = _GUICtrlListView_GetItemText($list, $i)
		$value = _GUICtrlListView_GetItemText($list, $i, 1)

		$Item[$i].key = $key
		$Item[$i].value = $value

		$Data &= __URIEncode($key) & "=" & __URIEncode($value)

		If $i < $listCount - 1 Then $Data &= "&"

	Next

	Return $Data

EndFunc

Func __GetKey($list)

	Local $listCount = _GUICtrlListView_GetItemCount($list)

	Local $Data

	For $i = 0 To $listCount - 1

		$key = _GUICtrlListView_GetItemText($list, $i)
		$value = _GUICtrlListView_GetItemText($list, $i, 1)
		$Data &= "Local $" & __normalize($key) & ' = "' & __URIEncode($value) & '"' & @CRLF
	Next

	Return $Data
EndFunc

Func __GetValue($list, $var)

	Local $listCount = _GUICtrlListView_GetItemCount($list)
	Local $Data = __GetKey($list)
	$Data &= $Data ? @CRLF : ""

	For $i = 0 To $listCount - 1

		If _GUICtrlListView_GetItemChecked($list, $i) = False Then ContinueLoop

		$key = _GUICtrlListView_GetItemText($list, $i)
		$value = _GUICtrlListView_GetItemText($list, $i, 1)

		$Data &= $var & ' &= "'
		If $i > 0 Then $Data &= "&"

		$Data &= __URIEncode($key) & '=" & $' & __normalize($key) & @CRLF
	Next

	$Data &= $Data ? @CRLF : ""

	Return $Data

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

Func __normalize($str)

	Return StringReplace($str, " ", "")

EndFunc
