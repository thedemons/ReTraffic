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
	$Item.__defineGetter("filter", ItemFilter)

	Return $Item
EndFunc

Func ItemDelete($this)

	Local $Item = $this.parent

	Local $itms[0]
	$Item.itms = $itms

	_GUICtrlListView_DeleteAllItems($listPost)
	_GUICtrlListView_DeleteAllItems($listRequest)
	_StLoadHtml($WebView, "")
	GUICtrlSetData($editRaw, $Item.raw)

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

	If $Index < 0 Or $Index >= UBound($itms) Then Return False

	Local $subItem = $itms[$Index]

	Local $Data = $this.arguments.length < 2 ? $subItem.raw : $this.arguments.values[1]

	; split url
	Local $URL = StringRegExp($Data, " (.*?) ", 1)
	If IsArray($URL) = False Then Return False

	; split url & request
	Local $Request = StringSplit($URL[0], "?", 1)
	$URL = $Request[1]

	; if any request /?a=b&c=d
	$Request = __GetData($subItem.request, $Request)

	; if any post data
	Local $Post = StringSplit($Data, @CRLF & @CRLF, 1)
	$Post = __GetData($subItem.post, $Post)


	$subItem.raw = $Data
	$subItem.url = $URL
	$subItem.request = $Request
	$subItem.post = $Post

	$itms[$Index] = $subItem
	$Item.itms = $itms

EndFunc

Func ItemLoad($this)

	Local $Item = $this.parent
	Local $Index = $Item.sel
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
	Local $Index = $Item.sel
	Local $isWebView = $this.arguments.length = 0 ? True : $this.arguments.values[0]

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

	If $isWebView Then
		$subItem.html = __Encode($Html)
		_StLoadHtml($WebView, $subItem.html)
	Else
		$subItem.html = $Html
	EndIf

	GUICtrlSetData($editText, $Html)


	$subItem.request = $Request
	$subItem.post = $Post
	$itms[$Index] = $subItem
	$Item.itms = $itms

	Return $subItem.html
EndFunc

Func ItemFilter($this)

	Local $Item = $this.parent
	Local $RequestCount = _GUICtrlListView_GetItemCount($listRequest)
	Local $PostCount = _GUICtrlListView_GetItemCount($listPost)

	__ListCheckAllItem($listRequest)
	__ListCheckAllItem($listPost)

	; get source target
	Local $Target = $Item.go(False)

	__TestItem($Item, $listRequest, $Target)
	__TestItem($Item, $listPost, $Target)

	$Item.go

EndFunc

Func __TestItem($Item, $list, $Target)

	Local $listCount = _GUICtrlListView_GetItemCount($list)

	For $i = 0 To $listCount - 1 Step + 3

		For $n = $i To $i + 2

			If $n > $listCount - 1 Then ExitLoop
			_GUICtrlListView_SetItemChecked($list, $n, False)
		Next

		$Html = $Item.go(False)

		If __CheckResult($Html, $Target) = False Then

			For $n = $i To $i + 2

				If $n > $listCount - 1 Then ExitLoop
				_GUICtrlListView_SetItemChecked($list, $n, True)
			Next

			For $n = $i To $i + 2

				If $n > $listCount - 1 Then ExitLoop

				_GUICtrlListView_SetItemChecked($list, $n, False)

				Local $value = _GUICtrlListView_GetItemText($list, $n, 1)

				$Html = $Item.go(False)
				If __CheckResult($Html, $Target) = False Then

					_GUICtrlListView_SetItemChecked($list, $n, True)
					If __isJSON($value) Then __TestItemJSON($Item, $n, $Target, $list = $listRequest ? 1 : 2)
				EndIf
			Next
		EndIf
	Next

EndFunc

Func __TestItemJSON($Item, $iRow, $Target, $list)

	Local $Index = $Item.sel
	Local $itms = $Item.itms

	If $Index < 0 Or $Index >= UBound($itms) Then Return False

	Local $subItem = $itms[$Index]
	Local $itemList = $list = 1 ? $subItem.request[$iRow] : $subItem.post[$iRow]
	Local $Json = $itemList.json

	Global $w = 643, $h = (UBound($Json) + 2) * 20 + 22, $l = (@DesktopWidth - $w) / 2, $t = (@DesktopHeight - $h + 40) / 2

	WinMove($GUI_JSON, "", $l, $t, $w, $h)
	GUICtrlSetPos($listJSON, 0, 0, $w, $h)
	GUICtrlSetPos($btnDone, 270, $h + 5, 100, 30)
	GUISetState(@SW_SHOW, $GUI_JSON)

	For $i = 0 To UBound($Json) - 1
		$Json[$i].check = True
		_GUICtrlListView_AddItem($listJSON, $Json[$i].key)
		_GUICtrlListView_AddSubItem($listJSON, $i, $Json[$i].value, 1)
		_GUICtrlListView_SetItemChecked($listJSON, $i, True)
	Next

	__ItemSetJsonData($Item, $iRow, $Json, $list)

	For $i = 0 To UBound($Json) - 1 Step + 3

		For $n = $i To $i + 2

			If $n > UBound($Json) - 1 Then ExitLoop
			$Json[$n].check = False
			_GUICtrlListView_SetItemChecked($listJSON, $n, False)
		Next

		__ItemSetJsonData($Item, $iRow, $Json, $list)

		$Html = $Item.go(False)

		If __CheckResult($Html, $Target) = False Then

			For $n = $i To $i + 2

				If $n > UBound($Json) - 1 Then ExitLoop
				$Json[$n].check = True
				_GUICtrlListView_SetItemChecked($listJSON, $n, True)
			Next

			__ItemSetJsonData($Item, $iRow, $Json, $list)

			For $n = $i To $i + 2

				If $n > UBound($Json) - 1 Then ExitLoop

				$Json[$n].check = False
				_GUICtrlListView_SetItemChecked($listJSON, $n, False)

				__ItemSetJsonData($Item, $iRow, $Json, $list)

				$Html = $Item.go(False)
				If __CheckResult($Html, $Target) = False Then

					$Json[$n].check = True
					_GUICtrlListView_SetItemChecked($listJSON, $n, True)
					__ItemSetJsonData($Item, $iRow, $Json, $list)
				EndIf
			Next
		EndIf
	Next
	GUISetState(@SW_HIDE, $GUI_JSON)

EndFunc

Func __ItemSetJsonData($Item, $iRow, $Json, $list)

	Local $Index = $Item.sel
	Local $itms = $Item.itms

	If $Index < 0 Or $Index >= UBound($itms) Then Return False

	Local $subItem = $itms[$Index]
	Local $itemList = $list = 1 ? $subItem.request[$iRow] : $subItem.post[$iRow]

	$itemList.json = $Json

	If $list = 1 Then

		Local $temp = $subItem.request
		$temp[$iRow] = $itemList
		$subItem.request = $temp
	Else

		Local $temp = $subItem.post
		$temp[$iRow] = $itemList
		$subItem.post = $temp
	EndIf

	$itms[$Index] = $subItem
	$Item.itms = $itms

EndFunc

Func __ListCheckAllItem($list, $check = True)

	Local $listCount = _GUICtrlListView_GetItemCount($list)

	For $i = 0 To $listCount - 1
		_GUICtrlListView_SetItemChecked($list, $i, $check)

	Next
EndFunc

Func ItemChange($this)

	Local $Item = $this.parent
	Local $iLv = $this.arguments.values[0]
	Local $iRow = $this.arguments.values[1]
	Local $iCol = $this.arguments.values[2]
	Local $list = $iLv = 1 ? $listRequest : $listPost

	Local $Data = _GUICtrlListView_GetItemText($list, $iRow, $iCol)

	If __isJSON($Data) Then Return ItemChangeJSON($Item, $list, $iRow)

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

Func ItemChangeJSON($Item, $iLv, $iRow)

	Local $Index = $Item.sel
	Local $itms = $Item.itms

	If $Index < 0 Or $Index >= UBound($itms) Then Return False

	Local $subItem = $itms[$Index]
	Local $itemList = $iLv = 1 ? $subItem.request[$iRow] : $subItem.post[$iRow]

	Local $Json = $itemList.json
	Global $w = 643, $h = (UBound($Json) + 2) * 20 + 22, $l = (@DesktopWidth - $w) / 2, $t = (@DesktopHeight - $h + 40) / 2

	WinMove($GUI_JSON, "", $l, $t, $w, $h + 70)
	GUICtrlSetPos($listJSON, 0, 0, $w, $h)
	GUICtrlSetPos($btnDone, 270, $h + 5, 100, 30)
	GUISetState(@SW_SHOW, $GUI_JSON)

	_GUICtrlListView_DeleteAllItems($listJSON)
	For $i = 0 to UBound($Json) -1
		_GUICtrlListView_AddItem($listJSON, $Json[$i].key)
		_GUICtrlListView_AddSubItem($listJSON, $i, $Json[$i].value, 1)
		_GUICtrlListView_SetItemChecked($listJSON, $i, $Json[$i].check)
	Next

	While 1

		Switch GUIGetMsg()

			Case -3
				GUISetState(@SW_HIDE, $GUI_JSON)
				Return False

			Case $btnDone

				For $i = 0 to UBound($Json) - 1

					$Json[$i].key = _GUICtrlListView_GetItemText($listJSON, $i)
					$Json[$i].value = _GUICtrlListView_GetItemText($listJSON, $i, 1)
					$Json[$i].check = _GUICtrlListView_GetItemChecked($listJSON, $i)
				Next

				$itemList.json = $Json
				If $iLv = 1 Then $subItem.request[$iRow] = $itemList
				If $iLv = 2 Then $subItem.post[$iRow] = $itemList
				$itms[$Index] = $subItem
				$Item.itms = $itms

				GUISetState(@SW_HIDE, $GUI_JSON)
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

	Local $subItem = $itms[$Index]

	Local $PostCount = _GUICtrlListView_GetItemCount($listPost)

	; header, cookie and ref
	Local $cookie = __GetCookieFromHeader($subItem.raw)
	Local $ref = __GetRefFromHeader($subItem.raw)

	Local $selUA = __GetComboSel($comboUA) ; get current selected user agent
	Local $UserAgent = $selUA >= 0 ? $aUserAgent[$selUA].str : ""

	Local $key, $value

	Local $Data = "; code generated by ReTraffic" & @CRLF
	$Data &= 'Local $URL = "' & $subItem.url & '?"' & @CRLF
	$Data &= 'Local $Cookie = "' & $cookie & '"' & @CRLF

	If $ref Then $Data &= 'Local $Ref = "' & $ref & '"' & @CRLF
	If $UserAgent Then $Data &= 'Local $UserAgent = "' & $UserAgent & '"' & @CRLF


	$Data &= "Local $Request, $DataRequest, $DataPost" & @CRLF & @CRLF

	$Data &= __GetValue($Item, 1, $listRequest, "$DataRequest")

	$Data &= __GetValue($Item, 2, $listPost, "$DataPost")

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

		If __isJSON($value) Then

			$Json = $Item[$i].json
			$value = "{"

			For $n = 0 to UBound($Json) - 1

				If $Json[$n].check = False Then ContinueLoop

				$value &= $Json[$n].key & ":" & $Json[$n].value
				If $i < UBound($Json) - 1 Then $value &= ","
			Next

			$value &= "}"
		EndIf

		$Item[$i].key = $key
		$Item[$i].value = $value

		$Data &= __URIEncode($key) & "=" & __URIEncode($value)

		If $i < $listCount - 1 Then $Data &= "&"

	Next

	Return $Data

EndFunc

Func __GetKey($Item, $iLV, $list)

	Local $Index = $Item.sel
	Local $itms = $Item.itms

	If $Index < 0 Or $Index >= UBound($itms) Then Return False

	Local $subItem = $itms[$Index]

	Local $listCount = _GUICtrlListView_GetItemCount($list)

	Local $Data

	For $i = 0 To $listCount - 1

		If _GUICtrlListView_GetItemChecked($list, $i) = False Then ContinueLoop

		$key = _GUICtrlListView_GetItemText($list, $i)
		$value = _GUICtrlListView_GetItemText($list, $i, 1)
		MsgBox(0,"",$key)

		If __isJSON($value) Then
			MsgBox(0,"",$value)

			Local $itemList = $iLv = 1 ? $subItem.request[$i] : $subItem.post[$i]
			Local $Json = $itemList.json

			$Data &= "Local $" & __normalize($key) & ' = "{"' & @CRLF

			For $n = 0 to UBound($Json) - 1

				If $Json[$n].check = False Then ContinueLoop

				$Data &=  "$" & __normalize($key) & " &= '" & $Json[$n].key & ":" & $Json[$n].value

				If $n < UBound($Json) - 1 Then
					$Data &= ",'" & @CRLF
				Else
					$Data &= "}'" & @CRLF
				EndIf
			Next

			$Data &= "$" & __normalize($key) & " = _URIEncode($" & __normalize($key) & ")" & @CRLF & @CRLF

		Else

			$Data &= "Local $" & __normalize($key) & ' = "' & __URIEncode($value) & '"' & @CRLF

			MsgBox(0,"",$Data)
		EndIf
	Next

	Return $Data
EndFunc

Func __GetValue($Item, $iLv, $list, $var)

	Local $listCount = _GUICtrlListView_GetItemCount($list)
	Local $Data = __GetKey($Item, $iLv, $list)
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

Func __GetData($listItem, $Data)

	If $Data[0] < 2 Then $Data = False
	$Data = $Data <> False ? $Data[ $Data[0] ] : False

	If $Data Then $Data = __Key2Array($listItem, $Data)

	Return $Data
EndFunc

Func __Key2Array($listItem, $Data)

	Local $aData = StringSplit($Data, "&", 1), $SplitData

	Local $Ret[ $aData[0] ], $key, $value

	For $i = 1 To $aData[0]

		; if data not contain "="
		If StringInStr( $aData[$i], "=") = False Then

			$key = __URIDecode( $aData[$i] )
			$value = ""
		Else

			$SplitData = StringSplit($aData[$i], "=", 1)
			$key = __URIDecode( $SplitData[1] )
			$value = __URIDecode( $SplitData[2] )
		EndIf

		$Ret[$i - 1] = IDispatch()
		$Ret[$i - 1].key = $key ; key
		$Ret[$i - 1].value = $value ; value
		$Ret[$i - 1].check = True


		If __isJSON($value) Then

			If $i <= UBound($listItem) Then
				$Ret[$i - 1].json = __JSON2Array($listItem[$i - 1].json, $value)
			Else
				$Ret[$i - 1].json = __JSON2Array("", $value)
			EndIf
		EndIf

		If $i <= UBound($listItem) Then $Ret[$i - 1].check = $listItem[$i - 1].check

	Next

	If UBound($Ret) = 0 Then Return False

	Return $Ret
EndFunc

Func __CheckResult($Html, $Target)

	Local $len = StringLen($Target)

	If StringLen($Html) < $len - $len  * 0.2 Or StringLen($Html) > $len + $len  * 0.2 Then Return False

	If StringInStr($Html, "error") And StringInStr($Target, "error") = False Then Return False

	If StringInStr($Html, "head") = False And StringInStr($Target, "head") Then Return False
	If StringInStr($Html, "head") And StringInStr($Target, "head") = False Then Return False
	If StringInStr($Html, "warning") And StringInStr($Target, "warning") = False Then Return False

	Return True
EndFunc

Func __normalize($str)

	Return StringReplace($str, " ", "")

EndFunc

Func _WM_NOTIFY_JSON($hWnd, $iMsg, $wParam, $lParam)

    #forceref $hWnd, $iMsg, $wParam

    ; Struct = $tagNMHDR and "int Item;int SubItem" from $tagNMLISTVIEW
    Local $tStruct = DllStructCreate("hwnd;uint_ptr;int_ptr;int;int", $lParam)
    If @error Then Return

    Switch DllStructGetData($tStruct, 1)
        Case $hLV_JSON
            $iLV = 1
        Case Else
            Return
    EndSwitch

    If BitAND(DllStructGetData($tStruct, 3), 0xFFFFFFFF) = $NM_DBLCLK Then

		MsgBox(0,"","")
	EndIf

EndFunc
