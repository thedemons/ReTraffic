#AutoIt3Wrapper_Run_AU3Check=n
#include-once
#include "header.au3"

Global $__AllList[0]

Func ListURL($GUI, $Text, $x, $y, $w, $h)

	Local $list = _GUICtrlListView_Create($GUI, "ID|" & $Text, $x, $y, $w, $h, $LVS_REPORT + $LVS_SINGLESEL)
	GUICtrlSetState($list, 256)

	_GUICtrlListView_SetExtendedListViewStyle($list, BitOR($LVS_EX_FULLROWSELECT,$LVS_EX_GRIDLINES))
	_GUICtrlListView_JustifyColumn($list, 1, 1)
	_GUICtrlListView_SetColumnWidth($list, 0, 25)
	_GUICtrlListView_SetColumnWidth($list, 1, $w/2.5)
	_GUICtrlListView_SetColumnWidth($list, 2, $w - $w/2.5 - 25)

	Local $UBound = UBound($__AllList)
	ReDim $__AllList[$UBound + 1]

	$__AllList[$UBound] = $list

	Local $hList = IDispatch()
	$hList.index = $UBound
	$hList.sel = -1

	$hList.__defineGetter("add", ListURLAdd)
	$hList.__defineGetter("check", ListURLCheck)
	$hList.__defineGetter("delete", ListURLDelete)

	Return $hList

EndFunc

Func ListURLDelete($this)

	Local $hList = $this.parent
	Local $list = $__AllList[$hList.index]

	_GUICtrlListView_DeleteAllItems($list)

EndFunc

Func ListURLAdd($this)

	Local $hList = $this.parent
	Local $list = $__AllList[$hList.index]
	Local $Add = $this.arguments.values[0]

	If UBound($Add) <> _GUICtrlListView_GetColumnCount($list) - 1 Then Return False

	Local $index = _GUICtrlListView_GetItemCount($list)

	_GUICtrlListView_AddItem($list, $index)

	; add sub item
	For $i = 0 to UBound($Add) - 1

		If StringLeft($Add[$i], 4) = "http" Then $Add[$i] = StringSplit($Add[$i], "//", 1)[2]

		_GUICtrlListView_AddSubItem($list, $index, $Add[$i], $i + 1)
	Next

	;select this item
	_GUICtrlListView_SetItemSelected($list, $index)

	$hList.sel = $index

	$Item.load
	$Item.sel = $index

EndFunc

Func ListURLCheck($this)

	Local $hList = $this.parent
	Local $list = $__AllList[$hList.index]

	Local $sel = _GUICtrlListView_GetSelectedIndices($list, True)

	If $sel[0] = 0 then Return False

	If $sel[1] = $hList.sel Then Return False

	$hlist.sel = $sel[1]

	$Item.sel = $sel[1]
	$Item.load()

	Return $sel[1]

EndFunc

Func AddNewItem()

	Local $aItem, $Text
	Local $Gui_AddNew = GUICreate("Add New item", 800, 600)
	Local $Edit = GUICtrlCreateEdit("", 0, 0, 800, 600)

	GUISetState()

	While 1

		$Text = ControlGetText("Progress Telerik Fiddler Web Debugger", "", "[NAME:txtRaw; INSTANCE:1]")
		$aItem = __GetItemFromData($Text)

		If $aItem = False Then

			$Text = ControlGetText("Progress Telerik Fiddler Web Debugger", "", "[NAME:txtRaw; INSTANCE:2]")
			$aItem = __GetItemFromData($Text)
		EndIf

		If $aItem = False Then $aItem = __GetItemFromData(ClipGet())

		If $aItem = False Then $aItem = __GetItemFromData( GUICtrlRead($Edit) )

		If $aItem <> False Then

			GUIDelete($Gui_AddNew)
			$List.add($aItem)

			ControlClick($GUI, "", "[CLASS:SysListView32; INSTANCE:1]", "left", 1, 0, 0)
			Return True

		EndIf

		If _IsPressed("0D") And $aItem = False Then MsgBox(0, "Lỗi", "Data không hợp lệ, vui lòng kiểm tra lại")


		Switch GUIGetMsg()
			Case - 3
				GUIDelete($Gui_AddNew)
				Return False
		EndSwitch

	WEnd

EndFunc

Func __GetItemFromData($Data)

	Local $URL = StringRegExp($Data, " (.*?) ", 1)
	If IsArray($URL) = False Then Return False

	If StringInStr($URL[0], "?") Then $URL[0] = StringSplit($URL[0], "?", 1)[1]

	Local $aHost = StringSplit($URL[0], "/", 1)
	If $aHost[0] < 3 Then Return False

	Local $Host = $aHost[1] & "//" & $aHost[3]
	Local $URL = "/"

	For $i = 4 to $aHost[0]

		$URL &= $aHost[$i]
		If $i < $aHost[0] Then $URL &= "/"
	Next

	Local $Ret[2] = [$Host, $URL]

	$Item.add($Data)

	Return $Ret
EndFunc
