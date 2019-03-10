#AutoIt3Wrapper_Run_AU3Check=n
#include "header.au3"

; start Sciter for web view
_StStartup()

; create gui
$GUI = GUICreate("ReTraffic v" & $__version, $GuiW, $GuiH)

$btnAddNew = GUICtrlCreateButton("Add New", 0, 0, 80, 30)
	GUICtrlSetFont(-1, 11)
$btnDelAll = GUICtrlCreateButton("Delete All", 80, 0, 80, 30)
	GUICtrlSetFont(-1, 11)
$btnGo = GUICtrlCreateButton("Go", 270, 0, 80, 30)
	GUICtrlSetFont(-1, 11, 800)
$btnFilter = GUICtrlCreateButton("Run Filter", 180, 0, 90, 30)
	GUICtrlSetFont(-1, 11, 800)

$List = ListURL($GUI, "Host|URL", 0, 30, 350, 763)
$Item = Item()

GUISetState()


; Tab Request
$GUI_Request = GUICreate("", 843, 500, 355, 0, $WS_POPUP, $WS_EX_MDICHILD, $GUI)
$tab_Request = GUICtrlCreateTab(0, 0, 843, 500)

$btnCopy = GUICtrlCreateButton("Copy Code", 750, 0, 94, 22)
GUICtrlSetFont(-1, 11, 800)

$comboUA = GUICtrlCreateCombo("User-Agent default", 140, 0, 115, 22, 3)

$tabRaw = GUICtrlCreateTabItem("Raw")
	$btnEditRaw = GUICtrlCreateButton("Edit", 650, 0, 100, 22)
	GUICtrlSetFont(-1, 11, 800)
	$editRaw = GUICtrlCreateEdit("ReTraffic v" & $__version & " created by Ho Hai Dang", 0, 22, 843, 475)
	GUICtrlSetFont(-1, 11, Default, Default, "CONSOLAS")
	GUICtrlSetState($editRaw, 128)

$tabRequest = GUICtrlCreateTabItem("Request")
	$listRequest = GUICtrlCreateListView("Key|Value", 0, 22, 843, 475, $WS_BORDER, $LVS_EX_FULLROWSELECT + $LVS_EX_CHECKBOXES + $LVS_EX_GRIDLINES)
	_GUICtrlListView_SetColumnWidth($listRequest, 0, 415)
	_GUICtrlListView_SetColumnWidth($listRequest, 1, 415)

$tabPost = GUICtrlCreateTabItem("Post")
	$listPost = GUICtrlCreateListView("Key|Value", 0, 22, 843, 475, $WS_BORDER, $LVS_EX_FULLROWSELECT + $LVS_EX_CHECKBOXES + $LVS_EX_GRIDLINES)
	_GUICtrlListView_SetColumnWidth($listPost, 0, 415)
	_GUICtrlListView_SetColumnWidth($listPost, 1, 415)

GUISetState()

; Tab Respone
$GUI_Respone = GUICreate("", 843, 293, 355, 500, $WS_POPUP, $WS_EX_MDICHILD, $GUI)
$tab_Respone = GUICtrlCreateTab(0, 0, 843, 300)

$tabWeb = GUICtrlCreateTabItem("Web view")
	$WebView = _StincGui($GUI_Respone, 2, 22, 837, 300)

$tabText = GUICtrlCreateTabItem("Text view")
	$editText = GUICtrlCreateEdit("", 0, 22, 837, 271)
	GUICtrlSetState(-1, 2048)
	GUICtrlSetFont(-1, 11, Default, Default, "CONSOLAS")

GUISetState()

; JSON gui
$GUI_JSON = GUICreate("JSON Data", 643, 100)
$listJSON = GUICtrlCreateListView("Key|Value", 0, 0, 643, 875, $WS_BORDER, $LVS_EX_FULLROWSELECT + $LVS_EX_CHECKBOXES + $LVS_EX_GRIDLINES)
_GUICtrlListView_SetColumnWidth($listJSON, 0, 315)
_GUICtrlListView_SetColumnWidth($listJSON, 1, 315)
$btnDone = GUICtrlCreateButton("Done", 550,0,100, 30)
GUICtrlSetFont(-1, 11)
;~ GUISetState()

; ===
$hLV_JSON = GUICtrlGetHandle($listJSON)
$hLV_Request = GUICtrlGetHandle($listRequest)
$hLV_Post = GUICtrlGetHandle($listPost)

GUIRegisterMsg($WM_NOTIFY, "_WM_NOTIFY_Handler")

;=============

_LoadUserAgent()

While 1

	$List.check()

	$aMsg = GUIGetMsg(1)
	Switch $aMsg[0]

		Case $tab_Request
			$sel = GUICtrlRead($tab_Request)

			If $prevSel = 0 And $sel = 1 And GUICtrlGetState($editRaw) <> 144 Then $Item.load(True)
			If $prevSel = 0 And $sel = 2 And GUICtrlGetState($editRaw) <> 144 Then $Item.load(True)

			$prevSel = $sel

		Case $tab_Respone
			$sel = GUICtrlRead($tab_Respone)
			If $sel = 1 Then GUICtrlSetState($tabText, 2048)

		Case $btnAddNew
			AddNewItem()

		Case $btnFilter

			If MsgBox(1, "Cảnh báo", "(BETA)" & @CRLF & "Chức năng này sẽ liên tục gửi request, bạn có đồng ý?") = 1 Then $Item.filter

		Case $btnCopy
			$Item.copy

		Case $btnEditRaw
			GUICtrlSetState($editRaw, GUICtrlGetState($editRaw) = 144 ? 64 : 128)

		Case $btnDelAll
			$Item.delete
			$List.delete

		Case $btnGo
			$Item.go()

		Case - 3
			Switch $aMsg[1]
				Case $GUI
					If MsgBox(1, "Thông báo", "Bạn có muốn thoát?") = 1 Then Exit
			EndSwitch

	EndSwitch

	Sleep(20)

WEnd

Func _LoadUserAgent()

	Local $Str

	For $i = 0 To UBound($aUserAgent) - 1

		$Str &= $aUserAgent[$i].name

		If $i < UBound($aUserAgent) - 1 Then $Str &= "|"

	Next

	GUICtrlSetData($comboUA, $Str)
EndFunc

Func _WM_NOTIFY_Handler($hWnd, $iMsg, $wParam, $lParam)

    #forceref $hWnd, $iMsg, $wParam

    ; Struct = $tagNMHDR and "int Item;int SubItem" from $tagNMLISTVIEW
    Local $tStruct = DllStructCreate("hwnd;uint_ptr;int_ptr;int;int", $lParam)
    If @error Then Return

    Switch DllStructGetData($tStruct, 1)
        Case $hLV_Request
            $iLV = 1
        Case $hLV_Post
            $iLV = 2
        Case Else
            Return
    EndSwitch

    If BitAND(DllStructGetData($tStruct, 3), 0xFFFFFFFF) = $NM_DBLCLK Then

		$iRow = DllStructGetData($tStruct, 4)
		$iCol = DllStructGetData($tStruct, 5)
		If $iRow >= 0 Then $Item.change($iLV, $iRow, $iCol)
	EndIf

EndFunc

