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

$List = ListURL($GUI, "Host|URL", 0, 30, 350, 763)
$Item = Item()

GUISetState()


; Tab Request
$GUI_Request = GUICreate("", 843, 500, 355, 0, $WS_POPUP, $WS_EX_MDICHILD, $GUI)
GUICtrlCreateTab(0, 0, 843, 500)

$btnCopy = GUICtrlCreateButton("Copy Code", 750, 0, 94, 22)
GUICtrlSetFont(-1, 11, 800)

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
GUICtrlCreateTab(0, 0, 843, 300)

$tabWeb = GUICtrlCreateTabItem("Web view")
	$WebView = _StincGui($GUI_Respone, 2, 22, 837, 300)

$tabText = GUICtrlCreateTabItem("Text view")
	$editText = GUICtrlCreateEdit("", 0, 22, 837, 271)
	GUICtrlSetFont(-1, 11, Default, Default, "CONSOLAS")

GUISetState()

$hLV_Request = GUICtrlGetHandle($listRequest)
$hLV_Post = GUICtrlGetHandle($listPost)
GUIRegisterMsg($WM_NOTIFY, "_WM_NOTIFY_Handler")

;=============

While 1

	$List.check()

	$aMsg = GUIGetMsg(1)
	Switch $aMsg[0]

		Case $btnAddNew
			AddNewItem()

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
					If MsgBox(1, "Message", "You gonna exit?") = 1 Then Exit
			EndSwitch

	EndSwitch

	Sleep(20)

WEnd

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