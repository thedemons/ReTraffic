#include-once
#include "header.au3"

Func __JSON2Array($itemJson, $str)

	Local $len = StringLen($str), $pLastKey = 2, $Split, $sRemoveKey
	Local $Data[0], $Index = 0

	For $i = 1 To $len

		$char = StringMid($str, $i, 1)
		If $char = ":" Then

			$Index = UBound($Data)

			ReDim $Data[$Index + 1]

			$Data[$Index] = IDispatch()
			$Data[$Index].check = True

			If $Index <= UBound($itemJson) - 1 Then $Data[$Index].check = $itemJson[$Index].check

			$Data[$Index].key = StringSplit(StringMid($str, $pLastKey, $i), ":", 1)[1]

			$sRemoveKey = StringTrimLeft($str, $i)

			$value = _JSON_GetBlock($sRemoveKey, False)

			; not a json block
			If $value = False Then

				$Split = StringSplit($sRemoveKey, ",", 1)

				If $Split[0] = 1 Then $Split[1] = StringTrimRight($Split[1], 1)

				$Data[$Index].value = $Split[1]

				$i += StringLen($Split[1])
			Else

				$Data[$Index].value = $value

				$i += StringLen($value)
			EndIf

			$pLastKey = $i + 2
		EndIf
	Next

	Return $Data

EndFunc

Func _JSON_GetBlock($str, $isFind = True)

	Local $len = StringLen($str)
	Local $pStart = 1, $Open = 0, $pEnd = 0

	If $isFind Then
		For $i = 1 to $len

			If StringMid($str, $i, 1) = "{" Then
				$pStart = $i
				ExitLoop
			EndIf
		Next
	EndIf

	If StringMid($str, $pStart, 1) <> "{" Then Return False

	For $i = $pStart + 1 To $len

		$char = StringMid($str, $i, 1)

		If $char = "{" Then $Open += 1
		If $char = "}" Then $Open -= 1

		If $Open < 0 Then
			$pEnd = $i
			ExitLoop
		EndIf
	Next

	If $pEnd > 0 Then Return StringMid($str, $pStart, $pEnd)

	Return False

EndFunc

Func __isJSON($str)

	If _JSON_GetBlock($str, False) Then Return True
	Return False

EndFunc