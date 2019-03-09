;Thanks [@ProAndy, Trancexx, Firefox - WinHttp UDF] autoitscript.com

#pragma compile(AutoItExecuteAllowed, True)
#Au3Stripper_Ignore_Funcs=_JS_Execute, _HttpRequest_BypassCloudflare, _HTMLEncode, __HTML_RegexpReplace, __IE_Init_GoogleBox, _Data2SendEncode
__HttpRequest_CheckUpdate(1403)

#include-once
#include <Array.au3>
#include <Crypt.au3>
#include <GDIPlus.au3>
#include <WinAPI.au3>


Global Const $g___ConsoleForceUTF8 = False
Global Const $g___ConsoleForceANSI = False
;-----------------------------------------------------------------------------------
Global $dll_WinHttp = DllOpen('winhttp.dll')
Global $dll_User32 = DllOpen('user32.dll')
Global $dll_Gdi32, $dll_WinInet
;-----------------------------------------------------------------------------------
Global $g___ChromeVersion = FileGetVersion('C:\Program Files (x86)\Google\Chrome\Application\chrome.exe')
If @error Then $g___ChromeVersion = FileGetVersion('C:\Program Files\Google\Chrome\Application\chrome.exe')
If @error Or $g___ChromeVersion = '' Or $g___ChromeVersion = '0.0.0.0' Then $g___ChromeVersion = '70.0.3538.102'
;------------------------------------------------------------------------------------
Global Const $g___UAHeader = 'Mozilla/5.0 (Windows NT ' & StringRegExpReplace(FileGetVersion('kernel32.dll'), '^(\d+\.\d+)(.*)$', '$1', 1) & ((StringInStr(@OSArch, '64') And Not @AutoItX64) ? '; WOW64' : '') & ') '
Global Const $g___defUserAgent = $g___UAHeader & '_HttpRequest' & ' (WinHTTP/5.1) like Gecko'
Global Const $g___defUserAgentW = $g___UAHeader & 'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/' & $g___ChromeVersion & ' Safari/537.36'
Global Const $g___defUserAgentA = 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'
Global Const $g___defUserAgentAO = 'Mozilla/5.0 (Linux; U; Android 4.2.2; en-us; SM-T217S Build/JDQ39) AppleWebKit/534.30 (KHTML, like Gecko) Version/4.0 Safari/534.30' ;Samsung Galaxy Tab3 7.0
;-----------------------------------------------------------------------------------------
Global $g___MaxSession_TT = 110, $g___MaxSession_USE = 100, $g___LastSession = 0
Global $g___sBaseURL[$g___MaxSession_TT], $g___UserAgent[$g___MaxSession_TT] = [$g___defUserAgentW]
Global $g___retData[$g___MaxSession_TT][2]
Global $g___ftpOpen[$g___MaxSession_TT], $g___ftpConnect[$g___MaxSession_TT]
Global $g___hOpen[$g___MaxSession_TT], $g___hConnect[$g___MaxSession_TT], $g___hRequest[$g___MaxSession_TT], $g___hWebSocket[$g___MaxSession_TT]
Global $g___hProxy[$g___MaxSession_TT][5] ;Proxy|ProxyBk|ProxyBypass|ProxyUserName|ProxyPassword
Global $g___hCredential[$g___MaxSession_TT][2] ;Username|Password
;------------------------------------------------------------------------------------
Global $g___hCookie[$g___MaxSession_TT], $g___hCookieLast = '', $g___hCookieDomain = '', $g___hCookieRemember = False
;------------------------------------------------------------------------------------
Global $g___CookieJarPath = ''
Global $g___CookieJarINI = ObjCreate("Scripting.Dictionary")
$g___CookieJarINI.CompareMode = 1
;------------------------------------------------------------------------------------
Global Const $def___sChr64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
Global Const $def___aChr64 = StringSplit($def___sChr64, "", 2)
Global Const $def___sPadding = '='
Global $g___sChr64 = $def___sChr64
Global $g___aChr64 = $def___aChr64
Global $g___sPadding = $def___sPadding
;------------------------------------------------------------------------------------
Global $g___hStatusCallback = DllCallbackRegister("__HttpRequest_StatusCallback", "none", "handle;dword_ptr;dword;ptr;dword")
Global $g___pStatusCallback = DllCallbackGetPtr($g___hStatusCallback)
;-----------------------------------------------------------------------------------------
Global $g___oError = ObjEvent("AutoIt.Error", "__ObjectErrDetect"), $g___oErrorStop = 0
;-----------------------------------------------------------------------------------------
Global $g___JsLibGunzip = '', $g___JsLibJSON = ''
;-----------------------------------------------------------------------------------------
Global $g___aReadWriteData = [['char', 'byte'], [StringMid, BinaryMid], [StringLen, BinaryLen]]
Global $g___HttpRequestReg = 'HKCU\Software\AutoIt v3\HttpRequest\'
Global $g___oDicEntity, $g___oDicHiddenSearch
Global $g___OnlineCompilerTimer = TimerInit()
Global $g___CancelReadWrite = True
Global $g___BytesPerLoop = 8192
Global $g___ErrorNotify = True
Global $g___LocationRedirect = ''
Global $g___CheckConnect = ''
Global $g___aVietPattern = ''
Global $g___OldConsole = ''
Global $g___sData2Send = ''
Global $g___MIMEData = ''
Global $g___HotkeySet = ''
Global $g___Boundary = ''
Global $g___ServerIP = ''
Global $g___CertInfo = ''
Global $g___TimeOut = ''
Global $g___aChrEnt = ''
;-----------------------------------------------------------------------------------------
OnAutoItExitRegister('__HttpRequest_CloseAll')
__HttpRequest_CancelReadWrite()
__SciTE_ConsoleWrite_FixFont()
;__SciTE_ConsoleClear()
ConsoleWrite(@CRLF)



Func _HttpRequest($iReturn, $sURL = '', $sData2Send = '', $sCookie = '', $sReferer = '', $sAdditional_Headers = '', $sMethod = '', $CallBackFunc_Progress = '')
	If StringRegExp($iReturn, '(?i)^\h*?curl\h+') Then
		Local $vData = _HttpRequest_ParseCURL($iReturn)
		Return SetError(@error, @extended, $vData)
	EndIf
	;-------------------------------------------------
	Local $aRetMode = __HttpRequest_iReturnSplit($iReturn)
	If @error Then Return SetError(2, -1, '')
	$g___LastSession = $aRetMode[8]
	;-------------------------------------------------
	If StringRegExp($sURL, '^\h*?/\w') Then $sURL = $g___sBaseURL[$g___LastSession] & $sURL
	;-------------------------------------------------
	Local $aURL = __HttpRequest_URLSplit($sURL)
	If @error Then Return SetError(1, -1, '')
	;-------------------------------------------------
	Local $vContentType = '', $vAcceptType = '', $vUserAgent = '', $vBoundary = '', $vReset = 0, $vUpload = 0, $vWebsocket = 0
	Local $sServerUserName = '', $sServerPassword = '', $sProxyUserName = '', $sProxyPassword = ''
	$g___LocationRedirect = ''
	$g___sData2Send = ''
	$g___retData[$g___LastSession][0] = ''
	$g___retData[$g___LastSession][1] = Binary('')
	;-------------------------------------------------
	If $aURL[0] = 3 Then Return _FtpRequest($aRetMode, $aURL, $sData2Send, $CallBackFunc_Progress)
	;-------------------------------------------------
	If $g___hRequest[$g___LastSession] Then $g___hRequest[$g___LastSession] = _WinHttpCloseHandle2($g___hRequest[$g___LastSession])
	If $g___hWebSocket[$g___LastSession] Then $g___hWebSocket[$g___LastSession] = _WinHttpWebSocketClose2($g___hWebSocket[$g___LastSession])
	;-------------------------------------------------
	If Not $g___hOpen[$g___LastSession] Or $g___hProxy[$g___LastSession][1] <> $g___hProxy[$g___LastSession][0] Then
		If $g___hConnect[$g___LastSession] Then $g___hConnect[$g___LastSession] = _WinHttpCloseHandle2($g___hConnect[$g___LastSession])
		If $g___hOpen[$g___LastSession] Then $g___hOpen[$g___LastSession] = _WinHttpCloseHandle2($g___hOpen[$g___LastSession])
		If $g___hProxy[$g___LastSession][1] <> $g___hProxy[$g___LastSession][0] Then $g___hProxy[$g___LastSession][1] = $g___hProxy[$g___LastSession][0]
		$g___hOpen[$g___LastSession] = _WinHttpOpen2($g___hProxy[$g___LastSession][0], $g___hProxy[$g___LastSession][2])
		_WinHttpSetStatusCallback2($g___hOpen[$g___LastSession], $g___pStatusCallback, 0x00014002)
		_WinHttpSetOption2($g___hOpen[$g___LastSession], 84, 0xA80) ;OPTION_SECURE_PROTOCOLS = 0xA80 (TSL), 0xA8 (SSL + TSL1.0), 0xA00 (TSL1_1 + TSL1_2)
		;_WinHttpSetOption2($g___hOpen[$g___LastSession], 118, 3) ;OPTION_DECOMPRESSION = DECOMPRESSION_FLAG_ALL
		_WinHttpSetOption2($g___hOpen[$g___LastSession], 88, 2) ;OPTION_REDIRECT_POLICY = REDIRECT_POLICY_ALWAYS
		;_WinHttpSetOption2($g___hOpen[$g___LastSession], 4, 20) ;OPTION_CONNECT_RETRIES
		;_WinHttpSetOption2($g___hOpen[$g___LastSession], 89, 20) ;OPTION_MAX_HTTP_AUTOMATIC_REDIRECTS
		;_WinHttpSetOption2($g___hOpen[$g___LastSession], 91, 128 * 1024) ;OPTION_MAX_RESPONSE_HEADER_SIZE. Default: 64Kb
		;_WinHttpSetOption2($g___hOpen[$g___LastSession], 92, 2 * 1024^2) ;OPTION_MAX_RESPONSE_DRAIN_SIZE. Default = 1Mb
		;_WinHttpSetOption2($g___hOpen[$g___LastSession], 79, 2) ;OPTION_ENABLE_FEATURE = ENABLE_SSL_REVERT_IMPERSONATION
		;_WinHttpSetOption2($g___hOpen[$g___LastSession], 133, 1) ;OPTION_ENABLE_HTTP_PROTOCOL = FLAG_HTTP2 (Supported on Windows10 version 1607 and newer)
		$vReset = 1
	EndIf
	;----------------------------------------------------
	If $vReset = 1 Or $g___CheckConnect <> $g___LastSession & $aURL[2] & $aURL[1] Then
		$g___CheckConnect = $g___LastSession & $aURL[2] & $aURL[1]
		If $g___hConnect[$g___LastSession] Then $g___hConnect[$g___LastSession] = _WinHttpCloseHandle2($g___hConnect[$g___LastSession])
		$g___hConnect[$g___LastSession] = _WinHttpConnect2($g___hOpen[$g___LastSession], $aURL[2], $aURL[1])
	EndIf
	;-------------------------------------------------
	If IsArray($sData2Send) Then $sData2Send = _HttpRequest_CreateDataForm($sData2Send)
	_ArrayDisplay($sData2Send)
	;-------------------------------------------------
	If $aURL[8] Or $aRetMode[13] Then $vWebsocket = 1
	;-------------------------------------------------
	$sMethod = ($vWebsocket ? 'GET' : ($sMethod ? $sMethod : ($sData2Send ? 'POST' : 'GET')))
	;-------------------------------------------------
	$g___hRequest[$g___LastSession] = _WinHttpOpenRequest2($g___hConnect[$g___LastSession], $sMethod, $aURL[3], ($aURL[0] - 1) * 0x800000)
	_WinHttpSetOption2($g___hRequest[$g___LastSession], 31, 0x3300) ;OPTION_SECURITY_FLAGS = SECURITY_FLAG_IGNORE_ALL
	_WinHttpSetOption2($g___hRequest[$g___LastSession], 110, 1) ;OPTION_UNSAFE_HEADER_PARSING
	_WinHttpSetOption2($g___hRequest[$g___LastSession], 47, 0) ;OPTION_CLIENT_CERT_CONTEXT= NO_CERT
	;_WinHttpSetOption2($g___hRequest[$g___LastSession], 79, 1) ;OPTION_ENABLE_FEATURE = ENABLE_SSL_REVOCATION
	;-----------------------------------------------------------
	If $g___TimeOut Then _WinHttpSetTimeouts2($g___hRequest[$g___LastSession], $g___TimeOut, $g___TimeOut, $g___TimeOut)
	;------------------------------------------------------------
	If $vWebsocket And _WinHttpSetOptionEx2($g___hRequest[$g___LastSession], 114, 0, True) = 0 Then ;OPTION_UPGRADE_TO_WEB_SOCKET
		Return SetError(113, __HttpRequest_ErrNotify('_HttpRequest', 'WebSocket đã upgrade thất bại', -1), '')
	EndIf
	;------------------------------------------------------------
	If $aRetMode[3] Then _WinHttpSetOption2($g___hRequest[$g___LastSession], 63, 2) ;WINHTTP_DISABLE_REDIRECTS
	;-------------------------------------------------------------------------------------------------------------------------------
	If $aRetMode[5] Then ;Proxy cục bộ
		$sProxyUserName = $aRetMode[6]
		$sProxyPassword = $aRetMode[7]
		_WinHttpSetProxy2($g___hRequest[$g___LastSession], $aRetMode[5])
	ElseIf $g___hProxy[$g___LastSession][0] Then ;Proxy toàn cục
		$sProxyUserName = $g___hProxy[$g___LastSession][3]
		$sProxyPassword = $g___hProxy[$g___LastSession][4]
	EndIf
	If $sProxyUserName Then _WinHttpSetCredentials2($g___hRequest[$g___LastSession], $sProxyUserName, $sProxyPassword, 1, 1)
	;------------------------------------------------------------------------------------------------------------------------------
	If $aURL[4] Then ;Set cục bộ - $aURL[4], $aURL[5] nghĩa là URL có kèm user/pass
		$sServerUserName = $aURL[4]
		$sServerPassword = $aURL[5]
	ElseIf $g___hCredential[$g___LastSession][0] Then ;Set toàn cục
		$sServerUserName = $g___hCredential[$g___LastSession][0]
		$sServerPassword = $g___hCredential[$g___LastSession][1]
	EndIf
	If $sServerUserName Then _WinHttpSetCredentials2($g___hRequest[$g___LastSession], $sServerUserName, $sServerPassword, 0, 1)
	;----------------------------------------------------------------------------------------------------------------------------------------
	#cs
		- A typical WinHTTP application completes the following steps In order To handle authentication.
		• Request a resource With WinHttpOpenRequest And WinHttpSendRequest.
		• Check the response headers With WinHttpQueryHeaders.
		• If a 401 Or 407 status code is returned indicating that authentication is required, call WinHttpQueryAuthSchemes To find an acceptable scheme.
		• Set the authentication scheme, username, And password With WinHttpSetCredentials.
		• Resend the request With the same request handle by calling WinHttpSendRequest.

		- The credentials set by WinHttpSetCredentials are only used For one request.
		• WinHTTP does Not cache the credentials To use In other requests, which means that applications must be written that can respond To multiple requests.
		• If an authenticated connection is re - used, other requests may Not be challenged, but your code should be able To respond To a request at any time.
	#ce
	;----------------------------------------------------------------------------------------------------------------------------------------
	If $sAdditional_Headers Then
		Local $aAddition = StringRegExp($sAdditional_Headers, '(?i)\h*?([\w\-]+)\h*:\h*(.*?)(?:\||$)', 3)
		$sAdditional_Headers = ''
		For $i = 0 To UBound($aAddition) - 1 Step 2
			Switch $aAddition[$i]
				Case 'Accept'
					$vAcceptType = $aAddition[$i + 1]
				Case 'Content-Type'
					$vContentType = $aAddition[$i] & ': ' & $aAddition[$i + 1]
				Case 'Referer'
					If Not $sReferer Then $sReferer = $aAddition[$i + 1]
				Case 'Cookie'
					If Not $sCookie Then $sCookie = $aAddition[$i + 1]
				Case 'User-Agent'
					$vUserAgent = $aAddition[$i + 1]
				Case Else
					$sAdditional_Headers &= $aAddition[$i] & ': ' & $aAddition[$i + 1] & @CRLF
			EndSwitch
		Next
	EndIf
	;-------------------------------------------------
	$sAdditional_Headers &= 'User-Agent: ' & ($vUserAgent ? $vUserAgent : $g___UserAgent[$g___LastSession]) & @CRLF
	$sAdditional_Headers &= 'Accept: ' & ($vAcceptType ? $vAcceptType : '*/*') & @CRLF
	$sAdditional_Headers &= 'DNT: 1' & @CRLF
	;-------------------------------------------------
	If $sReferer Then $sAdditional_Headers &= 'Referer: ' & StringRegExpReplace($sReferer, '(?i)^\h*?Referer\h*?:\h*', '', 1) & @CRLF
	;-------------------------------------------------
	If $sCookie Then
		If $sMethod = 'POST' And StringInStr($aURL[3], 'login', 0, 1) Then __HttpRequest_ErrNotify('_HttpRequest', 'Nạp Cookie vào request liên quan đến Login có thể khiến request thất bại', '', 'Warning')
		If $sCookie == -1 Or $sCookie = 'CookieJar' Then
			If Not $g___CookieJarPath Then Return SetError(9, __HttpRequest_ErrNotify('_HttpRequest', 'CookieJar chưa được active. Vui lòng khởi tạo _HttpRequest_CookieJarSet', -1), '')
			$sCookie = _HttpRequest_CookieJarSearch($sURL)
		Else
			$sCookie = StringRegExpReplace($sCookie, '(?i)^\h*?Cookie\h*?:\h*', '', 1)
		EndIf
		If $g___hCookieRemember And($g___hCookieLast <> $sCookie Or $g___hCookieDomain <> $aURL[9]) Then
			__CookieGlobal_Insert($aURL[9], $sCookie)
			$g___hCookieDomain = $aURL[9]
			$g___hCookieLast = $sCookie
		EndIf
	EndIf
	If $g___hCookieRemember And $g___hCookie[$g___LastSession] Then $sCookie = __CookieGlobal_Search($sURL)
	If $sCookie Then $sAdditional_Headers &= 'Cookie: ' & $sCookie & @CRLF
	;----------------------------------------------------------------------------------------------------------------------------------------
	If $sData2Send Then
		If Not $g___Boundary And StringInStr($vContentType, 'multipart', 0, 1) Then
			$vBoundary = StringRegExp($vContentType, '(?i);\h*?boundary\h*?=\h*?([\w\-]+)', 1)
			If Not @error Then
				$g___Boundary = '--' & $vBoundary[0]
				If Not StringRegExp($sData2Send, '(?is)^' & $g___Boundary) Then
					Return SetError(22, __HttpRequest_ErrNotify('_HttpRequest', '$sData2Send có Boundary không khớp với khai báo ở header Content-Type', -1), '')
				ElseIf Not StringRegExp($sData2Send, '(?is)' & $g___Boundary & '--\R*?$') Then
					Return SetError(23, __HttpRequest_ErrNotify('_HttpRequest', 'Chuỗi Boundary ở cuối $sData2Send phải có -- ở cuối', -1), '')
				EndIf
			EndIf
		EndIf
		;----------------------------------------------
		If $g___Boundary Then
			If Not $vContentType Then $vContentType = 'Content-Type: multipart/form-data; boundary=' & StringTrimLeft($g___Boundary, 2)
			$g___Boundary = ''
			$vUpload = 1
		Else
			If Not $vContentType Then
				If StringRegExp($sData2Send, '^\h*?[\{\[]') Then
					$vContentType = 'Content-Type: application/json'
				Else
					$vContentType = 'Content-Type: application/x-www-form-urlencoded'
					__Data2Send_CheckEncode($sData2Send)
				EndIf
			EndIf
			;If Not IsBinary($sData2Send) Then $sData2Send = StringToBinary($sData2Send, $aRetMode[11])
		EndIf
	EndIf
	;----------------------------------------------------------------------------------------------------------------------------------------
	If Not _WinHttpSendRequest2($g___hRequest[$g___LastSession], $sAdditional_Headers & $vContentType, $vWebsocket ? '' : $sData2Send, $vUpload, $CallBackFunc_Progress) Then
		If @error = 999 Then Return SetError(999, -1, '')
		Return SetError(4, __HttpRequest_ErrNotify('_HttpRequest', 'Gửi request thất bại', -1), '')
	EndIf
	;----------------------------------------------------------------------------------------------------------------------------------------
	If $aRetMode[14] Then Return True
	;----------------------------------------------------------------------------------------------------------------------------------------
	If Not _WinHttpReceiveResponse2($g___hRequest[$g___LastSession]) Then
		Local $ErrorCode = DllCall("kernel32.dll", "dword", "GetLastError")[0]
		If $ErrorCode = 0 Then $ErrorCode = 12003
		Local $ErrorString = _WinHttpGetResponseErrorCode2($ErrorCode)
		Return SetError(5, __HttpRequest_ErrNotify('_HttpRequest', 'Không nhận được response từ Server. Mã lỗi: ' & $ErrorCode & ' (' & $ErrorString & ')', -1), '')
	EndIf
	;----------------------------------------------------------------------------------------------------------------------------------------
	$g___sData2Send = $sData2Send
	Local $vResponse_StatusCode = _WinHttpQueryHeaders2($g___hRequest[$g___LastSession], 19)
	Switch $vResponse_StatusCode
		Case 0 ;Nếu không nhận được Status Code
			Return SetError(6, -1, '')
			;--------------------------
		Case 404 ; Nếu báo lỗi URL không tồn tại (HTTP_STATUS_NOT_FOUND)
			Local $aURLwithHashTag = StringRegExp($sURL, '(?m)(.*)(\#[\w\.\-]+)$', 3)
			If Not @error Then ;Nếu tồn tại chỉ định mục con trong URL
				__HttpRequest_ErrNotify('_HttpRequest', 'Vui lòng bỏ chỉ định mục con (HashTag) ở đuôi URL ( ' & $aURLwithHashTag[1] & ' ) để tránh lãng phí thời gian Redirect', '', 'Warning')
				Local $sHeader = _WinHttpQueryHeaders2($g___hRequest[$g___LastSession], 22)
				Local $vReturn = _HttpRequest($iReturn, $aURLwithHashTag[0], $sData2Send, $sCookie, $sReferer, $sAdditional_Headers, $sMethod, $CallBackFunc_Progress)
				Local $aExtraInfo = [@error, @extended]
				$g___retData[$g___LastSession][0] = $sHeader & @CRLF & 'Redirect → [' & $aURLwithHashTag[0] & ']' & @CRLF & $g___retData[$g___LastSession][0]
				If $iReturn = 1 Then
					$vReturn = $g___retData[$g___LastSession][0]
				ElseIf $iReturn = 4 Or $iReturn = 5 Then
					$vReturn[0] = $g___retData[$g___LastSession][0]
				EndIf
				Return SetError($aExtraInfo[0], $aExtraInfo[1], $vReturn)
			EndIf
			;--------------------------
		Case 401, 407 ;Nếu request yêu cầu Auth (HTTP_STATUS_DENIED hoặc HTTP_STATUS_PROXY_AUTH_REQ)
			If ($vResponse_StatusCode = 401 And $sServerUserName = '') Then
				__HttpRequest_ErrNotify('_HttpRequest', $aURL[2] & ' yêu cầu phải có quyền truy cập')
			ElseIf ($vResponse_StatusCode = 407 And $sProxyUserName = '') Then
				__HttpRequest_ErrNotify('_HttpRequest', 'Proxy này yêu cầu quyền phải có truy cập')
			Else
				For $i = 1 To 3
					_HttpRequest_ConsoleWrite('> Đang tiến hành Authentication ... (' & $i & ')' & @CRLF)
					Local $aSchemes = _WinHttpQueryAuthSchemes2($g___hRequest[$g___LastSession]) ;Return AuthScheme, AuthTarget, SupportedSchemes
					If @error Then ContinueLoop (1 + 0 * __HttpRequest_ErrNotify('_WinHttpQueryAuthSchemes2', 'Không lấy được Authorization Schemes'))
					If $aSchemes[1] = 0 Then ;AUTH_TARGET_SERVER
						_WinHttpSetCredentials2($g___hRequest[$g___LastSession], $sServerUserName, $sServerPassword, 0, $aSchemes[0]) ;https://airbrake.io/blog/http-errors/401-unauthorized-error
					Else ;AUTH_TARGET_PROXY
						_WinHttpSetCredentials2($g___hRequest[$g___LastSession], $sProxyUserName, $sProxyPassword, 1, $aSchemes[0]) ;https://airbrake.io/blog/http-errors/407-proxy-authentication-required
					EndIf
					If @error Then ContinueLoop (1 + 0 * __HttpRequest_ErrNotify('_WinHttpSetCredentials2', 'Cài đặt Credentials thất bại'))
					_WinHttpSendRequest2($g___hRequest[$g___LastSession])
					_WinHttpReceiveResponse2($g___hRequest[$g___LastSession])
					$vResponse_StatusCode = _WinHttpQueryHeaders2($g___hRequest[$g___LastSession], 19)
					If $vResponse_StatusCode <> 401 And $vResponse_StatusCode <> 407 Then ExitLoop
				Next
				If $i = 4 Then __HttpRequest_ErrNotify('_HttpRequest', 'Quá trình Authentication thất bại')
			EndIf
		Case 445 ;REQUEST_CONFLICT
			Local $iTimerInit = TimerInit()
			Do
				If TimerDiff($iTimerInit) > 20000 Then ExitLoop
				Sleep(Random(100, 300, 1))
				_WinHttpSendRequest2($g___hRequest[$g___LastSession])
				_WinHttpReceiveResponse2($g___hRequest[$g___LastSession])
				$vResponse_StatusCode = _WinHttpQueryHeaders2($g___hRequest[$g___LastSession], 19)
			Until $vResponse_StatusCode <> 445
	EndSwitch
	;--------------------------------------------------------
	$g___retData[$g___LastSession][0] &= __CookieJar_Insert($aURL[2], _WinHttpQueryHeaders2($g___hRequest[$g___LastSession], 22))
	;--------------------------------------------------------
	If $vWebsocket Then
		_WinHttpWebSocketRequest($sData2Send)
		If @error Then Return SetError(@error, $vResponse_StatusCode, $aRetMode[0] = 1 ? $g___retData[$g___LastSession][0] : False)
		Return SetError(0, $vResponse_StatusCode, $aRetMode[0] = 1 ? $g___retData[$g___LastSession][0] : True)
	EndIf
	;--------------------------------------------------------
	Switch $aRetMode[0]
		Case 0, 1
			If $aRetMode[2] Then
				$sCookie = _GetCookie($g___retData[$g___LastSession][0])
				Return SetError(@error ? 7 : 0, $vResponse_StatusCode, $sCookie)
			Else
				Return SetError(0, $vResponse_StatusCode, $g___retData[$g___LastSession][0])
			EndIf
			;------------------------------------------
		Case 2 To 5
			If $aRetMode[9] Then ;Ghi file: iReturn có dạng FilePath:Encoding. Khi $aRetMode[9] được set thì kiểu Data trả về sẽ tự động set về 3 (Binary) bất chấp đã điền kiểu Data trả về là gì
				_WinHttpReadData_Ex($g___hRequest[$g___LastSession], $CallBackFunc_Progress, $aRetMode[9], $aRetMode[10])
				Return SetError(@error, $vResponse_StatusCode, $g___retData[$g___LastSession][0])
			EndIf
			$g___retData[$g___LastSession][1] = _WinHttpReadData_Ex($g___hRequest[$g___LastSession], $CallBackFunc_Progress)
			If @error Then Return SetError(@error, $vResponse_StatusCode, '')
			;------------------------------------------
			If StringRegExp(BinaryMid($g___retData[$g___LastSession][1], 1, 1), '(?i)0x(1F|08|8B)') Then $g___retData[$g___LastSession][1] = __Gzip_Uncompress($g___retData[$g___LastSession][1])
			;------------------------------------------
			If $aRetMode[2] = 1 Or $aRetMode[0] = 3 Or $aRetMode[0] = 5 Then ;$aRetMode[2] = 1: force Binary
				If $aRetMode[0] < 4 Then
					Return SetError(0, $vResponse_StatusCode, $g___retData[$g___LastSession][1])
				Else
					Local $aRet = [$g___retData[$g___LastSession][0], $g___retData[$g___LastSession][1]]
					Return SetError(0, $vResponse_StatusCode, $aRet)
				EndIf
			Else
				Local $sRet = $g___retData[$g___LastSession][1]
				$sRet = BinaryToString($sRet, $aRetMode[11]) ; $aRetMode[11] = 1: force ANSI, = 0 (Default): UTF8
				If $aRetMode[12] Then ;force return Raw Text
					$sRet = _HTML_Execute($sRet)
				ElseIf $aRetMode[4] Then ;trả về dạng đầy đủ của link relative trong HTML source
					$sRet = _HTML_AbsoluteURL($sRet, $aURL[7] & '://' & $aURL[2] & $aURL[3], '', $aURL[7])
				EndIf
				If $aRetMode[0] < 4 Then
					Return SetError(0, $vResponse_StatusCode, $sRet)
				Else
					Local $aRet = [$g___retData[$g___LastSession][0], $sRet]
					Return SetError(0, $vResponse_StatusCode, $aRet)
				EndIf
			EndIf
			;------------------------------------------
		Case 6
			Local $aIPAndGeo = _GetIPAndGeoInfo()
			Return SetError(@error ? 8 : 0, $vResponse_StatusCode, $aIPAndGeo)
			;------------------------------------------
		Case 7, 8, 9
			Exit MsgBox(4096, 'Thông báo', '$iReturn 7, 8, 9 đã bị loại bỏ, xin vui lòng sửa lại code')
	EndSwitch
EndFunc



#Region <Quản lý các Session của _HttpRequest>
	Func _HttpRequest_SessionSet($sSessionNumber)
		If $sSessionNumber < 0 Or $sSessionNumber > $g___MaxSession_USE - 1 Then Exit MsgBox(4096, 'Lỗi', '$sSessionNumber chỉ có thể từ số từ 0 đến ' & $g___MaxSession_USE - 1)
		$g___LastSession = $sSessionNumber
	EndFunc

	Func _HttpRequest_SessionList()
		Local $aListSession[0], $iCounter = 0
		For $i = 0 To $g___MaxSession_USE - 1
			If $g___hOpen[$i] Then
				ReDim $aListSession[$iCounter + 1]
				$aListSession[$iCounter] = $i
				$iCounter += 1
			EndIf
		Next
		Return $aListSession
	EndFunc

	Func _HttpRequest_SessionClear($sSessionNumber = 0, $vClearProxy = False)
		$g___hCookieLast = ''
		If $sSessionNumber < 0 Or $sSessionNumber > $g___MaxSession_USE - 1 Then Exit MsgBox(4096, 'Lỗi', '$sSessionNumber chỉ có thể từ số từ 0 đến ' & $g___MaxSession_USE - 1)
		$g___retData[$sSessionNumber][0] = ''
		$g___retData[$sSessionNumber][1] = Binary('')
		$g___hCookie[$sSessionNumber] = ''
		If $g___hOpen[$sSessionNumber] Then $g___hOpen[$sSessionNumber] = 0 * _WinHttpCloseHandle2($g___hOpen[$sSessionNumber])
		If $g___ftpOpen[$sSessionNumber] Then $g___ftpOpen[$sSessionNumber] = 0 * _FTP_CloseHandle2($g___ftpOpen[$sSessionNumber])
		If $vClearProxy Then _HttpRequest_SetProxy()
		If $g___CookieJarPath Then _HttpRequest_CookieJarUpdateToFile()
	EndFunc
#EndRegion




Func _HttpRequest_Test($sData, $FilePath = Default, $iEncoding = Default, $iShellExecute = True)
	If Not $sData Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_Test', 'Không thể ghi dữ liệu vì $sData là rỗng'), '')
	If Not $FilePath Or IsKeyword($FilePath) Then $FilePath = @TempDir & '\Test.html'
	If StringRegExp($FilePath, '(?i)\.html$') Then $sData = StringRegExpReplace($sData, "(?i)<script>\h*?if \(document\.location\.protocol \!=\h*?[""']https:?[""']\h*?\).*?</script>", '', 1)
	If $iEncoding = Default Then $iEncoding = 128
	If IsBinary($sData) Or (StringRegExp($sData, '(?i)^0x[[:xdigit:]]+$') And Mod(StringLen($sData), 2) = 0) Then
		$iEncoding = 16
	ElseIf StringRegExp(_HttpRequest_DetectMIME($FilePath), '(?i)^(audio|image|video)\/') Then
		Return SetError(2, __HttpRequest_ErrNotify('_HttpRequest_Test', 'Vui lòng dùng _HttpRequest ở mode $iReturn = -2 hoặc $iReturn = 3 để lấy dữ liệu dạng Binary mới ghi được loại tập tin này'))
	EndIf
	Local $l___hOpen = FileOpen($FilePath, 2 + $iEncoding)
	FileWrite($l___hOpen, $sData)
	FileClose($l___hOpen)
	If $iShellExecute Or $iShellExecute = Default Then ShellExecute($FilePath)
EndFunc


Func _HttpRequest_CreateDataForm($a_FormItems) ;thêm dấu $ để nhận biết đó là 1 file, thêm dấu ~ để chuyển Unicode sang Ansi
	$g___Boundary = _BoundaryGenerator()
	Local $sData2Send = $g___Boundary & @CRLF, $vValue, $PatternError = 0
	;------------------------------------------------------------------------------------------
	If Not IsArray($a_FormItems) Then
		$PatternError = 1
	ElseIf UBound($a_FormItems, 0) < 1 And UBound($a_FormItems, 0) > 2 Then
		$PatternError = 1
	ElseIf UBound($a_FormItems, 0) = 1 Then
		For $i = 0 To UBound($a_FormItems) - 1
			If Not StringRegExp($a_FormItems[$i], '^([^=]+=|[^:]+: )') Then
				$PatternError = 1
				ExitLoop
			EndIf
		Next
	ElseIf UBound($a_FormItems, 0) = 2 And UBound($a_FormItems, 2) <> 2 Then
		$PatternError = 1
	EndIf
	;---------------------------
	If $PatternError = 1 Then
		Exit MsgBox(4096, 'Lỗi', 'Tham số của _HttpRequest_CreateDataForm phải là mảng có dạng như sau: [["key1", "value1"], ["key2", "value2"], ...] hoặc ["key1=value1", "key2=value2"], ...')
	EndIf
	;------------------------------------------------------------------------------------------
	If UBound($a_FormItems, 0) = 1 Then
		Local $ArrayTemp = $a_FormItems, $uBound = UBound($ArrayTemp), $aRegExp
		ReDim $a_FormItems[$uBound][2]
		For $i = 0 To $uBound - 1
			$ArrayTemp[$i] = StringRegExp($ArrayTemp[$i], '(?s)^([^\:\=]+)(?:\=|\:\s)(.*$)', 3)
			If @error Then Return SetError(2, __HttpRequest_ErrNotify('_HttpRequest_CreateDataForm', 'Lỗi không xác định'), '')
			$a_FormItems[$i][0] = ($ArrayTemp[$i])[0]
			$a_FormItems[$i][1] = ($ArrayTemp[$i])[1]
		Next
	EndIf
	;------------------------------------------------------------------------------------------
	If UBound($a_FormItems, 0) = 2 Then
		Local $l__uBound = UBound($a_FormItems) - 1
		For $i = 0 To $l__uBound
			$vValue = $a_FormItems[$i][1]
			Switch StringLeft($a_FormItems[$i][0], 1)
				Case '$'
					If FileExists($vValue) Then
						If StringRegExp($vValue, '^[^\\]+\.\w+$') Then $vValue = @ScriptDir & '\' & $vValue
						$a_FormItems[$i][0] = StringTrimLeft($a_FormItems[$i][0], 1)
						$vValue = _GetFileInfo($vValue)
						If @error Then Return SetError(3, __HttpRequest_ErrNotify('_HttpRequest_CreateDataForm', 'Không xác định được tập tin đầu vào #1'), '')
					ElseIf StringInStr($a_FormItems[$i][0], '/', 1, 1) Then
						Local $a_FormItems_Split = StringRegExp($a_FormItems[$i][0], '^\$([^\/]+)\/(.*)$', 3)
						If @error Then Return SetError(4, __HttpRequest_ErrNotify('_HttpRequest_CreateDataForm', 'Mẫu Key sai'), '')
						$a_FormItems[$i][0] = $a_FormItems_Split[0]
						Local $new_vValue[3] = [$a_FormItems_Split[1], _HttpRequest_DetectMIME($a_FormItems_Split[1]), (StringLeft($vValue, 2) = '0x' ? BinaryToString($vValue) : $vValue)]
						$vValue = $new_vValue
						$new_vValue = Null
					Else
						$a_FormItems[$i][0] = StringTrimLeft($a_FormItems[$i][0], 1)
						Local $new_vValue[3] = ['', 'application/octet-stream', '']
						$vValue = $new_vValue
						$new_vValue = Null
					EndIf
				Case '~'
					$a_FormItems[$i][0] = StringTrimLeft($a_FormItems[$i][0], 1)
					$vValue = _Utf8ToAnsi($vValue)
				Case Else
					If StringRegExp($vValue, '^\@[^\r\n]{1,200}\.\w+$') Then
						$vValue = StringTrimLeft($vValue, 1)
						If StringRegExp($vValue, '^[^\\]+\.\w+$') Then $vValue = @ScriptDir & '\' & $vValue
						$vValue = _GetFileInfo($vValue)
						If @error Then Return SetError(5, __HttpRequest_ErrNotify('_HttpRequest_CreateDataForm', 'Không xác định được tập tin đầu vào #2'), '')
					EndIf
			EndSwitch
			;------------------------------------------------------------------------------------------
			$sData2Send &= 'Content-Disposition: form-data; name="' & $a_FormItems[$i][0] & '"'
			If UBound($vValue) > 2 Then
				$sData2Send &= '; filename="' & $vValue[0] & '"' & @CRLF & 'Content-Type: ' & $vValue[1] & @CRLF & @CRLF & $vValue[2]
			Else
				$sData2Send &= @CRLF & @CRLF & $vValue
			EndIf
			;------------------------------------------------------------------------------------------
			$sData2Send &= @CRLF & $g___Boundary & @CRLF
		Next
	Else
		Return SetError(6, __HttpRequest_ErrNotify('_HttpRequest_CreateDataForm', '$a_FormItems phải là mảng 1D hoặc 2D Array'), '')
	EndIf
	;------------------------------------------------------------------------------------------
	;$sData2Send = StringRegExpReplace($sData2Send, '(?im)^(Content-Disposition: form-data; name=")"(.*?"\s*?;\s*?filename=)', '${1}${2}')
	;$sData2Send = StringRegExpReplace($sData2Send, '(?im)(Content-Type\s*?:\s*?.*)"$', '${1}')
	;------------------------------------------------------------------------------------------
	Return StringTrimRight($sData2Send, 2) & '--'
EndFunc


Func _HttpRequest_ErrorNotify($___ErrorNotify = True)
	If $___ErrorNotify = Default Then $___ErrorNotify = True
	$g___ErrorNotify = $___ErrorNotify
EndFunc


Func _HttpRequest_SetTimeout($__TimeOut = Default)
	If StringIsDigit($__TimeOut) Then $__TimeOut = Number($__TimeOut)
	If Not IsNumber($__TimeOut) Or $__TimeOut = Default Or $__TimeOut < 0 Then $__TimeOut = 30000
	$g___TimeOut = $__TimeOut
EndFunc


Func _HttpRequest_SetHotkeyStopRequest($__sHotKeyCancelReadWrite = '')
	If Not $__sHotKeyCancelReadWrite Or $__sHotKeyCancelReadWrite = Default Then $__sHotKeyCancelReadWrite = ''
	If $__sHotKeyCancelReadWrite Then
		If $g___HotkeySet Then HotKeySet($g___HotkeySet)
		HotKeySet($__sHotKeyCancelReadWrite, '__HttpRequest_CancelReadWrite')
		$g___HotkeySet = $__sHotKeyCancelReadWrite
	Else
		HotKeySet($__sHotKeyCancelReadWrite)
	EndIf
EndFunc


Func _HttpRequest_SetProxy($__Proxy = '', $___ProxyUserName = '', $___ProxyPassword = '', $___ProxyBypass = '', $iSession = Default)
	If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
	$__Proxy = StringStripWS($__Proxy, 8)
	If $__Proxy Then
		If Not StringRegExp($__Proxy, ':\d+$') Then Exit MsgBox(4096, 'Lỗi', 'Chưa set Port cho Proxy. Ví dụ mẫu Proxy đúng:' & @CR & @CR & '127.0.0.1:80')
		$g___hProxy[$iSession][3] = (($___ProxyUserName And Not IsKeyword($___ProxyUserName)) ? $___ProxyUserName : '')
		$g___hProxy[$iSession][4] = (($___ProxyPassword And Not IsKeyword($___ProxyPassword)) ? $___ProxyPassword : '')
		$g___hProxy[$iSession][2] = (($___ProxyBypass And Not IsKeyword($___ProxyBypass)) ? $___ProxyBypass : '')
		$g___hProxy[$iSession][0] = (($__Proxy And Not IsKeyword($__Proxy)) ? $__Proxy : '')
	Else
		$g___hProxy[$iSession][0] = ''
	EndIf
EndFunc


Func _HttpRequest_CheckProxyLive($__sProxy)
	Local $__RQ = _HttpRequest('2|%' & $__sProxy, 'http://httpbin.org/get')
	If Not @error And $__RQ And StringRegExp($__RQ, '"origin"\h*?:\h*?"[\d\.]{7,}"') Then Return True
	Return SetError(1, '', False)
EndFunc


Func _HttpRequest_SetUserAgent($___sUserAgent = Default, $iSession = Default)
	If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
	Local $BkUserAgent = $g___UserAgent[$iSession]
	If $___sUserAgent And Not IsKeyword($___sUserAgent) Then
		$g___UserAgent[$iSession] = $___sUserAgent
	Else
		$g___UserAgent[$iSession] = $g___defUserAgent
	EndIf
	Return $BkUserAgent
EndFunc


Func _HttpRequest_SetAuthorization($___sUserName = '', $___sPassword = '', $iSession = Default)
	If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
	Local $___sbkUP = $g___hCredential[$iSession][0] & ':' & $g___hCredential[$iSession][1]
	If IsKeyword($___sUserName) Then $___sUserName = ''
	If IsKeyword($___sPassword) Then $___sPassword = ''
	If $___sPassword == '' And StringInStr($___sUserName, ':', 1, 1) Then
		Local $aSplitUP = StringSplit($___sUserName, ':')
		$___sUserName = $aSplitUP[1]
		$___sPassword = $aSplitUP[2]
	EndIf
	$g___hCredential[$iSession][0] = $___sUserName
	$g___hCredential[$iSession][1] = $___sPassword
	Return $___sbkUP
EndFunc


Func _HttpRequest_QueryHeaders($iQueryFlag = Default, $iIndex = 0, $iSession = Default)
	If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
	If Not $g___hRequest[$iSession] Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_QueryHeaders', 'Handle của request đã hết hạn'), '')
	Select
		Case $iQueryFlag = Default Or $iQueryFlag = ''
			If $iSession = Default Then
				Return $g___retData[$g___LastSession][0]
			Else
				$iQueryFlag = 22
				ContinueCase
			EndIf
		Case StringIsDigit($iQueryFlag) Or IsNumber($iQueryFlag)
			Local $vRet = _WinHttpQueryHeaders2($g___hRequest[$iSession], $iQueryFlag == -1 ? 0x80000000 + 22 : $iQueryFlag, $iIndex)
			If @error Then Return SetError(2, __HttpRequest_ErrNotify('_HttpRequest_QueryHeaders', 'Truy vấn Response Headers thất bại'), '')
			Return $vRet & ($iQueryFlag == -1 ? '   ' & $g___sData2Send : '')
		Case Else
			If $iQueryFlag = 'Cookie' Or $iQueryFlag = 'Set-Cookie' Then
				Local $sCookie = _GetCookie($g___retData[$g___LastSession][0])
				If @error Then Return SetError(3, __HttpRequest_ErrNotify('_HttpRequest_QueryHeaders', 'Không truy vấn được Cookies từ Response Headers'), '')
				Return $sCookie
			Else
				Local $aResponseHeaders = StringRegExp($g___retData[$g___LastSession][0], '(?m)^\h*?\Q' & $iQueryFlag & '\E\h*?:\h*(.+)$', 1)
				If @error Then Return SetError(4, __HttpRequest_ErrNotify('_HttpRequest_QueryHeaders', 'Truy vấn ' & $iQueryFlag & ' từ Response Headers thất bại'), '')
				Return $aResponseHeaders[0]
			EndIf
	EndSelect
EndFunc


Func _HttpRequest_QueryData($iReadingMode = Default, $iSession = Default, $CallBackFunc_Progress = '') ; 0 ANSI, 1 UTF8, 2 Binary
	If $iReadingMode = Default Then $iReadingMode = 1
	If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
	If Not $g___hRequest[$iSession] Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_QueryData', 'Handle của request này đã hết hạn'), '')
	Local $outData = _WinHttpReadData_Ex($g___hRequest[$iSession], $CallBackFunc_Progress)
	If $outData == '' Then
		If $iReadingMode = 2 Then Return $g___retData[$g___LastSession][1]
		Return BinaryToString($g___retData[$g___LastSession][1], $iReadingMode = 1 ? 4 : 1)
	Else
		If StringRegExp(BinaryMid($outData, 1, 1), '(?i)0x(1F|08|8B)') Then $outData = __Gzip_Uncompress($outData)
		If $iReadingMode = 2 Then Return $outData
		Return BinaryToString($outData, $iReadingMode = 0 ? 1 : 4)
	EndIf
EndFunc


Func _HttpRequest_GetSize($iURL)
	Local $sHeader = _HttpRequest(1, $iURL, '', '', '', 'Range: bytes=0-0')
	If @error Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_GetSize', 'Gửi request với header Range thất bại'), Null)
	$sHeader = StringRegExp($sHeader, '(?im)^Content-Range.*?(\d+)$', 1)
	If @error Then Return SetError(2, __HttpRequest_ErrNotify('_HttpRequest_GetSize', 'Không tìm thấy header Content-Range từ Response'), Null)
	Return Number($sHeader[0])
EndFunc


Func _HttpRequest_FileSplitSize($iSize_or_URL, $iPart = Default, $iOffset = Default)
	If Not $iPart Or $iPart = Default Then $iPart = 8
	If Not $iOffset Or $iOffset = Default Then $iOffset = 0
	If Not StringIsDigit($iSize_or_URL) Then
		$iSize_or_URL = _HttpRequest_GetSize($iSize_or_URL)
		If $iSize_or_URL = 0 Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_FileSplitSize', 'Request lấy độ lớn của tập tin thất bại'), 0)
	EndIf
	;--------------------------------------------------------------------------------------------------------------------------------------
	If $iOffset And $iOffset * $iPart > $iSize_or_URL Then Return SetError(2, __HttpRequest_ErrNotify('_HttpRequest_FileSplitSize', '$iOffset đã nạp khiến phần chia nhỏ bị sai'), 0)
	;--------------------------------------------------------------------------------------------------------------------------------------
	Local $asPart[$iPart][2]
	Local $nPart = Floor($iSize_or_URL / $iPart)
	For $i = 0 To $iPart - 1
		$asPart[$i][0] = $i * $nPart + $iOffset
		$asPart[$i][1] = ($i + 1) * $nPart - 1 + $iOffset
	Next
	Local $nMod = Mod($iSize_or_URL, $iPart)
	If $nMod Then
		Local $nCount = 0
		For $i = 0 To $iPart - 1
			$asPart[$i][0] += $nCount
			If $nCount < $nMod Then $nCount += 1
		Next
		$nCount = 1
		For $i = 0 To $iPart - 1
			$asPart[$i][1] += $nCount
			If $nCount < $nMod Then $nCount += 1
		Next
	EndIf
	Local $aRange[$iPart]
	For $i = 0 To $iPart - 1
		If $iOffset > 0 And $i = $iPart - 1 Then $asPart[$i][1] -= $iOffset
		$aRange[$i] = 'Range: bytes=' & $asPart[$i][0] & '-' & $asPart[$i][1]
	Next
	Return SetError(0, $iSize_or_URL, $aRange)
EndFunc


Func _HttpRequest_SearchHiddenValues($iSourceHtml_or_URL, $iKeySearch = '', $iURIEncodeValue = True, $iType = Default)
	;$iKeySearch tách các KeyName bằng dấu |
	; $iType: hidden, text, hidden|text. default: hidden
	If $iType = Default Then $iType = 'hidden'
	If Not $iKeySearch Or IsKeyword($iKeySearch) Then $iKeySearch = ''
	If $iKeySearch Then $iKeySearch = StringSplit($iKeySearch, '|')
	;--------------------------------------------------------------------------------------------------------------------------------------
	If StringRegExp($iSourceHtml_or_URL, '(?i)^https?://') And Not StringRegExp($iSourceHtml_or_URL, '[\r\n]') Then
		$iSourceHtml_or_URL = _HttpRequest(2, $iSourceHtml_or_URL)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_SearchHiddenValues', 'Request lấy source thất bại'), '')
	EndIf
	;--------------------------------------------------------------------------------------------------------------------------------------
	Local $aInput = StringRegExp($iSourceHtml_or_URL, '(?i)<input (.*?type=\\?["''](?:' & $iType & ')\\?[''"] [\S\s]*?)\/?>', 3)
	If @error Then Return SetError(2, __HttpRequest_ErrNotify('_HttpRequest_SearchHiddenValues', 'Không tìm thấy Hidden Values'), '')
	$aInput = __ArrayDuplicate($aInput)
	;--------------------------------------------------------------------------------------------------------------------------------------
	Local $vName, $vValue, $_vName, $_vValue, $sRet, $isKeyExists, $aRet[0][2], $aCounter = 0
	If IsObj($g___oDicHiddenSearch) Then
		$g___oDicHiddenSearch.RemoveAll
	Else
		$g___oDicHiddenSearch = ObjCreate("Scripting.Dictionary")
		$g___oDicHiddenSearch.CompareMode = 1
	EndIf
	If @error Then Return SetError(3, __HttpRequest_ErrNotify('_HttpRequest_SearchHiddenValues', 'Không thể tạo Dictionary Object'), '')
	With $g___oDicHiddenSearch
		For $i = 0 To UBound($aInput) - 1
			$isKeyExists = 0
			$vName = StringRegExp($aInput[$i], '(?i)name\h*?=\h*?\\?[''"](.+?)\\?[''"]', 1)
			If @error Then ContinueLoop
			If ($iURIEncodeValue = True And .Exists(_URIEncode($vName[0]))) Or ($iURIEncodeValue = False And .Exists($vName[0])) Then
				$isKeyExists = 1
				For $k = 1 To 99
					If ($iURIEncodeValue = True And Not .Exists(_URIEncode($vName[0]) & '.' & $k)) Or ($iURIEncodeValue = False And Not .Exists($vName[0] & '.' & $k)) Then
						$vName[0] &= '.' & $k
						ExitLoop
					EndIf
				Next
			EndIf
			;-----------------------------------------
			If IsArray($iKeySearch) Then
				For $k = 1 To $iKeySearch[0]
					If StringRegExp($vName[0], '(?i)^\Q' & $iKeySearch[$k] & '\E\.?\d*?$') Then ExitLoop
				Next
				If $k > $iKeySearch[0] Then ContinueLoop
			EndIf
			;-----------------------------------------
			$vValue = StringRegExp($aInput[$i], '(?i)value\h*?=\h*?\\?[''"](.*?)\\?[''"]', 1)
			If @error Then ContinueLoop
			;-----------------------------------------
			$_vName = ($iURIEncodeValue ? _URIEncode($vName[0]) : $vName[0])
			$_vValue = ($iURIEncodeValue ? _URIEncode($vValue[0]) : $vValue[0])
			If $isKeyExists = 0 Then
				$sRet &= $_vName & '=' & $_vValue & '&'
				.Add($_vName & '.0', $_vValue)
				If $_vName <> $vName[0] Then .Add($vName[0] & '.0', $_vValue)
			EndIf
			.Add($_vName, $_vValue)
			If $_vName <> $vName[0] Then .Add($vName[0], $_vValue)
		Next
		;-----------------------------------------
		Local $aRet[.Count][2], $aCounter = 0
		For $oKey In $g___oDicHiddenSearch
			$aRet[$aCounter][0] = $oKey
			$aRet[$aCounter][1] = .Item($oKey)
			$aCounter += 1
		Next
		.Add('all_array', $aRet)
		;-------------
		.Add('all_string', StringTrimRight($sRet, 1))
		;-----------------------------------------
	EndWith
	Return $g___oDicHiddenSearch
EndFunc


Func _HttpRequest_OnlineCompiler($iCode, $iLanguage)
	;http://rextester.com/main
	;$iLanguage: 39 = Ada, 15 = Assembly, 38 = Bash, 1 = C#, 7 = C++ (gcc), 27 = C++ (clang), 28 = C++ (vc++), 6 = C (gcc), 26 = C (clang), 29 = C (vc), 36 = Client Side, 18 = Common Lisp, 30 = D, 41 = Elixir, 40 = Erlang, 3 = F#, 45 = Fortran, 20 = Go, 11 = Haskell, 4 = Java, 17 = Javascript, 43 = Kotlin, 14 = Lua, 33 = MySql, 23 = Node.js, 42 = Ocaml, 25 = Octave, 10 = Objective-C, 35 = Oracle, 9 = Pascal, 13 = Perl, 8 = Php, 34 = PostgreSQL, 19 = Prolog, 5 = Python, 24 = Python 3, 31 = R, 12 = Ruby, 21 = Scala, 22 = Scheme, 16 = Sql Server, 37 = Swift, 32 = Tcl, 2 = Visual Basic
	If TimerDiff($g___OnlineCompilerTimer) < 1500 Then
		Sleep(1500)
	Else
		$g___OnlineCompilerTimer = TimerInit()
	EndIf
	Local $jsonResult = _HttpRequest(2, 'https://rextester.com/rundotnet/api', 'LanguageChoice=' & $iLanguage & '&Program=' & _URIEncode($iCode))
	Local $aResult = StringRegExp($jsonResult, '(?i)^\{"Warnings":(null|".*?"),"Errors":(null|".*?"),"Result":"(.*?)(?:\\[rn]){0,}","Stats"', 3)
	If @error Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_OnlineCompiler', 'Compile Online thất bại'), '')
	If $aResult[0] <> 'null' Or $aResult[1] <> 'null' Then Return SetError(2, '', $jsonResult)
	Return $aResult[2]
EndFunc


Func _HttpRequest_BypassCloudflare($URL_in, $iTimeout = Default)
	If $iTimeout < 10000 Or $iTimeout = Default Or Not $iTimeout Then $iTimeout = 10000
	Local $aURL_in = StringRegExp($URL_in, '(?i)^(https?://)([^\/]+)', 3)
	If @error Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_BypassCloudflare', 'URL đầu vào không chính xác'), '')
	;-------------------------------------------------------------------------------------------------------------------
	Local $sourceHtml = _HttpRequest(2, $URL_in)
	If @error Then Return SetError(2, __HttpRequest_ErrNotify('_HttpRequest_BypassCloudflare', 'Request lấy Html thất bại'), '')
	;-------------------------------------------------------------------------------------------------------------------
	Local $URL_out = '', $Bypass_CF, $jschl_answer
	#Region Phân loại
		If StringInStr($sourceHtml, 'src="/cdn-cgi/scripts/cf.challenge.js"', 1, 1) And StringInStr($sourceHtml, 'g-recaptcha-response', 1, 1) Then
			Local $id_data_ray = StringRegExp($sourceHtml, '(?i)data-ray="(.+?)"', 1)
			If @error Then Return SetError(3, __HttpRequest_ErrNotify('_HttpRequest_BypassCloudflare', 'Không tìm thấy data-ray từ Html'), '')
			Local $g_recaptcha_response = _IE_RecaptchaBox($URL_in)
			If @error Then Return SetError(4, __HttpRequest_ErrNotify('_HttpRequest_BypassCloudflare', 'Giải ReCaptcha thất bại'), '')
			_HttpRequest_ConsoleWrite('> [CloudFlare] Đã nhận được g-recaptcha-response: ' & $g_recaptcha_response & @CRLF & @CRLF)
			$URL_out = $aURL_in[0] & $aURL_in[1] & '/cdn-cgi/l/chk_captcha?id=' & $id_data_ray[0] & '&g-recaptcha-response=' & $g_recaptcha_response
		Else
			$sourceHtml = StringReplace(StringReplace($sourceHtml, '={"', '.', 1, 1), '":', '+=', 1, 1)
			;-------------------------------------------------------------------------------------------------------------------
			Local $number_jschl_math = StringRegExp($sourceHtml, '\.\w+([\+\-\*\/])=([^\w\;\}]+)[;\}]', 3)
			If @error Then Return SetError(5, __HttpRequest_ErrNotify('_HttpRequest_BypassCloudflare', 'Không tìm được number_jschl_math từ Html'), '')
			For $i = 1 To UBound($number_jschl_math) - 1 Step 2
				$jschl_answer = '(' & $jschl_answer & $number_jschl_math[$i - 1] & $number_jschl_math[$i] & ')'
			Next
			$jschl_answer = Call('Execute', StringReplace(StringReplace(StringReplace(StringReplace($jschl_answer, '+!![]', '+1'), '!+[]', '+1'), '+[]', '+0'), '+(+', '&('))
			$jschl_answer = Round(StringFormat('%.10f', $jschl_answer), 10) ; Fixed Number
			$jschl_answer = $jschl_answer + StringLen($aURL_in[1]) ; + len Domain
			;-------------------------------------------------------------------------------------------------------------------
			Local $jschl_vc = StringRegExp($sourceHtml, '(?is)name="jschl_vc" .*?value="(.*?)".*?name="pass" .*?value="(.*?)"', 3)
			If @error Then Return SetError(6, __HttpRequest_ErrNotify('_HttpRequest_BypassCloudflare', 'Không tìm được jschl_vc từ Html'), '')
			;-------------------------------------------------------------------------------------------------------------------
			Local $challenge_form = StringRegExp($sourceHtml, '(?i)"challenge-form" action\h?=\h?"\/?([^"]+)"', 1)
			If @error Then Return SetError(7, __HttpRequest_ErrNotify('_HttpRequest_BypassCloudflare', 'Không tìm được challenge-form từ Html'), '')
			;-------------------------------------------------------------------------------------------------------------------
			$URL_out = $aURL_in[0] & $aURL_in[1] & '/' & $challenge_form[0] & '?jschl_vc=' & $jschl_vc[0] & '&pass=' & _URIEncode($jschl_vc[1]) & '&jschl_answer=' & $jschl_answer
			;-------------------------------------------------------------------------------------------------------------------
			_HttpRequest_ConsoleWrite('> [CloudFlare] Hãy chờ 5 giây ...')
			For $i = 1 To 40
				Sleep(100)
				ConsoleWrite('.')
			Next
			ConsoleWrite(@CRLF & @CRLF)
		EndIf
		;-------------------------------------------------------------------------------------------------------------------
		Local $sTimer = TimerInit()
		Do
			If TimerDiff($sTimer) > $iTimeout Then Return SetError(8, __HttpRequest_ErrNotify('_HttpRequest_BypassCloudflare', 'Timeout - Vượt CloudFlare thất bại'), '')
			Sleep(200)
			$Bypass_CF = StringRegExp(_HttpRequest(1, $URL_out, '', '', $URL_in & (StringRight($URL_in, 1) == '/' ? '' : '/')), '(?i)(cf_clearance=[^;]+)', 1)
		Until Not @error
		ConsoleWrite('> [CloudFlare] Bypass Cookie : ' & $Bypass_CF[0] & @CRLF & @CRLF)
		Return 'cf_clearance=' & $Bypass_CF[0] & ';'
	#EndRegion
EndFunc


Func _HttpRequest_SetCookieRemeber($iRemember = True)
	$g___hCookieRemember = $iRemember
EndFunc


Func _HttpRequest_ConsoleWrite($sString)
	ConsoleWrite($g___ConsoleForceANSI ? __RemoveVietMarktical($sString) : _Utf8ToAnsi($sString))
EndFunc


Func _HttpRequest_MsgBox($iFlag, $iTitle, $iText, $iTimeout = 0)
	Run(@AutoItExe & ' /AutoIt3ExecuteLine "MsgBox(' & $iFlag & ', ''' & StringReplace($iTitle, '"', '""') & ''', ''' & StringReplace($iText, '"', '""') & ''', ' & $iTimeout & ')"')
EndFunc


Func _HttpRequest_ReduceMem()
	Local $ahProc = DllCall('kernel32.dll', 'int', 'OpenProcess', 'int', 0x1F0FFF, 'int', False, 'int', @AutoItPID)
	If @error Or Not IsArray($ahProc) Then Return SetError(1)
	DllCall("psapi.dll", 'int', 'EmptyWorkingSet', 'long', $ahProc[0])
	DllCall('kernel32.dll', 'int', 'CloseHandle', 'int', $ahProc[0])
EndFunc


Func _HttpRequest_SetRoot($___BaseURL, $iSession = Default)
	If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
	Local $___bkBaseURL = $g___sBaseURL[$iSession]
	$g___sBaseURL[$iSession] = $___BaseURL
	Return $___bkBaseURL
EndFunc


Func _HttpRequest_ProxyNova_GetListCountries()
	Local $aCountriesList1D = StringRegExp(_HttpRequest(2, 'https://www.proxynova.com/proxy-server-list/'), '(?i)<option value="(\w+)">([^\(]+)\h*?\((\d+)\)</option>', 3)
	If @error Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_ProxyNova_GetListCountries', 'Không lấy được danh sách các nước từ trang proxynova.com'), '')
	Local $aCountriesList2D[1][3] = [['COUNTRY', 'ABBR', 'QUANTITY']], $iCounter = 1
	For $i = 0 To UBound($aCountriesList1D) - 1 Step 3
		If $aCountriesList1D[$i + 2] > 0 Then
			ReDim $aCountriesList2D[$iCounter + 1][3]
			$aCountriesList2D[$iCounter][0] = $aCountriesList1D[$i + 1]
			$aCountriesList2D[$iCounter][1] = $aCountriesList1D[$i + 0]
			$aCountriesList2D[$iCounter][2] = $aCountriesList1D[$i + 2]
			$iCounter += 1
		EndIf
	Next
	$aCountriesList2D[0][0] = $iCounter
	Return $aCountriesList2D
EndFunc


;$iProxyAnonymity: Transparent, Anonymous, Elite
;							+ Transparent - target server knows your IP address and it knows that you're using a proxy.
;							+ Anonymous - target server does not know your IP address, but it knows that you're using a proxy.
;							+ Elite - target server does not know your IP address, or that the request is relayed through a proxy server.
Func _HttpRequest_ProxyNova_GetListProxies($iCountrySelect = Default, $iProxySpeedLimit = Default, $iProxyUptimeSelect = Default, $iProxyAnonymity = Default)
	If $iProxySpeedLimit = Default Then $iProxySpeedLimit = 200
	If $iProxyUptimeSelect = Default Then $iProxyUptimeSelect = 20
	If $iCountrySelect = Default Then $iCountrySelect = ''
	If $iProxyAnonymity = Default Then $iProxyAnonymity = ''
	If $iCountrySelect Then $iCountrySelect = 'country-' & $iCountrySelect
	Local $aListProxy1D = _HttpRequest('_2', 'https://www.proxynova.com/proxy-server-list/' & $iCountrySelect)
	$aListProxy1D = StringRegExp($aListProxy1D, '(?i)(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\h+(\d+)[\h\r\n]+(\d+)\h+ms\h+(\d+)\%\h+\(\d+\)\h+(.+?)\h+(Transparent|Anonymous|Elite)', 3)
	If @error Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_ProxyNova_GetListProxies', 'Không lấy được danh sách Proxy'), '')
	Local $aListProxy2D[1][5] = [['PROXY:PORT', 'SPEED (ms)', 'UPTIME (%)', 'COUNTRY - PROVINCE', 'ANONYMITY']], $iCounter = 1
	For $i = 0 To UBound($aListProxy1D) - 1 Step 6
		If $aListProxy1D[$i + 2] < $iProxySpeedLimit Then ContinueLoop
		If $aListProxy1D[$i + 3] < $iProxyUptimeSelect Then ContinueLoop
		If $iProxyAnonymity And $aListProxy1D[$i + 5] <> $iProxyAnonymity Then ContinueLoop
		ReDim $aListProxy2D[$iCounter + 1][5]
		For $j = 0 To 4
			$aListProxy2D[$iCounter][$j] = $aListProxy1D[$i + $j + 1]
		Next
		$aListProxy2D[$iCounter][0] = $aListProxy1D[$i] & ':' & $aListProxy2D[$iCounter][0]
		$iCounter += 1
	Next
	Return $aListProxy2D
EndFunc



;===========================================================



Func _BoundaryGenerator()
	Local $sData = ""
	For $i = 1 To 12
		$sData &= Random(1, 9, 1)
	Next
	Return ('-----------------------------' & $sData)
EndFunc


Func _Data2SendEncode($_Key1, $_Value1 = '', $_Key2 = '', $_Value2 = '', $_Key3 = '', $_Value3 = '', $_Key4 = '', $_Value4 = '', $_Key5 = '', $_Value5 = '', $_Key6 = '', $_Value6 = '', $_Key7 = '', $_Value7 = '', $_Key8 = '', $_Value8 = '', $_Key9 = '', $_Value9 = '', $_Key10 = '', $_Value10 = '', $_Key11 = '', $_Value11 = '', $_Key12 = '', $_Value12 = '', $_Key13 = '', $_Value13 = '', $_Key14 = '', $_Value14 = '', $_Key15 = '', $_Value15 = '', $_Key16 = '', $_Value16 = '', $_Key17 = '', $_Value17 = '', $_Key18 = '', $_Value18 = '', $_Key19 = '', $_Value19 = '', $_Key20 = '', $_Value20 = '')
	Local $sResult = '', $sKey
	If @NumParams = 1 Then
		Local $sData2Send = $_Key1
		Local $aData2Send = StringRegExp($sData2Send, '(?:^|&)([^=]+=?=?)(?:=)([^&]*)', 3), $uBound = UBound($aData2Send)
		If Mod($uBound, 2) Then Return $sData2Send
		For $i = 0 To $uBound - 1 Step 2
			If Not StringRegExp($aData2Send[$i], '\%\w\w?') Then $aData2Send[$i] = _URIEncode($aData2Send[$i])
			If Not StringRegExp($aData2Send[$i + 1], '\%\w\w?') Then $aData2Send[$i + 1] = _URIEncode($aData2Send[$i + 1])
			$sResult &= $aData2Send[$i] & '=' & $aData2Send[$i + 1] & '&'
		Next
	Else
		For $i = 1 To 20
			$sKey = Eval('_Key' & $i)
			If $sKey == '' Then ExitLoop
			$sResult &= _URIEncode($sKey) & '=' & _URIEncode(Eval('_Value' & $i)) & '&'
		Next
	EndIf
	Return StringTrimRight($sResult, 1)
EndFunc


Func _Utf8ToAnsi($sData)
	Return BinaryToString(StringToBinary($sData, 4), 1)
EndFunc


Func _AnsiToUtf8($sData)
	Return BinaryToString(StringToBinary($sData, 1), 4)
EndFunc


Func _URIEncode($sData, $vUTF8 = True)
	If $sData == '' Then Return ''
	If $vUTF8 = True Then $sData = _Utf8ToAnsi($sData)
	Return _HTMLEncode($sData, '%', '', 2, False)
EndFunc


Func _URIDecode($sData, $vUTF8 = True, $iEntities = 0)
	If $sData == '' Then Return ''
	$sData = _HTMLDecode(StringReplace($sData, '+', ' ', 0, 1), '%', '', 2, True, $iEntities)
	If $vUTF8 Then $sData = _AnsiToUtf8($sData)
	Return $sData
EndFunc
Func __URIEncode($sData)
    ; Prog@ndy
    Local $aData = StringSplit(BinaryToString(StringToBinary($sData,4),1),"")
    Local $nChar
    $sData=""
    For $i = 1 To $aData[0]
        ; ConsoleWrite($aData[$i] & @CRLF)
        $nChar = Asc($aData[$i])
        Switch $nChar
            Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
                $sData &= $aData[$i]
            Case 32
                $sData &= "+"
            Case Else
                $sData &= "%" & Hex($nChar,2)
        EndSwitch
    Next
    Return $sData
EndFunc

Func __URIDecode($sData)
    ; Prog@ndy
    Local $aData = StringSplit(StringReplace($sData,"+"," ",0,1),"%")
    $sData = ""
    For $i = 2 To $aData[0]
        $aData[1] &= Chr(Dec(StringLeft($aData[$i],2))) & StringTrimLeft($aData[$i],2)
    Next
    Return BinaryToString(StringToBinary($aData[1],1),4)
EndFunc

Func _HTMLEncode($sData, $Escape_Character_Head = '\u', $Escape_Character_Tail = Default, $iHexLength = Default, $iPassSpace = True)
	If $sData == '' Then Return ''
	If $iHexLength = Default Then $iHexLength = 4
	If $Escape_Character_Tail = Default Then $Escape_Character_Tail = ''
	Local $Asc_or_AscW = ($iHexLength = 2 ? 'Asc' : 'AscW')
	Local $sResult = Call('Execute', '"' & StringReplace(StringRegExpReplace($sData, '([^\w\-\.\~' & ($iPassSpace ? '\h' : '') & '])', '" & "\' & $Escape_Character_Head & '" & Hex(' & $Asc_or_AscW & '("$1"), ' & $iHexLength & ') & "' & $Escape_Character_Tail), $Asc_or_AscW & '(""")', $Asc_or_AscW & '("""")', 0, 1) & '"')
	If $sResult == '' Then Return SetError(1, __HttpRequest_ErrNotify('_HTMLEncode', 'Encode thất bại'), $sData)
	Return $sResult
EndFunc


Func _HTMLDecode($sData, $Escape_Character_Head = '\u', $Escape_Character_Tail = Default, $iHexLength = Default, $isHexNumber = True, $iEntities = 1)
	If $sData == '' Then Return ''
	Switch $iEntities
		Case 1
			$sData = __HTML_Entities_Decode($sData, False)
		Case 2
			$sData = __HTML_Entities_Decode($sData, True)
	EndSwitch
	If StringRegExp($sData, '&#[[:xdigit:]]{2};') Then $sData = __HTML_RegexpReplace($sData, '&#', ';', '2', False)
	If StringRegExp($sData, '&#[[:xdigit:]]{3,4};') Then $sData = __HTML_RegexpReplace($sData, '&#', ';', '3,4', False)
	If $iHexLength = Default Then
		If StringRegExp($sData, '\Q' & $Escape_Character_Head & '\E\w{2}(\Q' & $Escape_Character_Head & '\E|$)') Then
			$iHexLength = 2
;~ 		ElseIf StringRegExp($sData, '\Q' & $Escape_Character_Head & '\E\w{4}(\Q' & $Escape_Character_Head & '\E|$)') Then
;~ 			$iHexLength = 4
		ElseIf $Escape_Character_Tail And $Escape_Character_Tail <> Default Then
			$iHexLength = '2,4'
		Else
			$iHexLength = '3,4'
		EndIf
	EndIf
	If $Escape_Character_Tail = Default Then $Escape_Character_Tail = ';?'
	Return __HTML_RegexpReplace($sData, $Escape_Character_Head, $Escape_Character_Tail, $iHexLength, $isHexNumber)
EndFunc


;===============================================================

Func _Cookie_JSON2SemicolonFormat($jsonCookie)
	Local $aCookie = StringRegExp($jsonCookie, '(?i)"name":"([^"]+)".*?"value":"(.*?)"', 3), $sCookie = ''
	If @error Then Return SetError(1, '', '')
	For $i = 0 To UBound($aCookie) - 1 Step 2
		$sCookie &= $aCookie[$i] & '=' & $aCookie[$i + 1] & ';'
	Next
	Return $sCookie
EndFunc


Func _Live_HTTP_Headers_Form2Array($iType = 0)
	Local $arr_form = StringRegExp(ClipGet(), '(?im)Content-Disposition: form-data; name="([^"]*?)"\R\R^(.*?)$', 3)
	Local $str_form = @TAB & @TAB & '['
	For $i = 0 To UBound($arr_form) - 1 Step 2
		If $iType = 0 Then
			$str_form &= '["' & $arr_form[$i] & '", "' & $arr_form[$i + 1] & '"], _' & @CRLF & @TAB & @TAB
		Else
			$str_form &= '"' & $arr_form[$i] & '=' & $arr_form[$i + 1] & '", _' & @CRLF & @TAB & @TAB
		EndIf
	Next
	ClipPut('Local $FormData = _' & @CRLF & StringTrimRight($str_form, 7) & ']')
EndFunc


Func _GetFreeProxy($sFilePathToExport = '', $iGetFastProxyList = False, $vIncludeDateInHeader = False)
	Local $aLink = StringRegExp(_HttpRequest(2, 'http://www.proxyserverlist24.top/'), "(?i)href='(http://www.proxyserverlist24.top/.*?([\d\-]+)-" & ($iGetFastProxyList ? 'fast' : 'free') & "-proxy-server-list.*?\.html)'", 1)
	If @error Then Return SetError(1, __HttpRequest_ErrNotify('_GetFreeProxy', 'Không tìm thấy đường dẫn lấy Proxy List'), '')
	Local $asProxy = StringRegExp(_HttpRequest(2, $aLink[0]), '(?ms)^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\:\d+\R.+\R^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\:\d+', 1)
	If @error Then Return SetError(2, __HttpRequest_ErrNotify('_GetFreeProxy', 'Không tách được danh sách Proxy'), '')
	If $sFilePathToExport Then
		Local $hFileOpen = FileOpen($sFilePathToExport, 2)
		FileWrite($hFileOpen, ($vIncludeDateInHeader ? StringReplace($aLink[1], '-', '', 0, 1) & @CRLF : '') & $asProxy[0])
		FileClose($hFileOpen)
	EndIf
	Return SetExtended(StringReplace($aLink[1], '-', '', 0, 1), StringSplit(StringStripCR($asProxy[0]), @LF))
EndFunc


Func _GetCertificateInfo($iSession = Default)
	If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
	If Not $g___hRequest[$iSession] Then Return SetError(1, __HttpRequest_ErrNotify('_GetCertificateInfo', 'Phải thực hiện request đến trang đích trước'), '')
	Local $tBuffer = _WinHttpQueryOptionEx2($g___hRequest[$iSession], 32)
	If @error Then Return SetError(2, __HttpRequest_ErrNotify('_HttpRequest_CertificateInfo', 'Yêu cầu phải là https mới lấy được thông tin Certificate'), '')
	Local $tCertInfo = DllStructCreate("dword ExpiryTime[2]; dword StartTime[2]; ptr SubjectInfo; ptr IssuerInfo; ptr ProtocolName; ptr SignatureAlgName; ptr EncryptionAlgName; dword KeySize", DllStructGetPtr($tBuffer))
	Return DllStructGetData(DllStructCreate("wchar[256]", DllStructGetData($tCertInfo, "IssuerInfo")), 1)
EndFunc


Func _GetNameDNS($iSession = Default)
	If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
	If Not $g___hRequest[$iSession] Then Return SetError(1, __HttpRequest_ErrNotify('_GetNameDNS', 'Phải thực hiện request đến trang đích trước'), '')
	Local $tBuffer, $pCert_Context, $tCert_Info, $tCert_Encoding, $tCert_Ext, $aCall
	$tBuffer = DllStructCreate("ptr")
	DllCall($dll_WinHttp, "bool", 'WinHttpQueryOption', "handle", $g___hRequest[$iSession], "dword", 78, "struct*", $tBuffer, "dword*", DllStructGetSize($tBuffer))
	If @error Then Return SetError(2, __HttpRequest_ErrNotify('_GetNameDNS', 'Cài đặt option lấy DNS thất bại. Yêu cầu phải là https mới lấy được DNS'), '')
	$pCert_Context = DllStructGetData($tBuffer, 1)
	$tCert_Encoding = DllStructCreate("dword dwCertEncodingType; ptr pbCertEncoded; dword cbCertEncoded; ptr pCertInfo; handle hCertStore", $pCert_Context)
	If $g___CertInfo = '' Then
		$g___CertInfo = '«dw dwVersion; «dw SerialNumber_cbData; p SerialNumber_pbData»; «p SignatureAlgorithm_pszObjId; «dw SignatureAlgorithm_Parameters_cbData; p SignatureAlgorithm_Parameters_pbData»»; «dw Issuer_cbData; p Issuer_pbData»; «dw NotBefore_dwLowDateTime; dw NotBefore_dwHighDateTime»; «dw NotAfter_dwLowDateTime; dw NotAfter_dwHighDateTime»; «dw Subject_cbData; p Subject_pbData»; ««p SubjectPublicKeyInfo_Algorithm_pszObjId; «dw SubjectPublicKeyInfo_Parameters_cbData; p SubjectPublicKeyInfo_Parameters_pbData»»; «dw SubjectPublicKeyInfo_PublicKey_cbData; p ParametersSubjectPublicKeyInfo_pbData; dw SubjectPublicKeyInfo_PublicKey_cUnusedBits»»; «dw IssuerUniqueId_cbData; p IssuerUniqueId_pbData; dw IssuerUniqueId_cUnusedBits»; «dw dwSubjectUniqueId_cbData; p SubjectUniqueId_pbData; dw SubjectUniqueId_cUnusedBits»; dw cExtension; p rgExtension»;'
		$g___CertInfo = StringReplace(StringReplace(StringReplace(StringReplace($g___CertInfo, 'dw ', 'dword '), 'p ', 'ptr '), '«', 'struct;'), '»', ';endstruct')
	EndIf
	$tCert_Info = DllStructCreate($g___CertInfo, DllStructGetData($tCert_Encoding, 'pCertInfo'))
	$aCall = DllCall("Crypt32.dll", "ptr", "CertFindExtension", "str", "2.5.29.17", "dword", DllStructGetData($tCert_Info, 'cExtension'), "ptr", DllStructGetData($tCert_Info, 'rgExtension'))
	If @error Then Return SetError(3, __HttpRequest_ErrNotify('_GetNameDNS', 'Không tìm thấy Chứng nhận của DNS'), '')
	$tCert_Ext = DllStructCreate("struct;ptr pszObjId;bool fCritical;struct;dword Value_cbData;ptr Value_pbData;endstruct;endstruct;", $aCall[0])
	$aCall = DllCall("Crypt32.dll", "int", "CryptFormatObject", "dword", 1, "dword", 0, "dword", 1, "ptr", 0, "ptr", DllStructGetData($tCert_Ext, 'pszObjId'), "ptr", DllStructGetData($tCert_Ext, 'Value_pbData'), "dword", DllStructGetData($tCert_Ext, 'Value_cbData'), 'wstr', "", "dword*", 65536)
	If @error Then Return SetError(4, __HttpRequest_ErrNotify('_GetNameDNS', 'Không định dạng được Chứng nhận của DNS'), '')
	DllCall("Crypt32.dll", "dword", "CertFreeCertificateContext", "ptr", $pCert_Context)
	Return StringReplace($aCall[8], 'DNS Name=', '')
EndFunc


Func _GetCookie($sHeader = '', $iSession = Default, $iTrimCookie = True, $Excluded_Values = '')
	If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
	If $sHeader == '' Or $sHeader = Default Or ($sHeader And StringLeft($sHeader, 5) <> 'HTTP/') Then
		If $g___retData[$g___LastSession][0] Then
			$sHeader = $g___retData[$g___LastSession][0]
		ElseIf IsPtr($g___hRequest[$iSession]) Then
			$sHeader = _WinHttpQueryHeaders2($g___hRequest[$iSession], 22)
			If @error Or $sHeader == '' Then Return SetError(1, __HttpRequest_ErrNotify('_GetCookie', 'Không truy vấn được Response Headers'), '')
		EndIf
	EndIf
	Local $__aRH = StringRegExp($sHeader, '(?im)^Set-Cookie:\h*?([^=]+)=(?!deleted;)(.*)$', 3)
	If @error Or Not IsArray($__aRH) Then Return SetError(2, __HttpRequest_ErrNotify('_GetCookie', 'Không tìm thấy header Set-Cookie từ Response'), '')
	;----------------------------------------------------------------------------------
	Local $__sRH = '', $__uBound = UBound($__aRH)
	For $i = $__uBound - 2 To 0 Step -2
		If $__aRH[$i] == '' Or ($Excluded_Values And StringInStr('|' & $Excluded_Values & '|', '|' & StringStripWS($__aRH[$i], 3) & '|')) Then ContinueLoop
		$__sRH = $__aRH[$i] & '=' & $__aRH[$i + 1] & '; ' & $__sRH
		For $k = 0 To $i Step 2
			If $__aRH[$k] == $__aRH[$i] Then $__aRH[$k] = ''
		Next
	Next
	;----------------------------------------------------------------------------------
	If $iTrimCookie Then
		Local $aOptionalFilter = 'priority\h*?=\h*?(?:high|low)|Expires\h*?=\h*?|Path\h*?=\h*?|Domain\h*?=\h*?|Max-age\h*?=\h*?|SameSite\h*?=\h*?|HttpOnly|Secure'
		$__sRH = StringRegExpReplace(StringRegExpReplace($__sRH, '(?i);\h*?(' & $aOptionalFilter & ')([^;]*)', ';'), '(?:;\h?){2,}', ';')
	EndIf
	;----------------------------------------------------------------------------------
	Return $__sRH
EndFunc


Func _GetLocationRedirect($sHeader = '', $iIndex = -1, $iSession = Default)
	If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
	If Not $sHeader Or $sHeader = Default Or ($sHeader And StringLeft($sHeader, 5) <> 'HTTP/') Then
		If $g___retData[$g___LastSession][0] Then
			$sHeader = $g___retData[$g___LastSession][0]
		ElseIf $g___LocationRedirect Then
			Return $g___LocationRedirect
		ElseIf IsPtr($g___hRequest[$iSession]) Then
			$sHeader = _WinHttpQueryHeaders2($g___hRequest[$iSession], 22)
			If @error Or $sHeader == '' Then Return SetError(1, __HttpRequest_ErrNotify('_GetLocationRedirect', 'Không truy vấn được Response Headers'), '')
		Else
			Return SetError(2, __HttpRequest_ErrNotify('_GetLocationRedirect', 'Lỗi không xác định'), '')
		EndIf
	EndIf
	Local $__aRH = StringRegExp($sHeader, '(?im)^Location:\h?(.+)$', 3)
	If @error Or Not IsArray($__aRH) Then Return SetError(3, __HttpRequest_ErrNotify('_GetLocationRedirect', 'Không tìm thấy header Location từ Response'), '')
	Local $uBoundRH = UBound($__aRH) - 1
	If $iIndex < 0 Or $iIndex > $uBoundRH Or $iIndex = Default Or $iIndex == '' Then $iIndex = $uBoundRH
	Return StringStripWS($__aRH[$iIndex], 3)
EndFunc


Func _GetIPAndGeoInfo($iIP = Default)
	If $iIP = Default Then $iIP = $g___ServerIP
	If $iIP = '' Then Return SetError(1, __HttpRequest_ErrNotify('_GetIPAndGeoInfo', 'Không tìm thấy IP - Phải request đến trang đích trước khi sử dụng hàm này hoặc nhập 1 IP bạn biết vào'), '')
	Local $sHTML = _HttpRequest(2, 'https://gfx.robtex.com/ipinfo.js?ip=' & $iIP)
	Local $aInfo = [$iIP, 'country', 'city', 'asname', 'net', 'netdescr', 'as'], $regSource
	For $i = 1 To 6
		$regSource = StringRegExp($sHTML, '(?i)\(m\h*?==?\h*?"' & $aInfo[$i] & '"\)\h*?a.innerHTML\h*?=\h*?"\,?\h?(.*?)"', 1)
		If @error Then Return SetError(2, __HttpRequest_ErrNotify('_GetIPAndGeoInfo', 'Không tìm được thông tin từ IP'), $iIP)
		$aInfo[$i] = $regSource[0]
	Next
	Return $aInfo
EndFunc


Func _GetFileInfo($sFilePath, $vDataTypeReturn = 1) ; 2: Base64, 1: String, 0: Binary
	If Not FileExists($sFilePath) Then Return SetError(1, __HttpRequest_ErrNotify('_GetFileInfo', 'Đường dẫn tập tin không tồn tại'), '')
	If $vDataTypeReturn = Default Or $vDataTypeReturn == '' Then $vDataTypeReturn = 1
	Local $sFileName = StringRegExp($sFilePath, '[\\\/]([^\\\/]+\.\w+)$', 1)
	If @error Then Return SetError(2, __HttpRequest_ErrNotify('_GetFileInfo', 'Không tách được tên tập tin từ đường dẫn'), '')
	$sFileName = $sFileName[0]
	Local $hFileOpen = FileOpen($sFilePath, 16)
	If @error Then Return SetError(3, __HttpRequest_ErrNotify('_GetFileInfo', 'Không thể mở tập tin'), '')
	Local $sFileData = FileRead($hFileOpen)
	FileClose($hFileOpen)
	Local $sFileType = _HttpRequest_DetectMIME($sFileName)
	Switch $vDataTypeReturn
		Case 2
			$sFileData = _B64Encode($sFileData, 0, True, True)
			$sFileType &= ';base64'
		Case 1
			$sFileData = BinaryToString($sFileData)
	EndSwitch
	Local $aReturn[4] = [$sFileName, $sFileType, $sFileData, FileGetSize($sFilePath)]
	Return $aReturn
EndFunc


Func _GetHttpTime($sHttpTime = '')
	Local $tSystemTime = DllStructCreate('word Year;word Month;word DayOfWeek;word Day;word Hour;word Minute;word Second;word Milliseconds')
	Local $tTime = DllStructCreate("wchar[62]")
	If $sHttpTime Then
		DllStructSetData($tTime, 1, $sHttpTime)
		Local $aCall = DllCall($dll_WinHttp, "bool", 'WinHttpTimeToSystemTime', "struct*", $tTime, "struct*", $tSystemTime)
		If @error Or Not $aCall[0] Then Return SetError(3, __HttpRequest_ErrNotify('_GetHttpTime', 'Không thể gọi chức năng WinHttpTimeToSystemTime của WinHttp'), "")
		Local $aRet[6]
		For $i = 0 To 5
			$aRet[$i] = DllStructGetData($tSystemTime, $i + ($i < 2 ? 1 : 2))
		Next
		Return SetError(0, 0, $aRet)
	Else
		DllCall("kernel32.dll", "none", "GetSystemTime", "struct*", $tSystemTime)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('_GetHttpTime', 'Không thế truy vấn Time hệ thống'), "")
		Local $aCall = DllCall($dll_WinHttp, "bool", 'WinHttpTimeFromSystemTime', "struct*", $tSystemTime, "struct*", $tTime)
		If @error Or Not $aCall[0] Then Return SetError(2, __HttpRequest_ErrNotify('_GetHttpTime', 'Không thể gọi chức năng WinHttpTimeFromSystemTime của WinHttp'), "")
		Return DllStructGetData($tTime, 1)
	EndIf
EndFunc


Func _GetTimeStamp($Include_MSec = False, $sDateTime = Default) ;D/M/YYYY h:m:s
	If $sDateTime = Default Then
		Local $tSystemTime = DllStructCreate('struct;word Year;word Month;word Dow;word Day;word Hour;word Minute;word Second;word MSeconds;endstruct')
		DllCall("kernel32.dll", "none", "GetSystemTime", "struct*", $tSystemTime)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('_GetTimeStamp', 'GetSystemTime thất bại'), '')
		Local $aInfo[7] = [DllStructGetData($tSystemTime, "Day"), DllStructGetData($tSystemTime, "Month"), DllStructGetData($tSystemTime, "Year"), DllStructGetData($tSystemTime, "Hour"), DllStructGetData($tSystemTime, "Minute"), DllStructGetData($tSystemTime, "Second"), ($Include_MSec ? DllStructGetData($tSystemTime, "MSeconds") : 0)]
	Else
		Local $aInfo = StringRegExp($sDateTime, '\d+', 3)
		If @error Or UBound($aInfo) <> 6 Then Return SetError(2, __HttpRequest_ErrNotify('_GetTimeStamp', '$sDateTime không đúng định dạng'), '')
		ReDim $aInfo[7]
		$aInfo[6] = ($Include_MSec ? @MSEC : 0)
	EndIf
	$aInfo[2] -= ($aInfo[1] < 3 ? 1 : 0)
	Return ((Int(Int($aInfo[2] / 100) / 4) - Int($aInfo[2] / 100) + $aInfo[0] + Int(365.25 * ($aInfo[2] + 4716)) + Int(30.6 * (($aInfo[1] < 3 ? $aInfo[1] + 12 : $aInfo[1]) + 1)) - 2442110) * 86400 + ($aInfo[3] * 3600 + $aInfo[4] * 60 + $aInfo[5])) * ($Include_MSec ? 1000 : 1) + $aInfo[6]
EndFunc


Func _GetDateFromTimeStamp($iTimeStamp, $vLocalTime = False)
	Local $Msec = 0
	If StringLen($iTimeStamp) = 13 Then
		$iTimeStamp = Floor($iTimeStamp / 1000)
		$Msec = Mod($iTimeStamp, 1000)
	EndIf
	Local $iDayToAdd = Int($iTimeStamp / 86400), $iTimeVal = Mod($iTimeStamp, 86400)
	If $iTimeVal < 0 Then
		$iDayToAdd -= 1
		$iTimeVal += 86400
	EndIf
	Local $i_wFactor = Int((573371.75 + $iDayToAdd) / 36524.25), $i_bFactor = 2442113 + $iDayToAdd + $i_wFactor - Int($i_wFactor / 4), $i_cFactor = Int(($i_bFactor - 122.1) / 365.25), $i_dFactor = Int(365.25 * $i_cFactor), $i_eFactor = Int(($i_bFactor - $i_dFactor) / 30.6001), $aDatePart[3], $aTimePart[3]
	$aDatePart[2] = $i_bFactor - $i_dFactor - Int(30.6001 * $i_eFactor)
	$aDatePart[1] = $i_eFactor - 1 - 12 * ($i_eFactor - 2 > 11)
	$aDatePart[0] = $i_cFactor - 4716 + ($aDatePart[1] < 3)
	$aTimePart[0] = Int($iTimeVal / 3600)
	$iTimeVal = Mod($iTimeVal, 3600)
	$aTimePart[1] = Int($iTimeVal / 60)
	$aTimePart[2] = Mod($iTimeVal, 60)
	If $vLocalTime Then
		Local $tUTC = DllStructCreate('struct;word Year;word Month;word Dow;word Day;word Hour;word Minute;word Second;word MSeconds;endstruct')
		DllStructSetData($tUTC, "Month", $aDatePart[1])
		DllStructSetData($tUTC, "Day", $aDatePart[2])
		DllStructSetData($tUTC, "Year", $aDatePart[0])
		DllStructSetData($tUTC, "Hour", $aTimePart[0])
		DllStructSetData($tUTC, "Minute", $aTimePart[1])
		DllStructSetData($tUTC, "Second", $aTimePart[2])
		Local $tLocal = DllStructCreate('struct;word Year;word Month;word Dow;word Day;word Hour;word Minute;word Second;word MSeconds;endstruct')
		DllCall("kernel32.dll", "bool", "SystemTimeToTzSpecificLocalTime", "struct*", 0, "struct*", DllStructGetPtr($tUTC), "struct*", $tLocal)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('_GetDateFromTimeStamp', 'Chuyển SystemTime sang LocalTime thất bại'), '')
		$aDatePart[2] = DllStructGetData($tLocal, "Day")
		$aDatePart[1] = DllStructGetData($tLocal, "Month")
		$aDatePart[0] = DllStructGetData($tLocal, "Year")
		$aTimePart[0] = DllStructGetData($tLocal, "Hour")
		$aTimePart[1] = DllStructGetData($tLocal, "Minute")
		$aTimePart[2] = DllStructGetData($tLocal, "Second")
	EndIf
	Return StringFormat("%02d/%02d/%04d %02d:%02d:%02d", $aDatePart[2], $aDatePart[1], $aDatePart[0], $aTimePart[0], $aTimePart[1], $aTimePart[2]) & ($Msec > 0 ? ':' & $Msec : '')
EndFunc


Func _B64Encode($binaryData, $iLinebreak = 0, $safeB64 = False, $iRunByMachineCode = True, $iCompressData = False)
	If $binaryData == '' Then Return SetError(1, __HttpRequest_ErrNotify('_B64Encode', '$binaryData rỗng'), '')
	$iLinebreak = Number($iLinebreak)
	If $iLinebreak = Default Then $iLinebreak = 0
	If $safeB64 = Default Then $safeB64 = False
	If $iRunByMachineCode = Default Then $iRunByMachineCode = False
	;----------------------------------------------------------------------------------------
	If $iCompressData Then $binaryData = __LZNT_Compress($binaryData)
	If Not $iRunByMachineCode Then
		Local $lenData = StringLen($binaryData) - 2, $iOdd = Mod($lenData, 3), $spDec = '', $base64Data = ''
		For $i = 3 To $lenData - $iOdd Step 3
			$spDec = Dec(StringMid($binaryData, $i, 3))
			$base64Data &= $g___aChr64[$spDec / 64] & $g___aChr64[Mod($spDec, 64)]
		Next
		If $iOdd Then
			$spDec = BitShift(Dec(StringMid($binaryData, $i, 3)), -8 / $iOdd)
			$base64Data &= $g___aChr64[$spDec / 64] & ($iOdd = 2 ? $g___aChr64[Mod($spDec, 64)] & $g___sPadding & $g___sPadding : $g___sPadding)
		EndIf
	Else
		Local $tStruct = DllStructCreate("byte[" & BinaryLen($binaryData) & "]")
		DllStructSetData($tStruct, 1, $binaryData)
		Local $tsInt = DllStructCreate("int")
		Local $a_Call = DllCall("Crypt32.dll", "int", "CryptBinaryToString", "ptr", DllStructGetPtr($tStruct), "int", DllStructGetSize($tStruct), "int", 1, "ptr", 0, "ptr", DllStructGetPtr($tsInt))
		If @error Or Not $a_Call[0] Then Return SetError(2, __HttpRequest_ErrNotify('_B64Encode', 'Gọi chức năng CryptBinaryToString từ Crypt32.dll thất bại #1'), $binaryData)
		Local $tsChr = DllStructCreate("char[" & DllStructGetData($tsInt, 1) & "]")
		$a_Call = DllCall("Crypt32.dll", "int", "CryptBinaryToString", "ptr", DllStructGetPtr($tStruct), "int", DllStructGetSize($tStruct), "int", 1, "ptr", DllStructGetPtr($tsChr), "ptr", DllStructGetPtr($tsInt))
		If @error Or Not $a_Call[0] Then Return SetError(3, __HttpRequest_ErrNotify('_B64Encode', 'Gọi chức năng CryptBinaryToString từ Crypt32.dll thất bại #2'), $binaryData)
		Local $base64Data = StringStripWS(DllStructGetData($tsChr, 1), 8)
	EndIf
	If $iLinebreak Then $base64Data = StringRegExpReplace($base64Data, '(.{' & $iLinebreak & '})', '${1}' & @LF)
	If $safeB64 Then $base64Data = StringReplace(StringReplace($base64Data, '+', '-', 0, 1), '/', '_', 0, 1)
	Return $base64Data
EndFunc


Func _B64Decode($base64Data, $iRunByMachineCode = True, $iUnCompressData = False)
	If $base64Data == '' Then Return SetError(1, __HttpRequest_ErrNotify('_B64Decode', '$base64Data rỗng'), '')
	If $iRunByMachineCode = Default Then $iRunByMachineCode = False
	$base64Data = StringStripWS($base64Data, 8)
	;----------------------------------------------------------------------------------------
	If Not $iRunByMachineCode Then
		If Mod(StringLen($base64Data), 2) Then Return SetError(2, __HttpRequest_ErrNotify('_B64Decode', '$base64Data không phải là dữ liệu kiểu B64'), $base64Data)
		Local $aData = StringSplit($base64Data, ''), $binaryData = '0x', $iOdd = UBound(StringRegExp($base64Data, $g___sPadding, 3))
		For $i = 1 To $aData[0] - $iOdd * 2 Step 2
			$binaryData &= Hex((StringInStr($g___sChr64, $aData[$i], 1, 1) - 1) * 64 + StringInStr($g___sChr64, $aData[$i + 1], 1, 1) - 1, 3)
		Next
		If $iOdd Then $binaryData &= Hex(BitShift((StringInStr($g___sChr64, $aData[$i], 1, 1) - 1) * 64 + ($iOdd - 1) * (StringInStr($g___sChr64, $aData[$i + 1], 1, 1) - 1), 8 / $iOdd), $iOdd)
	Else
		Local $tStruct = DllStructCreate("int")
		Local $a_Call = DllCall("Crypt32.dll", "int", "CryptStringToBinary", "str", $base64Data, "int", 0, "int", 1, "ptr", 0, "ptr", DllStructGetPtr($tStruct, 1), "ptr", 0, "ptr", 0)
		If @error Or Not $a_Call[0] Then Return SetError(3, __HttpRequest_ErrNotify('_B64Decode', 'Gọi chức năng CryptStringToBinary từ Crypt32.dll thất bại #1'), $base64Data)
		Local $tsByte = DllStructCreate("byte[" & DllStructGetData($tStruct, 1) & "]")
		$a_Call = DllCall("Crypt32.dll", "int", "CryptStringToBinary", "str", $base64Data, "int", 0, "int", 1, "ptr", DllStructGetPtr($tsByte), "ptr", DllStructGetPtr($tStruct, 1), "ptr", 0, "ptr", 0)
		If @error Or Not $a_Call[0] Then Return SetError(4, __HttpRequest_ErrNotify('_B64Decode', 'Gọi chức năng CryptStringToBinary từ Crypt32.dll thất bại #2'), $base64Data)
		Local $binaryData = DllStructGetData($tsByte, 1)
	EndIf
	If $iUnCompressData Then $binaryData = __LZNT_Decompress($binaryData)
	Return $binaryData
EndFunc


Func _B64SetupDatabase($___sChr64, $___sPadding = '=')
	If StringInStr($___sChr64, $___sPadding, 1, 1) Then Return SetError(1, __HttpRequest_ErrNotify('_B64SetupDatabase', 'Tham số $___sChr64 không được bao gồm dấu ='), False)
	Local $___aChr64 = StringSplit($___sChr64, "", 2)
	Local $___iCounter = 0, $___uBound = UBound($___aChr64) - 1
	If $___uBound <> 63 Then Return SetError(2, __HttpRequest_ErrNotify('_B64SetupDatabase', 'Tham số $___sChr64 phải là chuỗi dài 64 ký tự'), False)
	For $i = 0 To $___uBound
		For $k = 0 To $___uBound
			If $___aChr64[$i] == $___aChr64[$k] Then $___iCounter += 1
		Next
		If $___iCounter = 2 Then Return SetError(3, __HttpRequest_ErrNotify('_B64SetupDatabase', 'Cài đặt Database thất bại'), False)
		$___iCounter = 0
	Next
	$g___sChr64 = $___sChr64
	$g___aChr64 = $___aChr64
	$g___sPadding = $___sPadding
	Return True
EndFunc



#Region Crypt
	Func __LZNT_Decompress($bBinary)
		$bBinary = Binary($bBinary)
		Local $tInput = DllStructCreate("byte[" & BinaryLen($bBinary) & "]")
		DllStructSetData($tInput, 1, $bBinary)
		Local $tBuffer = DllStructCreate("byte[" & 16 * DllStructGetSize($tInput) & "]")
		Local $a_Call = DllCall("ntdll.dll", "int", "RtlDecompressBuffer", "ushort", 2, "ptr", DllStructGetPtr($tBuffer), "dword", DllStructGetSize($tBuffer), "ptr", DllStructGetPtr($tInput), "dword", DllStructGetSize($tInput), "dword*", 0)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('__LZNT_Decompress', ' Decompress Buffer thất bại'), '')
		Local $tOutput = DllStructCreate("byte[" & $a_Call[6] & "]", DllStructGetPtr($tBuffer))
		Return SetError(0, 0, DllStructGetData($tOutput, 1))
	EndFunc

	Func __LZNT_Compress($bBinary)
		$bBinary = Binary($bBinary)
		Local $tInput = DllStructCreate("byte[" & BinaryLen($bBinary) & "]")
		DllStructSetData($tInput, 1, $bBinary)
		Local $a_Call = DllCall("ntdll.dll", "int", "RtlGetCompressionWorkSpaceSize", "ushort", 2, "dword*", 0, "dword*", 0)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('__LZNT_Compress', 'Tạo WorkSpace thất bại'), "")
		Local $tWorkSpace = DllStructCreate("byte[" & $a_Call[2] & "]")
		Local $tBuffer = DllStructCreate("byte[" & 16 * DllStructGetSize($tInput) & "]")
		Local $a_Call = DllCall("ntdll.dll", "int", "RtlCompressBuffer", "ushort", 2, "ptr", DllStructGetPtr($tInput), "dword", DllStructGetSize($tInput), "ptr", DllStructGetPtr($tBuffer), "dword", DllStructGetSize($tBuffer), "dword", 4096, "dword*", 0, "ptr", DllStructGetPtr($tWorkSpace))
		If @error Then Return SetError(2, __HttpRequest_ErrNotify('__LZNT_Compress', 'Compress Buffer thất bại'), '')
		Local $tOutput = DllStructCreate("byte[" & $a_Call[7] & "]", DllStructGetPtr($tBuffer))
		Return SetError(0, 0, DllStructGetData($tOutput, 1))
	EndFunc

	Func _GetMD5($sFilePath_or_Data)
		Return _GetHash($sFilePath_or_Data, 0x00008003)
	EndFunc

	Func _GetMD5Decrypt($sMD5Encrypt, $iSuperMode = False)
		If $iSuperMode Then
			Local $_a = StringRegExp(_HttpRequest(2, 'https://www.md5online.org/'), '(?i)name="a" value="(.*?)"', 1)
			If @error Then SetError(1, __HttpRequest_ErrNotify('_GetMD5Decrypt', 'Mở trang md5online.org thất bại'), $sMD5Encrypt)
			Local $g_captcha = _IE_RecaptchaBox('https://www.md5online.org/')
			If @error Then Return SetError(2, __HttpRequest_ErrNotify('_GetMD5Decrypt', 'Chạy ReCaptcha thất bại'), $sMD5Encrypt)
			Local $aDecrypt = StringRegExp(_HttpRequest(2, _
					'https://www.md5online.org/', _
					'md5=' & $sMD5Encrypt & '&g-recaptcha-response=' & $g_captcha & '&action=decrypt&a=' & $_a), _
					'(?i)<span class="result".*?>Found :\h*?<b>(.*?)</b></span>', 1)
			If @error Then Return SetError(3, __HttpRequest_ErrNotify('_GetMD5Decrypt', 'Không tìm thấy chuỗi MD5 này trên Database'), $sMD5Encrypt)
			Return $aDecrypt[0]
		Else
			Local $vTimer = RegRead($g___HttpRequestReg & 'MD5Decrypt', 'Timer')
			Local $vCounter = RegRead($g___HttpRequestReg & 'MD5Decrypt', 'Counter')
			;------------------------------------------------------------------------------------------------------------------------
			If Not $vTimer Or TimerDiff($vTimer) > 3600 * 1000 Then
				$vTimer = TimerInit()
				$vCounter = 0
				RegWrite($g___HttpRequestReg & 'MD5Decrypt', 'Timer', 'REG_SZ', $vTimer)
				RegWrite($g___HttpRequestReg & 'MD5Decrypt', 'Counter', 'REG_SZ', $vCounter)
			Else
				RegWrite($g___HttpRequestReg & 'MD5Decrypt', 'Counter', 'REG_SZ', $vCounter + 1)
			EndIf
			;------------------------------------------------------------------------------------------------------------------------
			If $vCounter > 15 Then Return SetError(4, __HttpRequest_ErrNotify('_GetMD5Decrypt', 'Chỉ có thể gửi request giải MD5 15 lần trong 1 giờ'), $sMD5Encrypt)
			;------------------------------------------------------------------------------------------------------------------------
			Local $aDecrypt = StringRegExp(_HttpRequest(2, 'http://md5.my-addr.com/md5_decrypt-md5_cracker_online/md5_decoder_tool.php', 'md5=' & $sMD5Encrypt), '(?i)Hashed string</span>: (.*?)</div>', 1)
			If @error Then Return SetError(5, __HttpRequest_ErrNotify('_GetMD5Decrypt', 'Không tìm thấy chuỗi MD5 này trên Database'), $sMD5Encrypt)
			Return $aDecrypt[0]
		EndIf
	EndFunc

	Func _GetSHA256($sFilePath_or_Data)
		Return _GetHash($sFilePath_or_Data, 0x0000800c)
	EndFunc

	Func _GetSHA1($sFilePath_or_Data)
		Return _GetHash($sFilePath_or_Data, 0x00008004)
	EndFunc

	Func _GetHash($sFilePath_or_Data, $iAlgID)
		If StringRegExp($sFilePath_or_Data, '(?i)^[A-Z]:\') And FileExists($sFilePath_or_Data) Then
			Return StringLower(Hex(_Crypt_HashFile($sFilePath_or_Data, $iAlgID)))
		Else
			$sFilePath_or_Data = StringToBinary($sFilePath_or_Data, 4)
			Return StringLower(Hex(_Crypt_HashData($sFilePath_or_Data, $iAlgID)))
		EndIf
	EndFunc

	Func _GetHMAC_Ex($bData, $bKey, $sAlgorithm = 'SHA256', $bRaw_Output = False) ;$sAlgorithm = SHA512, SHA256, SHA1, SHA384, MD5, RIPEMD160  - Author: DannyFire
		Local $oHashHMACErrorHandler = ObjEvent("AutoIt.Error", "_HashHMACErrorHandler")
		Local $oHMAC = ObjCreate("System.Security.Cryptography.HMAC" & $sAlgorithm)
		If @error Then SetError(1, 0, "")
		$oHMAC.key = Binary($bKey)
		Local $bHash = $oHMAC.ComputeHash_2(Binary($bData))
		Return SetError(0, 0, $bRaw_Output ? $bHash : StringLower(StringMid($bHash, 3)))
	EndFunc

	Func _GetHMAC($sString, $iKey, $iAlgID = 0x0000800c) ;$CALG_SHA_256
		_Crypt_Startup()
		Local $iBlockSize = 64
		Local $a_oPadding[$iBlockSize], $a_iPadding[$iBlockSize]
		Local $oPadding = Binary(''), $iPadding = Binary('')
		$iKey = Binary($iKey)
		If BinaryLen($iKey) > $iBlockSize Then
			$iKey = _Crypt_HashData($iKey, $iAlgID)
			If @error Then Return SetError(1, __HttpRequest_ErrNotify('_GetHMAC', '_Crypt_HashData thất bại #1'), -1)
		EndIf
		For $i = 1 To BinaryLen($iKey)
			$a_iPadding[$i - 1] = Number(BinaryMid($iKey, $i, 1))
			$a_oPadding[$i - 1] = Number(BinaryMid($iKey, $i, 1))
		Next
		For $i = 0 To $iBlockSize - 1
			$a_oPadding[$i] = BitXOR($a_oPadding[$i], 0x5C)
			$a_iPadding[$i] = BitXOR($a_iPadding[$i], 0x36)
		Next
		For $i = 0 To $iBlockSize - 1
			$iPadding &= Binary('0x' & Hex($a_iPadding[$i], 2))
			$oPadding &= Binary('0x' & Hex($a_oPadding[$i], 2))
		Next
		Local $HashS1 = _Crypt_HashData($iPadding & Binary($sString), $iAlgID)
		If @error Then Return SetError(2, __HttpRequest_ErrNotify('_GetHMAC', '_Crypt_HashData thất bại #2'), -1)
		Local $HashS2 = _Crypt_HashData($oPadding & $HashS1, $iAlgID)
		If @error Then Return SetError(3, __HttpRequest_ErrNotify('_GetHMAC', '_Crypt_HashData thất bại #3'), -1)
		_Crypt_Shutdown()
		Return StringLower(Hex($HashS2))
	EndFunc

;~ 	Func _hmac_sha256($sString, $iKey, $iAlgID = 0x0000800c, $iBlockSize = 64) ;Chưa test
;~ 		Local Const $oConst = 0x5C, $iConst = 0x36
;~ 		Local $a_opad[$iBlockSize], $a_ipad[$iBlockSize]
;~ 		Local $opad = Binary(''), $ipad = Binary('')
;~ 		$iKey = Binary($iKey)
;~ 		If BinaryLen($iKey) > $iBlockSize Then $iKey = _Crypt_HashData($iKey, $iAlgID)
;~ 		For $i = 1 To BinaryLen($iKey)
;~ 			$a_ipad[$i - 1] = Number(BinaryMid($iKey, $i, 1))
;~ 			$a_opad[$i - 1] = Number(BinaryMid($iKey, $i, 1))
;~ 		Next
;~ 		For $i = 0 To $iBlockSize - 1
;~ 			$a_opad[$i] = BitXOR($a_opad[$i], $oConst)
;~ 			$a_ipad[$i] = BitXOR($a_ipad[$i], $iConst)
;~ 		Next
;~ 		For $i = 0 To $iBlockSize - 1
;~ 			$ipad &= Binary(Hex($a_ipad[$i], 2))
;~ 			$opad &= Binary(Hex($a_opad[$i], 2))
;~ 		Next
;~ 		Return StringRegExpReplace(_Crypt_HashData($opad & _Crypt_HashData($ipad & Binary($sString), $iAlgID), $iAlgID), "0x", "")
;~ 	EndFunc
#EndRegion



#Region <FUNC đã đổi tên và sẽ bị loại bỏ ở phiên bản sau>
	#cs
		« - - - - - - - - - - -Huân Hoàng - - - - - - - - -»
		« - - - - - - - - - - -Rainy Pham - - - - - - - - -»
	#ce

	Func _HttpRequest_SetSession($sSessionNumber)
		_HttpRequest_SessionSet($sSessionNumber)
	EndFunc

	Func _HttpRequest_ClearSession($sSessionNumber = 0, $vClearProxy = False)
		_HttpRequest_SessionClear($sSessionNumber, $vClearProxy)
	EndFunc

	Func _TimeStampUNIX($iMSec = @MSEC, $iSec = @SEC, $iMin = @MIN, $iHour = @HOUR, $iDay = @MDAY, $iMonth = @MON, $iYear = @YEAR)
		__ConsoleOldFuncWarning('_TimeStampUNIX', '_GetTimeStamp')
	EndFunc

	Func _URLDecode($iParam1 = '', $iParam2 = '', $iParam3 = '', $iParam4 = '', $iParam5 = '', $iParam6 = '')
		__ConsoleOldFuncWarning('_URLDecode', '_HTMLDecode')
	EndFunc

	Func _WinHttpBoundaryGenerator()
		__ConsoleOldFuncWarning('_WinHttpBoundaryGenerator', '_BoundaryGenerator')
	EndFunc

	Func _HttpRequest_CreateDataFormSimple($a_FormItems)
		__ConsoleOldFuncWarning('_HttpRequest_CreateDataFormSimple', '_HttpRequest_CreateDataForm')
	EndFunc

	Func _HttpRequest_ClearCookies($sSessionNumber = 0)
		__ConsoleOldFuncWarning('_HttpRequest_ClearCookies', '_HttpRequest_ClearSession.')
	EndFunc

	Func _HttpRequest_NewSession($sSessionNumber = 0)
		__ConsoleOldFuncWarning('_HttpRequest_NewSession', '_HttpRequest_ClearSession.')
	EndFunc

	Func _GetFileInfos($sFilePath, $vDataTypeReturn = 1)
		__ConsoleOldFuncWarning('_GetFileInfos', '_GetFileInfo')
	EndFunc

	Func _GetLocation_Redirect($__sHeader = '', $iIndex = -1)
		__ConsoleOldFuncWarning('_GetLocation_Redirect', '_GetLocationRedirect')
	EndFunc

	Func _FileWrite_Test($sData, $FilePath = Default, $iMode = 0)
		__ConsoleOldFuncWarning('_FileWrite_Test', '_HttpRequest_Test')
	EndFunc

	Func _GetHiddenValues($iSourceHtml_or_URL, $iKeySearch = '', $iReturnArray = False, $iInputType = 0)
		__ConsoleOldFuncWarning('_GetHiddenValues', '_HttpRequest_SearchHiddenValues')
	EndFunc

	Func _TimeStamp2Date($iTimeStamp, $vLocalTime = False)
		__ConsoleOldFuncWarning('_TimeStamp2Date', '_GetDateFromTimeStamp')
	EndFunc

	Func _HttpRequest_GetImageBinaryDimension($sBinaryData_Or_FilePath, $Release_hBitmap = True, $isFilePath = False)
		__ConsoleOldFuncWarning('_HttpRequest_GetImageBinaryDimension', '_Image_GetDimension', False)
		Local $vValue = _Image_GetDimension($sBinaryData_Or_FilePath, $Release_hBitmap, $isFilePath)
		Return SetError(@error, @error, $vValue)
	EndFunc

	Func _HttpRequest_SetImageBinaryToGUI($sBinaryData_Or_FilePath_Or_URL_Or_arrayHBitmap, $idCtrl_Or_hWnd, $width_Image = Default, $height_Image = Default)
		__ConsoleOldFuncWarning('_HttpRequest_SetImageBinaryToGUI', '_Image_SetGUI', False)
		Local $vValue = _Image_SetGUI($sBinaryData_Or_FilePath_Or_URL_Or_arrayHBitmap, $idCtrl_Or_hWnd, $width_Image, $height_Image)
		Return SetError(@error, @error, $vValue)
	EndFunc

	Func _HttpRequest_SimpleCaptchaGUI($BinaryCaptcha, $___x = -1, $___y = -1, $___hParent = Default)
		__ConsoleOldFuncWarning('_HttpRequest_SimpleCaptchaGUI', '_Image_SetSimpleCaptchaGUI', False)
		Local $vValue = _Image_SetSimpleCaptchaGUI($BinaryCaptcha, $___x, $___y, $___hParent)
		Return SetError(@error, @error, $vValue)
	EndFunc


	Func __ConsoleOldFuncWarning($oldName, $newName, $vExit = True)
		MsgBox(4096, 'Lưu ý', 'Hàm "' & $oldName & '" đã đổi tên thành "' & $newName & '". Vui lòng sử dụng tên hàm mới bởi ' & $oldName & ' sẽ bị loại bỏ ở các phiên bản sau.')
		If $vExit Then Exit
	EndFunc

	Func __SciTE_TextSplit($nameVar, $nCharPerLine = 101)
		Local $sStr = StringReplace(ClipGet(), "'", "''", 0, 1), $sRet = ''
		Do
			$sRet &= $nameVar & " &= '" & StringLeft($sStr, $nCharPerLine) & "'" & @CRLF
			$sStr = StringTrimLeft($sStr, $nCharPerLine)
		Until StringLen($sStr) = 0
		ClipPut($sRet)
	EndFunc

	Func __SciTE_RunOnDetach()
		If Not @Compiled And $CmdLine[0] = 0 Then
			_HttpRequest_ConsoleWrite('> __SciTE_RunOnDetach được khởi tạo : SciTe sẽ không chờ cho đến khi code chạy xong.' & @CRLF)
			Exit Run(FileGetShortName(@AutoItExe) & ' "' & @ScriptFullPath & '" --detach-scite', @WorkingDir, @SW_HIDE)
		EndIf
	EndFunc

	Func __SciTE_ConsoleWrite_FixFont()
		If @Compiled Or ($CmdLine[0] > 0 And $CmdLine[1] = '--hh-multi-process') Then Return
		;----------------------------------------------------------------------------------------------------------------------------
		Local $SciTE_Link_A = StringRegExpReplace(@AutoItExe, '(?i)\w+\.exe$', '') & 'SciTE\'
		Local $SciTE_Link_B = StringRegExpReplace(__WinAPI_GetProcessFileName(ProcessExists('SciTE.exe')), '(?i)\w+\.exe$', '')
		If $SciTE_Link_B = '' Then $SciTE_Link_B = $SciTE_Link_A
		If $SciTE_Link_A = $SciTE_Link_B Then
			Local $SciTEProp_Link = [@LocalAppDataDir & '\AutoIt v3\SciTE\SciTEUser.properties', $SciTE_Link_A & 'SciTEUser.properties', $SciTE_Link_A & 'SciTEGlobal.properties']
		Else
			Local $SciTEProp_Link = [@LocalAppDataDir & '\AutoIt v3\SciTE\SciTEUser.properties', $SciTE_Link_A & 'SciTEUser.properties', $SciTE_Link_A & 'SciTEGlobal.properties', $SciTE_Link_B & 'SciTEUser.properties', $SciTE_Link_B & 'SciTEGlobal.properties']
		EndIf
		For $i = 0 To UBound($SciTEProp_Link) - 1
			Local $SciTEUserProp_Change = 0
			If FileExists($SciTEProp_Link[$i]) Then
				Local $SciTEUserProp_Data = FileRead($SciTEProp_Link[$i])
				If Not StringRegExp($SciTEUserProp_Data, '(?im)^\h*?\Qoutput.code.page\E\h*?=\h*?65001') Or StringRegExp($SciTEUserProp_Data, '(?im)^\h*?\Qoutput.code.page\E\h*?=\h*?0') Then
					$SciTEUserProp_Data = StringRegExpReplace($SciTEUserProp_Data, '(?im)^\h*?\Qoutput.code.page\E.*$\R', '')
					$SciTEUserProp_Data &= @CRLF & 'output.code.page=65001'
					$SciTEUserProp_Change = 1
				EndIf
				If Not StringRegExp($SciTEUserProp_Data, '(?im)^\h*?\Qcode.page\E\h*?=\h*?65001') Or StringRegExp($SciTEUserProp_Data, '(?im)^\h*?\Qoutput.code.page\E\h*?=\h*?0') Then
					$SciTEUserProp_Data = StringRegExpReplace($SciTEUserProp_Data, '(?im)^\h*?\Qcode.page\E.*$\R', '')
					$SciTEUserProp_Data &= @CRLF & 'code.page=65001'
					$SciTEUserProp_Change = 1
				EndIf
			Else
				$SciTEUserProp_Change = 1
				$SciTEUserProp_Data = 'output.code.page=65001' & @CRLF & 'code.page=65001' & @CRLF
			EndIf
			If $SciTEUserProp_Change = 1 Then
				Local $hOpen = FileOpen($SciTEProp_Link[$i], 2 + 8)
				FileWrite($hOpen, $SciTEUserProp_Data)
				FileClose($hOpen)
			EndIf
		Next
		;----------------------------------------------------------------------------------------------------------------------------
		If $g___ConsoleForceUTF8 = True Then
			Local $SciTEProp_Link = @ScriptDir & '\SciTE.properties'
			If Not FileExists($SciTEProp_Link) Then
				Local $hOpen = FileOpen($SciTEProp_Link, 2 + 8)
				FileWrite($hOpen, 'output.code.page=65001' & @CRLF & 'code.page=65001' & @CRLF)
				FileClose($hOpen)
			EndIf
		EndIf
	EndFunc

	Func __SciTE_ConsoleClear()
		__SciTE_Command("menucommand:420")
	EndFunc

	Func __SciTE_Command($sCmd)
		If @Compiled Then Return
		Local $CmdStruct = DllStructCreate('Char[' & StringLen($sCmd) + 1 & ']')
		DllStructSetData($CmdStruct, 1, $sCmd)
		Local $COPYDATA = DllStructCreate('Ptr;DWord;Ptr')
		DllStructSetData($COPYDATA, 1, 1)
		DllStructSetData($COPYDATA, 2, StringLen($sCmd) + 1)
		DllStructSetData($COPYDATA, 3, DllStructGetPtr($CmdStruct))
		DllCall($dll_User32, 'None', 'SendMessage', 'HWnd', WinGetHandle("DirectorExtension"), 'Int', 74, 'HWnd', 0, 'Ptr', DllStructGetPtr($COPYDATA))
	EndFunc

	Func __WinAPI_GetProcessFileName($iPID)
		If $iPID = 0 Then Return SetError(1, __HttpRequest_ErrNotify('__WinAPI_GetProcessFileName', 'PID không tồn tại (PID = 0)'), '')
		;-----------------------------------------------------
		Local $__tOSVI__ = DllStructCreate('struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct')
		DllStructSetData($__tOSVI__, 1, DllStructGetSize($__tOSVI__))
		Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $__tOSVI__)
		If @error Or Not $aRet[0] Then Return SetError(2, 0, '')
		Local $__WINVER__ = BitOR(BitShift(DllStructGetData($__tOSVI__, 2), -8), DllStructGetData($__tOSVI__, 3))
		;-----------------------------------------------------
		Local $hProcess = DllCall('kernel32.dll', 'handle', 'OpenProcess', 'dword', $__WINVER__ < 0x0600 ? 0x00000410 : 0x00001010, 'bool', 0, 'dword', $iPID)
		If @error Or Not $hProcess[0] Then Return SetError(3, 0, '')
		;-----------------------------------------------------
		Local $aFileNameExW = DllCall('psapi.dll', 'dword', 'GetModuleFileNameExW', 'handle', $hProcess[0], 'handle', 0, 'wstr', '', 'int', 4096)
		If @error Or Not $aFileNameExW[0] Then Return SetError(4, 0, '')
		;-----------------------------------------------------
		DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess[0])
		;-----------------------------------------------------
		Return $aFileNameExW[3]
	EndFunc

	Func __RemoveVietMarktical($sText)
		If $g___aVietPattern = '' Then Global $g___aVietPattern = [['áàảãạăắằẳẵặâấầẩẫậ', 'a'], ['đ', 'd'], ['éèẻẽẹêếềểễệ', 'e'], ['íìỉĩị', 'i'], ['óòỏõọôốồổỗộơớờởỡợ', 'o'], ['úùủũụưứừửữự', 'u'], ['ýỳỷỹỵ', 'y']]
		For $i = 0 To 6
			$sText = StringRegExpReplace($sText, '[' & $g___aVietPattern[$i][0] & ']', $g___aVietPattern[$i][1])
			$sText = StringRegExpReplace($sText, '[' & StringUpper($g___aVietPattern[$i][0]) & ']', StringUpper($g___aVietPattern[$i][1]))
		Next
		Return $sText
	EndFunc
#EndRegion





#Region <Google Request>
	Func _HttpRequest_GoogleLogin($sUser, $sPass, $sRedirectURL = Default, $iReturn = Default, $iUserAgent = Default, $iSetPrevUAAfterLogin = True, $___x = -1, $___y = -1, $___hParent = Default, $vUselessParam = 0)
		If $iReturn = Default Then $iReturn = 2
		If $sRedirectURL = Default Then $sRedirectURL = ''
		If $iUserAgent = Default Then $iUserAgent = $g___defUserAgent
		$sUser = _URIEncode(StringRegExpReplace($sUser, '(?i)@gmail.com[\.\w]*?$', '', 1))
		$sPass = _URIEncode($sPass)
		Local $BkUserAgent = _HttpRequest_SetUserAgent($iUserAgent)
		Local $sHeader = ''
		;---------------------------------------------------------------------------------------
		Local $rq1 = _HttpRequest(2, 'https://accounts.google.com/signin/v1/lookup', 'Email=' & $sUser)
		If StringInStr($rq1, '"errormsg_0_Email"', 0, 1) Then
			Return SetError(1 * _HttpRequest_SetUserAgent($BkUserAgent), __HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Google không nhận dạng được email này'), $rq1)
		EndIf
		$sHeader &= $g___retData[$g___LastSession][0] & @CRLF & @CRLF
		$aHiddenValue = StringRegExp($rq1, '(?i)name="(gxf|GALX|ProfileInformation|SessionState)".*?value="(.*?)"', 3)
		If @error Or UBound($aHiddenValue) <> 8 Then
			Return SetError(2 * _HttpRequest_SetUserAgent($BkUserAgent), __HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Không tìm thấy tham số đăng nhập tài khoản từ Html #2'), $rq1)
		EndIf
		;---------------------------------------------------------------------------------------
		Local $rq2 = _HttpRequest(2, 'https://accounts.google.com/signin/challenge/sl/password', $aHiddenValue[0] & '=' & _URIEncode($aHiddenValue[1]) & '&' & $aHiddenValue[2] & '=' & _URIEncode($aHiddenValue[3]) & '&' & $aHiddenValue[4] & '=' & _URIEncode($aHiddenValue[5]) & '&' & $aHiddenValue[6] & '=' & _URIEncode($aHiddenValue[7]) & '&Email=' & $sUser & '&Passwd=' & $sPass & '&signIn=Sign+in&PersistentCookie=yes&Page=PasswordSeparationSignIn&flowName=GlifWebSignIn&_utf8=%E2%98%83&bgresponse=&continue=' & _URIEncode($sRedirectURL))
		Local $CaptchaCheck = StringInStr($rq2, '<div class="captcha-container">', 1, 1)
		If StringInStr($rq2, '"errormsg_0_Passwd"', 0, 1) And Not $CaptchaCheck Then
			Return SetError(3 * _HttpRequest_SetUserAgent($BkUserAgent), __HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Mật khẩu không chính xác #1'), $rq2)
		EndIf
		If $CaptchaCheck Then
			For $i = 0 To 2
				$aHiddenValue = StringRegExp($rq2, '(?i)name="(gxf|ProfileInformation|SessionState)".*?value="(.*?)"', 3)
				If @error Or UBound($aHiddenValue) <> 6 Then
					Return SetError(4 * _HttpRequest_SetUserAgent($BkUserAgent), __HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Không tìm thấy tham số đăng nhập tài khoản từ Html #3'), $rq2)
				EndIf
				Local $TokenLogin = StringRegExp($rq2, '(?i)name="logintoken".*?value="(.*?)"', 1)
				If @error Then Return SetError(5 * _HttpRequest_SetUserAgent($BkUserAgent), __HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Không tìm thấy tham số đăng nhập tài khoản từ Html #4'), $rq2)
				Local $CaptchaLink = 'https://accounts.google.com/Captcha?v=2&ctoken=' & $TokenLogin[0]
				TrayTip('Google Captcha', 'Nhập Captcha để tiếp tục', 30, 2)
				Local $CaptchaValue = _Image_SetSimpleCaptchaGUI(_HttpRequest(-2, $CaptchaLink), $___x, $___y, $___hParent)
				TrayTip('', '', 0)
				$rq2 = _HttpRequest(2, 'https://accounts.google.com/signin/challenge/sl/password', 'Page=PasswordSeparationSignIn&continue=' & _URIEncode($sRedirectURL) & '&flowName=GlifWebSignIn&_utf8=%E2%98%83&bgresponse=&Email=' & $sUser & '&Passwd=' & $sPass & '&' & $aHiddenValue[0] & '=' & _URIEncode($aHiddenValue[1]) & '&' & $aHiddenValue[2] & '=' & _URIEncode($aHiddenValue[3]) & '&' & $aHiddenValue[4] & '=' & _URIEncode($aHiddenValue[5]) & '&logintoken=' & $TokenLogin[0] & '&url=' & _URIEncode($CaptchaLink) & '&logintoken_audio=' & $TokenLogin[0] & '&url_audio=' & _URIEncode($CaptchaLink & '&kind=audio') & '&logincaptcha=' & $CaptchaValue & '&signIn=Sign+in&PersistentCookie=yes')
				If Not StringInStr($rq2, '<div class="captcha-container">', 1, 1) Then ExitLoop
				MsgBox(4096, 'Lỗi', 'Bạn đã nhập sai Captcha. Vui lòng thử lại.')
			Next
			If $i = 3 Then Exit MsgBox(4096, 'Lỗi', 'Bạn đã nhập sai Captcha liên tiếp 3 lần. Code sẽ tắt để đảm bảo tài khoản không bị khoá.')
		EndIf
		If StringInStr($rq2, '"errormsg_0_Passwd"', 0, 1) Then Return SetError(3 * _HttpRequest_SetUserAgent($BkUserAgent), __HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Mật khẩu không chính xác #2'), $rq2)
		$sHeader &= $g___retData[$g___LastSession][0] & @CRLF & @CRLF
		;---------------------------------------------------------------------------------------
		If StringInStr($rq2, 'data-phone-step-skip-link', 1, 1) Then
			_HttpRequest_ConsoleWrite('<Account bị kiểm tra Recovery phone number Or Recovery email>' & @CRLF)
			If $vUselessParam = 0 Then
				Return _HttpRequest_GoogleLogin($sUser, $sPass, $sRedirectURL, $iReturn, $iSetPrevUAAfterLogin, $___x, $___y, $___hParent, 1)
			Else
				Return SetError(6 * _HttpRequest_SetUserAgent($BkUserAgent), __HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Tài khoản yêu cầu cập nhật thông tin. Hãy đăng nhập trên trình duyệt kiểm tra lại'), $rq2)
			EndIf

		ElseIf StringInStr($rq2, 'https://accounts.google.com/signin/newfeatures/options', 1, 1) Then
			_HttpRequest_ConsoleWrite('<Account bị kiểm tra New Features>' & @CRLF)
			Local $aParam = StringRegExp(StringReplace($rq2, '&amp;', '&', 0, 1), '(?i)<input type="hidden" name="(.*?)" value="(.*?)"', 3)
			Local $sParam = ''
			For $i = 0 To UBound($aParam) - 1 Step 2
				$sParam &= $aParam[$i] & '=' & _URIEncode($aParam[$i + 1]) & '&'
			Next
			_HttpRequest(0, 'https://accounts.google.com/signin/newfeatures/save', StringTrimRight($sParam, 1))
			If $vUselessParam < 2 Then
				Return _HttpRequest_GoogleLogin($sUser, $sPass, $sRedirectURL, $iReturn, $iSetPrevUAAfterLogin, $___x, $___y, $___hParent, 2)
			Else
				Return SetError(5 * _HttpRequest_SetUserAgent($BkUserAgent), __HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Tài khoản yêu cầu cập nhật thông tin. Hãy đăng nhập trên trình duyệt kiểm tra lại'), $rq2)
			EndIf

		ElseIf StringInStr($rq2, 'action="/signin/challenge/az', 1, 1) Then
			Exit MsgBox(4096 + 48, 'Thông báo', 'Tài khoản của bạn bị bắt buộc phải verify bằng Điện thoại.' & @CRLF & 'Vui lòng cập nhật lại tài khoản trên trình duyệt rồi thử lại.')
		EndIf
		;---------------------------------------------------------------------------------------
		If StringRegExp($rq2, 'action="\/signin\/challenge\/kpp\/[45]"') Then
			_HttpRequest_SetUserAgent($BkUserAgent)
			Return SetError(7, __HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Tài khoản cần được xác thực - Vui lòng mở Gmail trên trình duyệt và báo cáo an toàn nếu nhận được thông báo Activity'), $rq2)
		ElseIf Not StringInStr($sHeader, 'SAPISID', 0, 1) Then
			_HttpRequest_SetUserAgent($BkUserAgent)
			Return SetError(8, __HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Đăng nhập thất bại không rõ nguyên do. Vui lòng LogOut (nếu đã đăng nhập) và LogIn lại tài khoản trên trình duyệt'), $rq2)
		EndIf
		;---------------------------------------------------------------------------------------
		If $iSetPrevUAAfterLogin Then _HttpRequest_SetUserAgent($BkUserAgent)
		Switch $iReturn
			Case -1
				Return _GetCookie($sHeader)
			Case 0
				Return ''
			Case 1
				Return $sHeader
			Case 2
				Return $rq2
			Case 4
				Local $aRet = [$sHeader, $rq2]
				Return $aRet
			Case Else
				__HttpRequest_ErrNotify('_HttpRequest_GoogleLogin', 'Chỉ chấp nhận $iReturn = -1 hoặc 0 hoặc 1 hoặc 2 hoặc 4')
				Return $rq2
		EndSwitch
	EndFunc

	Func _HttpRequest_Google_SAPISIDHASH($SAPISID, $xOrigin) ;https://stackoverflow.com/questions/16907352/reverse-engineering-javascript-behind-google-button
		Local $sTimeStamp = _GetTimeStamp()
		Return 'SAPISIDHASH ' & $sTimeStamp & '_' & _GetSHA1($sTimeStamp & ' ' & $SAPISID & ' ' & $xOrigin)
	EndFunc

	Func _HttpRequest_Google_CheckNewDevice()
		Local $sHTML = _HttpRequest(2, 'https://accounts.google.com/b/0/DisplayUnlockCaptcha')
		Local $aHiddenValue = StringRegExp($sHTML, '(?i)id="(timeStmp|secTok)"[\s\S]*?value=[''"](.+?)[''"]', 3)
		If Not @error And UBound($aHiddenValue) = 4 Then _HttpRequest(0, 'https://accounts.google.com/b/0/DisplayUnlockCaptcha', $aHiddenValue[0] & '=' & _URIEncode($aHiddenValue[1]) & '&' & $aHiddenValue[2] & '=' & _URIEncode($aHiddenValue[3]) & '&submitChallenge=Continue')
		$sHTML = _HttpRequest(2, 'https://myaccount.google.com/security-checkup?continue=https://myaccount.google.com/')
		Local $aDeviceID = StringRegExp($sHTML, '(?i)data-event-id=("-?\d+")', 3)
		If @error Then Return SetError(1, '', False)
		$aDeviceID = __ArrayDuplicate($aDeviceID)
		For $i = 0 To UBound($aDeviceID) - 1
			__Google_SettingsOnOff($sHTML, 161362964, $aDeviceID[$i], 2)
			If @error > 0 And @error < 3 Then Return SetError(2, '', False)
		Next
		Return True
	EndFunc

	Func _HttpRequest_Google_AllowLessSecureApps($iState)
		Local $sHTML = _HttpRequest(2, 'https://myaccount.google.com/security')
		__Google_SettingsOnOff($sHTML, 139777153, $iState)
		$sHTML = _HttpRequest(2, 'https://myaccount.google.com/security-checkup?continue=https://myaccount.google.com/')
		Local $aEventID = StringRegExp($sHTML, '(?i)true,("-?\d+"),\[4,1,5\]', 3)
		If @error Then Return SetError(1, '', False)
		For $i = 0 To UBound($aEventID) - 1
			__Google_SettingsOnOff($sHTML, 161362964, $aEventID[$i], 2)
			If @error > 0 And @error < 3 Then Return SetError(2, '', False)
		Next
		Return True
	EndFunc

	Func __Google_SettingsOnOff($sHTML, $iExtension, $iState, $iAddtionalData = '')
		Local $at = StringRegExp($sHTML, "(?i)\Q'https:\/\/www.google.com\/settings',\E'(.*?)'", 1)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('__Google_SettingsOnOff', 'Chưa đăng nhập Google hoặc Đăng nhập thất bại'))
		Local $boq = StringRegExp($sHTML, '(?i)"(boq_identity.*?)"', 1)
		If @error Then Return SetError(2, __HttpRequest_ErrNotify('__Google_SettingsOnOff', 'Chưa đăng nhập Google hoặc Đăng nhập thất bại'))
		If Not IsNumber($iAddtionalData) Then $iAddtionalData = '"' & StringReplace($iAddtionalData, '\u003d', '=') & '"'
		_HttpRequest(0, 'https://myaccount.google.com/_/AccountSettingsUi/mutate?ds.extension=' & $iExtension & '&f.sid=&bl=' & $boq[0] & '&hl=en&_reqid=&rt=c', 'f.req=' & _URIEncode('["af.maf",[["af.add",' & $iExtension & ',[{"' & $iExtension & '":[' & $iState & ($iAddtionalData ? ',' & $iAddtionalData : '') & ']}]]]]') & '&at=' & _URIEncode($at[0]) & '&')
		If @error Then Return SetError(3)
	EndFunc
#EndRegion




#Region <UDF WinHttp by Trancexx, ProAndy>
	Func _WinHttpGetResponseErrorCode2($iErrorCode)
		$iErrorCode = StringRegExp('OUT_OF_HANDLES12001,TIMEOUT12002,UNKNOWN12003,INTERNAL_ERROR12004,INVALID_URL12005,UNRECOGNIZED_SCHEME12006,NAME_NOT_RESOLVED12007,INVALID_OPTION12009,OPTION_NOT_SETTABLE12011,SHUTDOWN12012,LOGIN_FAILURE12015,OPERATION_CANCELLED12017,INCORRECT_HANDLE_TYPE12018,INCORRECT_HANDLE_STATE12019,CANNOT_CONNECT12029,CONNECTION_ERROR12030,RESEND_REQUEST12032,SECURE_CERT_DATE_INVALID12037,SECURE_CERT_CN_INVALID12038,CLIENT_AUTH_CERT_NEEDED12044,SECURE_INVALID_CA12045,SECURE_CERT_REV_FAILED12057,CANNOT_CALL_BEFORE_OPEN12100,CANNOT_CALL_BEFORE_SEND12101,CANNOT_CALL_AFTER_SEND12102,CANNOT_CALL_AFTER_OPEN12103,HEADER_NOT_FOUND12150,INVALID_SERVER_RESPONSE12152,INVALID_HEADER12153,INVALID_QUERY_REQUEST12154,HEADER_ALREADY_EXISTS12155,REDIRECT_FAILED12156,SECURE_CHANNEL_ERROR12157,BAD_AUTO_PROXY_SCRIPT12166,UNABLE_TO_DOWNLOAD_SCRIPT12167,SECURE_INVALID_CERT12169,SECURE_CERT_REVOKED12170,NOT_INITIALIZED12172,SECURE_FAILURE12175,AUTO_PROXY_SERVICE_ERROR12178,SECURE_CERT_WRONG_USAGE12179,AUTODETECTION_FAILED12180,HEADER_COUNT_EXCEEDED12181,HEADER_SIZE_OVERFLOW12182,CHUNKED_ENCODING_HEADER_SIZE_OVERFLOW12183,RESPONSE_DRAIN_OVERFLOW12184,CLIENT_CERT_NO_PRIVATE_KEY12185,CLIENT_CERT_NO_ACCESS_PRIVATE_KEY12186', '(?:^|,)([A-Z_]+)' & $iErrorCode, 1)
		If @error Then Return 'ERROR_WINHTTP_UNKNOWN'
		Return 'ERROR_WINHTTP_' & $iErrorCode[0]
	EndFunc

	Func _WinHttpQueryHeaders2($hRequest, $iInfoLevel = 22, $iIndex = 0, $vBuffer = 8192)
		If $iInfoLevel = 19 Then $vBuffer = 8
		Switch $iInfoLevel
			Case 80
				Local $vCert = _GetCertificateInfo()
				Return SetError(@error, 0, $vCert)
			Case 81
				Local $vDS = _GetNameDNS()
				Return SetError(@error, 0, $vDS)
			Case Else
				Local $aCall = DllCall($dll_WinHttp, "bool", 'WinHttpQueryHeaders', "handle", $hRequest, "dword", $iInfoLevel, 'wstr', '', 'wstr', "", "dword*", $vBuffer, "dword*", $iIndex)
				If @error Or Not $aCall[0] Then
					If $aCall[5] And $vBuffer < $aCall[5] Then
						$aCall = DllCall($dll_WinHttp, "bool", 'WinHttpQueryHeaders', "handle", $hRequest, "dword", $iInfoLevel, 'wstr', '', 'wstr', "", "dword*", $aCall[5], "dword*", $iIndex)
						If @error Or Not $aCall[0] Then Return SetError(2, 0, 0)
						Return $aCall[4]
					Else
						Return SetError(1, 0, 0)
					EndIf
				EndIf
				Return $aCall[4]
		EndSwitch
	EndFunc

	Func _WinHttpAddRequestHeaders2($hRequest, $sHeader, $iModifier = Default)
		;WINHTTP_ADDREQ_FLAG_ADD = 0x20000000
		;WINHTTP_ADDREQ_FLAG_REPLACE = 0x80000000
		;WINHTTP_ADDREQ_FLAG_ADD_IF_NEW = 0x10000000
		;WINHTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA = 0x40000000
		;WINHTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON = 0x01000000
		If $iModifier = Default Then $iModifier = 0x10000000
		DllCall($dll_WinHttp, "bool", 'WinHttpAddRequestHeaders', "handle", $hRequest, 'wstr', $sHeader, "dword", -1, "dword", $iModifier)
	EndFunc

	Func _WinHttpOpen2($iProxy, $iProxyBypass)
		Local $aCall = DllCall($dll_WinHttp, "handle", "WinHttpOpen", 'wstr', '', "dword", $iProxy ? 3 : 1, 'wstr', $iProxy, 'wstr', $iProxyBypass, "dword", 0)
		If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
		Return $aCall[0]
	EndFunc

	Func _WinHttpSendRequest2($hRequest, $sHeaders = '', $sData2Send = '', $iUpload = 0, $CallBackFunc_Progress = '')
		Local $pData2Send = 0, $lData2Send = 0
		If $sData2Send Then
			$lData2Send = BinaryLen($sData2Send)
			If $iUpload = 0 Then
				Local $tData2Send = DllStructCreate('byte[' & $lData2Send & ']')
				DllStructSetData($tData2Send, 1, $sData2Send)
				$pData2Send = DllStructGetPtr($tData2Send)
			EndIf
		EndIf
		Local $aCall = DllCall($dll_WinHttp, "bool", 'WinHttpSendRequest', "handle", $hRequest, 'wstr', $sHeaders, "dword", 0, "ptr", $pData2Send, "dword", $lData2Send, "dword", $lData2Send, "dword_ptr", 0)
		If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
		If $iUpload Then
			_WinHttpWriteData_Ex($hRequest, $sData2Send, $lData2Send, $CallBackFunc_Progress)
			If @error Then Return SetError(@error, 0, 0)
		EndIf
		Return 1
	EndFunc

	Func _WinHttpWriteData_Ex($hRequest, $sData2Send, $lData2Send, $CallBackFunc_Progress = '', $iBytesPerLoop = $g___BytesPerLoop)
		Local $tBuffer, $iDataMid, $iCheckCallbackFunc = 0, $vNowSizeBytes = 0, $vTotalSizeBytes = -1, $aCall, $isBinData2Send = IsBinary($sData2Send)
		If $CallBackFunc_Progress <> '' Then
			$iCheckCallbackFunc = 1
			$vTotalSizeBytes = $lData2Send
			If $vTotalSizeBytes > 2147483647 Then Return SetError(101, __HttpRequest_ErrNotify('_WinHttpWriteData_Ex', 'Tập tin quá lớn', 101), 0)
		EndIf
		;----------------------------------
		Do
			$iDataMid = $g___aReadWriteData[1][$isBinData2Send]($sData2Send, $vNowSizeBytes + 1, $iBytesPerLoop)
			$iDataMidLen = $g___aReadWriteData[2][$isBinData2Send]($iDataMid)
			$tBuffer = DllStructCreate($g___aReadWriteData[0][$isBinData2Send] & "[" & ($iDataMidLen + 1) & "]")
			DllStructSetData($tBuffer, 1, $iDataMid)
			$aCall = DllCall($dll_WinHttp, "bool", 'WinHttpWriteData', "handle", $hRequest, "struct*", $tBuffer, "dword", $iDataMidLen, "dword*", 0)
			If @error Or Not $aCall[0] Then ExitLoop
			$vNowSizeBytes += $iDataMidLen
			$tBuffer = ''
			;--------------------------------------------------------------------------------
			If $g___CancelReadWrite Then
				$g___CancelReadWrite = False
				Return SetError(999, __HttpRequest_ErrNotify('_WinHttpWriteData_Ex', 'Đã huỷ request', 999), 0)
			ElseIf $iCheckCallbackFunc Then
				$CallBackFunc_Progress($vNowSizeBytes, $vTotalSizeBytes)
			EndIf
		Until $aCall[4] < $iBytesPerLoop
		Return 1
	EndFunc

	Func _WinHttpReadData_Ex($hRequest, $CallBackFunc_Progress = '', $iFileSavePath = '', $iEncodingOfFileSave = 0, $iBytesPerLoop = $g___BytesPerLoop)
		Local $vBinaryData = Binary(''), $aCall, $iCheckCallbackFunc = 0, $vNowSizeBytes = 1, $vTotalSizeBytes = -1
		Local $tBuffer = DllStructCreate("byte[" & $iBytesPerLoop & "]")
		;----------------------------------
		If $CallBackFunc_Progress <> '' Then
			$iCheckCallbackFunc = 1
			$vTotalSizeBytes = Number(_WinHttpQueryHeaders2($hRequest, 5)) ;QUERY_CONTENT_LENGTH
			If $vTotalSizeBytes > 2147483647 Then Return SetError(102, __HttpRequest_ErrNotify('_WinHttpReadData_Ex', 'Tập tin quá lớn', -1), 0)
		EndIf
		;----------------------------------
		If $iFileSavePath Then
			If $iEncodingOfFileSave = 0 Then $iEncodingOfFileSave = 16
			If FileExists($iFileSavePath) Then
				FileOpen($iFileSavePath, 2)
				__HttpRequest_ErrNotify('_WinHttpReadData_Ex', 'Đã ghi đè lên tập tin cũ tồn tại: "' & $iFileSavePath & '"', '', 'Warning')
			EndIf
			Local $hFileOpen = FileOpen($iFileSavePath, 1 + $iEncodingOfFileSave)
			While 1
				$aCall = DllCall($dll_WinHttp, "bool", 'WinHttpReadData', "handle", $hRequest, "struct*", $tBuffer, "dword", $iBytesPerLoop, 'dword*', 0)
				If @error Or Not $aCall[0] Or Not $aCall[4] Then ExitLoop
				$vNowSizeBytes += $aCall[4]
				If $aCall[4] < $iBytesPerLoop Then
					FileWrite($hFileOpen, BinaryMid(DllStructGetData($tBuffer, 1), 1, $aCall[4]))
					If $iCheckCallbackFunc Then $CallBackFunc_Progress($vNowSizeBytes, $vTotalSizeBytes)
					ExitLoop
				Else
					FileWrite($hFileOpen, DllStructGetData($tBuffer, 1))
				EndIf
				;--------------------------------------------------------------------------------
				If $g___CancelReadWrite Then
					$g___CancelReadWrite = False
					If $iFileSavePath Then FileClose($hFileOpen)
					$tBuffer = ''
					Return SetError(998, __HttpRequest_ErrNotify('_WinHttpReadData_Ex', 'Đã huỷ request', -1), 0)
				ElseIf $iCheckCallbackFunc Then
					$CallBackFunc_Progress($vNowSizeBytes, $vTotalSizeBytes)
				EndIf
			WEnd
			$tBuffer = ''
			FileClose($hFileOpen)

		Else

			While 1
				$aCall = DllCall($dll_WinHttp, "bool", 'WinHttpReadData', "handle", $hRequest, "struct*", $tBuffer, "dword", $iBytesPerLoop, 'dword*', 0)
				If @error Or Not $aCall[0] Or Not $aCall[4] Then ExitLoop
				$vNowSizeBytes += $aCall[4]
				If $aCall[4] < $iBytesPerLoop Then
					$vBinaryData &= BinaryMid(DllStructGetData($tBuffer, 1), 1, $aCall[4])
					If $iCheckCallbackFunc Then $CallBackFunc_Progress($vNowSizeBytes, $vTotalSizeBytes)
					ExitLoop
				Else
					$vBinaryData &= DllStructGetData($tBuffer, 1)
				EndIf
				;--------------------------------------------------------------------------------
				If $g___CancelReadWrite Then
					$g___CancelReadWrite = False
					If $iFileSavePath Then FileClose($hFileOpen)
					$tBuffer = ''
					Return SetError(998, __HttpRequest_ErrNotify('_WinHttpReadData_Ex', 'Đã huỷ request', -1), 0)
				ElseIf $iCheckCallbackFunc Then
					$CallBackFunc_Progress($vNowSizeBytes, $vTotalSizeBytes)
				EndIf
			WEnd
			$tBuffer = ''
			Return $vBinaryData
		EndIf
	EndFunc

	Func _WinHttpConnect2($hSession, $sServerName, $iServerPort)
		Local $aCall = DllCall($dll_WinHttp, "handle", 'WinHttpConnect', "handle", $hSession, 'wstr', $sServerName, "dword", $iServerPort, "dword", 0)
		If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
		Return $aCall[0]
	EndFunc

	Func _WinHttpSetTimeouts2($hInternet, $iConnectTimeout = 30000, $iSendTimeout = 30000, $iReceiveTimeout = 30000)
		DllCall($dll_WinHttp, "bool", 'WinHttpSetTimeouts', "handle", $hInternet, "int", 0, "int", $iConnectTimeout, "int", $iSendTimeout, "int", $iReceiveTimeout)
	EndFunc

	Func _WinHttpCloseHandle2($hInternet)
		DllCall($dll_WinHttp, "bool", 'WinHttpCloseHandle', "handle", $hInternet)
	EndFunc

	Func _WinHttpOpenRequest2($hConnect, $sVerb, $sObjectName = '', $iFlags = 0x40, $sVersion = 'HTTP/1.1')
		Local $aCall = DllCall($dll_WinHttp, "handle", 'WinHttpOpenRequest', "handle", $hConnect, 'wstr', StringUpper($sVerb), 'wstr', $sObjectName, 'wstr', StringUpper($sVersion), 'wstr', '', "ptr", 0, "dword", $iFlags)
		If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
		Return $aCall[0]
	EndFunc

	Func _WinHttpReceiveResponse2($hRequest)
		Local $aCall = DllCall($dll_WinHttp, "bool", 'WinHttpReceiveResponse', "handle", $hRequest, "ptr", 0)
		If Not @error And $aCall[0] Then Return 1
	EndFunc

	Func _WinHttpSetOptionEx2($hInternet, $iOption, $vBuffer = 0, $iNoParam = False)
		Local $tBuffer, $iBuffer
		If $iNoParam Then
			Local $aCall = DllCall($dll_WinHttp, "bool", "WinHttpSetOption", "handle", $hInternet, "dword", $iOption, "ptr", 0, "dword", 0)
			If @error Or Not $aCall[0] Then Return SetError(1, 0, False)
			Return True
		ElseIf IsBinary($vBuffer) Or IsNumber($vBuffer) Then
			$iBuffer = BinaryLen($vBuffer)
			$tBuffer = DllStructCreate("byte[" & $iBuffer & "]")
			DllStructSetData($tBuffer, 1, $vBuffer)
		ElseIf IsDllStruct($vBuffer) Then
			$tBuffer = $vBuffer
			$iBuffer = DllStructGetSize($tBuffer)
		Else
			$tBuffer = DllStructCreate("wchar[" & (StringLen($vBuffer) + 1) & "]")
			$iBuffer = DllStructGetSize($tBuffer)
			DllStructSetData($tBuffer, 1, $vBuffer)
		EndIf
		Local $avResult = DllCall($dll_WinHttp, "bool", 'WinHttpSetOption', "handle", $hInternet, "dword", $iOption, "ptr", DllStructGetPtr($tBuffer), "dword", $iBuffer)
		If @error Or Not $avResult[0] Then Return SetError(2, 0, False)
		Return True
	EndFunc

	Func _WinHttpSetOption2($hInternet, $iOption, $vSetting, $iSize = -1)
		Local $sType
		If IsBinary($vSetting) Then
			$iSize = DllStructCreate("byte[" & BinaryLen($vSetting) & "]")
			DllStructSetData($iSize, 1, $vSetting)
			$vSetting = $iSize
			$iSize = DllStructGetSize($vSetting)
		EndIf
		Switch $iOption
			Case 2 To 7, 12, 13, 31, 36, 58, 63, 68, 73, 74, 77, 79, 80, 83 To 85, 88 To 92, 96, 100, 101, 110, 118
				$sType = "dword*"
				$iSize = 4
			Case 1, 86
				$sType = "ptr*"
				$iSize = 4
				If @AutoItX64 Then $iSize = 8
				If Not IsPtr($vSetting) Then Return SetError(1, 0, 0)
			Case 45
				$sType = "dword_ptr*"
				$iSize = 4
				If @AutoItX64 Then $iSize = 8
			Case 41, 0x1000 To 0x1003
				$sType = "wstr"
				If (IsDllStruct($vSetting) Or IsPtr($vSetting)) Then Return SetError(2, 0, 0)
				If $iSize < 1 Then $iSize = StringLen($vSetting)
			Case 38, 47, 59, 97, 98
				$sType = "ptr"
				If Not (IsDllStruct($vSetting) Or IsPtr($vSetting)) Then Return SetError(3, 0, 0)
			Case Else
				Return SetError(4, 0, 0)
		EndSwitch
		If $iSize < 1 Then
			If IsDllStruct($vSetting) Then
				$iSize = DllStructGetSize($vSetting)
			Else
				Return SetError(5, 0, 0)
			EndIf
		EndIf
		Local $aCall = DllCall($dll_WinHttp, "bool", 'WinHttpSetOption', "handle", $hInternet, "dword", $iOption, $sType, IsDllStruct($vSetting) ? DllStructGetPtr($vSetting) : $vSetting, "dword", $iSize)
		If @error Or Not $aCall[0] Then Return SetError(6, 0, 0)
		Return 1
	EndFunc

	Func _WinHttpQueryOptionEx2($hInternet, $iOption, $iBufferSize = 2048)
		Local $tBufferLength = DllStructCreate("dword")
		DllStructSetData($tBufferLength, 1, $iBufferSize)
		Local $tBuffer = DllStructCreate("byte[" & $iBufferSize & "]")
		Local $avResult = DllCall($dll_WinHttp, "bool", 'WinHttpQueryOption', "handle", $hInternet, "dword", $iOption, "ptr", DllStructGetPtr($tBuffer), "ptr", DllStructGetPtr($tBufferLength))
		If @error Or Not $avResult[0] Then Return SetError(1, 0, "")
		Return $tBuffer
	EndFunc

	Func _WinHttpQueryOption2($hRequest, $iOption)
		Local $aCall = DllCall($dll_WinHttp, "bool", 'WinHttpQueryOption', "handle", $hRequest, "dword", $iOption, "ptr", 0, "dword*", 0)
		If @error Or $aCall[0] Then Return SetError(1, 0, "")
		Local $iSize = $aCall[4], $tBuffer
		Switch $iOption
			Case 34, 41, 81, 82, 93, 0x1000 To 0x1003
				$tBuffer = DllStructCreate("wchar[" & $iSize + 1 & "]")
			Case 1, 21, 78
				$tBuffer = DllStructCreate("ptr")
			Case 0 To 7, 9, 24, 31, 36, 73, 74, 83, 89
				$tBuffer = DllStructCreate("int")
			Case 45
				$tBuffer = DllStructCreate("dword_ptr")
			Case Else
				$tBuffer = DllStructCreate("byte[" & $iSize & "]")
		EndSwitch
		$aCall = DllCall($dll_WinHttp, "bool", 'WinHttpQueryOption', "handle", $hRequest, "dword", $iOption, "struct*", $tBuffer, "dword*", $iSize)
		If @error Or Not $aCall[0] Then Return SetError(2, 0, "")
		Return DllStructGetData($tBuffer, 1)
	EndFunc

	Func _WinHttpSetProxy2($hInternet, $sProxy = "", $sProxyBypass = "")
		Local $tProxy = DllStructCreate("wchar sProxy[" & StringLen($sProxy) + 1 & "];wchar sProxyBypass[" & StringLen($sProxyBypass) + 1 & "]")
		$tProxy.sProxy = $sProxy
		$tProxy.sProxyBypass = $sProxyBypass
		;------------------------------------------------------------
		Local $tProxyInfo = DllStructCreate("dword AccessType;ptr Proxy;ptr ProxyBypass")
		$tProxyInfo.AccessType = 3
		$tProxyInfo.Proxy = DllStructGetPtr($tProxy, 1)
		$tProxyInfo.ProxyBypass = DllStructGetPtr($tProxy, 2)
		;------------------------------------------------------------
		_WinHttpSetOptionEx2($hInternet, 38, $tProxyInfo)
		If @error Then Return SetError(1, 0, 0)
		Return 1
	EndFunc

	Func _WinHttpSetStatusCallback2($hInternet, $pCallback, $nStatusRev)
		DllCall($dll_WinHttp, "ptr", 'WinHttpSetStatusCallback', "handle", $hInternet, "ptr", $pCallback, "dword", $nStatusRev, "ptr", 0)
	EndFunc

	Func _WinHttpSetCredentials2($hRequest, $sUserName, $sPassword, $iAuthTargets, $iAuthScheme)
		;$iAuthTargets: Server = 0x0 ;Proxy = 0x1
		;$iAuthScheme: BASIC = 0x1 ;NTLM = 0x2 ;PASSPORT = 0x4 ;DIGEST = 0x8 ;NEGOTIATE = 0x10
		If $iAuthScheme = 0x4 Then
			_WinHttpSetOption2($hRequest, 83, 0x10000000) ;OPTION_CONFIGURE_PASSPORT_AUTH = ENABLE_PASSPORT_AUTH
			If $iAuthTargets = 0x0 Then
				_WinHttpSetOption2($hRequest, 0x1000, $sUserName) ;OPTION_USERNAME
				_WinHttpSetOption2($hRequest, 0x1001, $sPassword) ;OPTION_PASSWORD
			Else
				_WinHttpSetOption2($hRequest, 0x1002, $sUserName) ;OPTION_PROXY_USERNAME
				_WinHttpSetOption2($hRequest, 0x1003, $sPassword) ;OPTION_PROXY_PASSWORD
			EndIf
		EndIf
		Local $aCall = DllCall($dll_WinHttp, "bool", 'WinHttpSetCredentials', "handle", $hRequest, "dword", $iAuthTargets, "dword", $iAuthScheme, 'wstr', $sUserName, 'wstr', $sPassword, "ptr", 0)
		If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
		Return 1
	EndFunc

	Func _WinHttpQueryAuthSchemes2($hRequest) ;Return AuthScheme, AuthTarget, SupportedSchemes
		Local $aCall = DllCall($dll_WinHttp, "bool", "WinHttpQueryAuthSchemes", "handle", $hRequest, "dword*", 0, "dword*", 0, "dword*", 0)
		If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
		Local $aRet = [$aCall[3], $aCall[4], $aCall[2]]
		Return $aRet
	EndFunc
#EndRegion



#Region <Chạy code javascript, php>
	Func _JSON_Beautify($sJSON)
		Local $rq = _HttpRequest(2, 'https://jsonformatter.curiousconcept.com/process', 'jsondata=' & _URIEncode($sJSON) & '&jsonstandard=1&jsontemplate=1')
		Return StringRegExpReplace(StringRegExpReplace(StringRegExpReplace($rq, '(?i).*?","jsoncopy":"(.*?)","template":"1"\}.+', '${1}'), '\\(\W)', '\1'), '\\[rn]', @CRLF)
	EndFunc

	Func _JS_Beautify($jsCode)
		$jsCode = StringRegExp(_HttpRequest(2, 'https://javascriptbeautifier.com/compress', 'js_code=' & _URIEncode($jsCode) & '&js_code_result=&beautify=true'), '(?is)"mini":"(.+?)"\}$', 1)
		If @error Or StringInStr($jsCode[0], 'undefined","error"', 0, 1, 1, 30) Then Return SetError(1, __HttpRequest_ErrNotify('_JS_Beautify', 'Làm đẹp Js thất bại'), '')
		Return StringRegExpReplace(StringReplace($jsCode[0], '\n', @CRLF), '\\([''"])', '$1')
	EndFunc

	Func _JS_Compress($jsCode)
		$jsCode = _HttpRequest(2, 'https://javascript-minifier.com/raw', 'input=' & _URIEncode($jsCode))
		If @error Or $jsCode = '' Or StringInStr($jsCode, '// Error:', 0, 1, 1, 10) Then Return SetError(1, __HttpRequest_ErrNotify('_JS_Compress', 'Làm gọn Js thất bại'), $jsCode)
		Return $jsCode
	EndFunc

	Func _JS_DeObfuscator($jsCode)
		$jsCode = StringRegExp(_HttpRequest(2, 'https://www.javascriptdeobfuscator.com/', 'code=' & _URIEncode($jsCode)), '(?is)<textarea[^>]+ name="code">(.*?)</textarea>', 1)
		If @error Then SetError(1, __HttpRequest_ErrNotify('_JS_DeObfuscator', 'Giải rối Js thất bại'), '')
		Return StringRegExpReplace($jsCode[0], '[\x00-\x19]', '')
	EndFunc

	Func _JS_ToStringAu3($jsMode = 0) ;0: Chỉ chuyển sang string, 1: Compress trước khi chuyển, 2: beautify trước khi chuyển
		Local $jsCode = ClipGet()
		If Not $jsCode Then Return SetError(1, MsgBox(4096, 'Lỗi', 'Hãy copy đoạn js cần chuyển thành string Au3 trước khi chạy hàm này'), '')
		;-----------------------------------------------------------------------------------------------------------
		Switch $jsMode
			Case 1
				$jsCode = _JS_Compress($jsCode)
			Case 2
				$jsCode = _JS_Beautify($jsCode)
		EndSwitch
		If @error Or Not $jsCode Then Return SetError(2, __HttpRequest_ErrNotify('_JS_ToStringAu3', 'Làm đẹp/Làm gọn Js thất bại'), '')
		;-----------------------------------------------------------------------------------------------------------
		$jsCode = StringStripCR(StringRegExpReplace($jsCode, '(?m)^\h*$[\r\n]+', ''))
		$jsCode = StringRegExpReplace(StringReplace($jsCode, "'", "''", 0, 1), '(?m)^', "'")
		$jsCode = StringTrimRight(StringRegExpReplace($jsCode, '(?m)($)', "' & @CRLF & _"), 3)
		ClipPut($jsCode)
		MsgBox(4096, 'Thông báo', 'Đã lưu kết quả chuyển đổi vào Clipboard')
	EndFunc


	Func _JS_Execute($LibraryJS, $sCodeJS, $Name_Var_Return_Val, $ModeIE = False, $PathTempLibJS = Default)
		If FileExists($PathTempLibJS) Then
			If StringRight($PathTempLibJS, 1) <> '\' Then $PathTempLibJS &= '\'
		Else
			$PathTempLibJS = @TempDir & '\'
		EndIf
		Local $TempPath, $hOpen, $iError = 0, $sLibraryJS = ''
		;--------------------------------------------------------------------------------------------------
		If FileExists($sCodeJS) Then $sCodeJS = FileRead($sCodeJS)
		If StringInStr($Name_Var_Return_Val, '.', 1, 1) Then
			Local $Name_Var_Return_Val_tmp = $Name_Var_Return_Val
			$Name_Var_Return_Val = StringReplace($Name_Var_Return_Val, '.', '', 1, 1)
			$sCodeJS = StringReplace($sCodeJS, $Name_Var_Return_Val_tmp, $Name_Var_Return_Val, 1, 1)
		EndIf
		$sCodeJS = StringRegExpReplace($sCodeJS, '(?i)(location.href=)(".*?"|''.*?'')', '')
		;--------------------------------------------------------------------------------------------------
		If $LibraryJS Or IsArray($LibraryJS) Then
			If Not IsArray($LibraryJS) Then $LibraryJS = StringSplit($LibraryJS, '|', 2)
			For $i = 0 To UBound($LibraryJS) - 1
				If StringRegExp($LibraryJS[$i], '(?i)^https?://') Then
					$TempPath = $PathTempLibJS & StringRight(StringRegExpReplace($LibraryJS[$i], '(?i)(\.js|\W+)', '-'), 200) & '.js'
					If FileExists($TempPath) And FileGetSize($TempPath) > 2 Then
						$LibraryJS[$i] = FileRead($TempPath)
					Else
						$LibraryJS[$i] = _HttpRequest(2, $LibraryJS[$i])
						If @error Or Not $LibraryJS[$i] Then $iError = 301
						$hOpen = FileOpen($TempPath, 2 + 32 + 8)
						FileWrite($hOpen, $LibraryJS[$i])
						FileClose($hOpen)
					EndIf
				Else
					$LibraryJS[$i] = FileRead($LibraryJS[$i])
				EndIf
				$sLibraryJS &= $LibraryJS[$i] & ';' & @CRLF
			Next
		EndIf
		If Not $g___JsLibJSON Then
			$g___JsLibJSON &= 'if(typeof JSON!=="object"){JSON={}}(function(){"use strict";var rx_one=/^[\],:{}\s]*$/;var rx_two=/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g;var rx_three=/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g;var rx_four=/(?:^|:|,)(?:\s*\[)+/g;var rx_escapable=/[\\"\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;var rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;function f(n){return(n<10)?"0"+n:n}function this_value(){return this.valueOf()}if(typeof Date.prototype.toJSON!=="function"){Date.prototype.toJSON=function(){return isFinite(this.valueOf())?(this.getUTCFullYear()+"-"+f(this.getUTCMonth()+1)+"-"+f(this.getUTCDate())+"T"+f(this.getUTCHours())+":"+f(this.getUTCMinutes())+":"+f(this.getUTCSeconds())+"Z"):null};Boolean.prototype.toJSON=this_value;Number.prototype.toJSON=this_value;String.prototype.toJSON=this_value}var gap;var indent;var meta;var rep;function quote(string){rx_escapable.lastIndex=0;return rx_escapable.test(string)?"\""+string.replace(rx_escapable,function(a){var c=meta[a];return typeof c==="string"?c:"\\u"+("0000"+a.charCodeAt(0).toString(16)).slice(-4)})+"\"":"\""+string+"\""}function str(key,holder){var i;var k;var v;var length;var mind=gap;var partial;var value=holder[key];if(value&&typeof value==="object"&&typeof value.toJSON==="function"){value=value.toJSON(key)}if(typeof rep==="function"){value=rep.call(holder,key,value)}switch(typeof value){case"string":return quote(value);case"number":return(isFinite(value))?String(value):"null";case"boolean":case"null":return String(value);case"object":if(!value){return"null"}'
			$g___JsLibJSON &= 'gap+=indent;partial=[];if(Object.prototype.toString.apply(value)==="[object Array]"){length=value.length;for(i=0;i<length;i+=1){partial[i]=str(i,value)||"null"}v=partial.length===0?"[]":gap?("[\n"+gap+partial.join(",\n"+gap)+"\n"+mind+"]"):"["+partial.join(",")+"]";gap=mind;return v}if(rep&&typeof rep==="object"){length=rep.length;for(i=0;i<length;i+=1){if(typeof rep[i]==="string"){k=rep[i];v=str(k,value);if(v){partial.push(quote(k)+((gap)?": ":":")+v)}}}}else{for(k in value){if(Object.prototype.hasOwnProperty.call(value,k)){v=str(k,value);if(v){partial.push(quote(k)+((gap)?": ":":")+v)}}}}v=partial.length===0?"{}":gap?"{\n"+gap+partial.join(",\n"+gap)+"\n"+mind+"}":"{"+partial.join(",")+"}";gap=mind;return v}}if(typeof JSON.stringify!=="function"){meta={"\b":"\\b","\t":"\\t","\n":"\\n","\f":"\\f","\r":"\\r","\"":"\\\"","\\":"\\\\"};JSON.stringify=function(value,replacer,space){var i;gap="";indent="";if(typeof space==="number"){for(i=0;i<space;i+=1){indent+=" "}}else if(typeof space==="string"){indent=space}rep=replacer;if(replacer&&typeof replacer!=="function"&&(typeof replacer!=="object"||typeof replacer.length!=="number")){throw new Error("JSON.stringify");}return str("",{"":value})}}if(typeof JSON.parse!=="function"){JSON.parse=function(text,reviver){var j;function walk(holder,key){var k;var v;var value=holder[key];if(value&&typeof value==="object"){for(k in value){if(Object.prototype.hasOwnProperty.call(value,k)){v=walk(value,k);if(v!==undefined){value[k]=v}else{delete value[k]}}}}return reviver.call(holder,key,value)}text=String(text);rx_dangerous.lastIndex=0;if(rx_dangerous.test(text)){text=text.replace(rx_dangerous,function(a){return("\\u"+("0000"+a.charCodeAt(0).toString(16)).slice(-4))})}if(rx_one.test(text.replace(rx_two,"@").replace(rx_three,"]").replace(rx_four,""))){j=eval("("+text+")");return(typeof reviver==="function")?walk({"":j},""):j}throw new SyntaxError("JSON.parse");}}}())'
		EndIf
		;-----------------------------------------------------------------------------------------------------------
		$sCodeJS = '<script>' & @CRLF & $sCodeJS & '; document.write(' & $Name_Var_Return_Val & ');' & @CRLF & '</script>' & @CRLF
		$sCodeJS = '<script>' & @CRLF & $sLibraryJS & ';' & @CRLF & $g___JsLibJSON & ';' & @CRLF & '</script>' & @CRLF & $sCodeJS
		If $ModeIE Then
			Local $oIE = ObjCreate("InternetExplorer.Application")
			With $oIE
				.navigate('about:blank')
				.document.write($sCodeJS)
				.document.close()
				While .busy()
				WEnd
				$sCodeJS = .document.body.innerText
				If StringRight($sCodeJS, 1) == ' ' Then $sCodeJS = StringTrimRight($sCodeJS, 1)
				While .busy()
				WEnd
				.quit()
				;ProcessClose('ielowutil.exe')
			EndWith
		Else
			$sCodeJS = _HTML_Execute($sCodeJS)
			If @error Then $iError = 302
		EndIf
		;-----------------------------------------------------------------------------------------------------------
		Return SetError($iError, '', $sCodeJS)
	EndFunc

	Func _PHP_Execute($phpData, $Name_Var_Return_Val, $BinaryMode = False, $phpVersion = Default)
		If Not $phpVersion Or $phpVersion = Default Then $phpVersion = '7.1.0'
		If StringLeft($Name_Var_Return_Val, 1) <> '$' Then $Name_Var_Return_Val = '$' & $Name_Var_Return_Val
		If Not StringInStr($phpData, '<?php') Then $phpData = '<?php' & @CRLF & $phpData
		If StringRegExp($phpData, '(?is)\?>\h*?$') Then $phpData = StringRegExpReplace($phpData, '(?is)\?>\h*?$', '')
		Local $rq = _HttpRequest($BinaryMode ? -2 : 2, 'http://sandbox.onlinephpfunctions.com/', 'code=' & _URIEncode($phpData & ';' & @CRLF & 'echo ' & $Name_Var_Return_Val & ';') & '&phpVersion=' & StringReplace($phpVersion, '.', '_', 0, 1) & '&output=Textbox&ajaxResult=1')
		$phpData = StringRegExp($rq, '(?i)' & $BinaryMode ? '3C746578746172656120(?:..)+3E(.*?)3C2F74657874617265613E' : '<textarea [^>]+>(.*?)</textarea>', 1)
		If @error Then Return SetError(1, '', '')
		Return $BinaryMode ? Binary('0x' & $phpData[0]) : _HTMLDecode($phpData[0])
	EndFunc
#EndRegion


#Region <HTML Entities>
	Func __HTML_Entities_Decode($sHTML, $iModeIE = False)
		If $iModeIE = Default Then $iModeIE = False
		If $iModeIE Then
			$sHTML = _HTML_Execute(StringReplace($sHTML, '&#xD;', '<hr>'))
		Else
			If Not IsObj($g___oDicEntity) Then
				$g___oDicEntity = __HTML_Entities_Init()
				If @error Then Return SetError(2, __HttpRequest_ErrNotify('__HTML_Entities_Decode', 'Khởi tạo __HTML_Entities_Init thất bại'), $sHTML)
			EndIf
			;---------------------------------------------------------------------
			Local $aText = StringRegExp($sHTML, '\&\#(\d+)\;', 3)
			If Not @error Then
				For $i = 0 To UBound($aText) - 1
					$sHTML = StringReplace($sHTML, '&#' & $aText[$i] & ';', ChrW($aText[$i]), 0, 1)
				Next
			EndIf
			;---------------------------------------------------------------------
			$aText = StringRegExp($sHTML, '\&([a-zA-Z]{2,10})\;', 3)
			If Not @error Then
				For $i = 0 To UBound($aText) - 1
					$sHTML = StringReplace($sHTML, '&' & $aText[$i] & ';', $g___oDicEntity.item($aText[$i]), 0, 1)
				Next
			EndIf
		EndIf
		Return $sHTML
	EndFunc

	Func __HTML_Entities_Init()
		If $g___aChrEnt == '' Then
			Local $aisEntities[246][2] = [[34, 'quot'], [38, 'amp'], [39, 'apos'], [60, 'lt'], [62, 'gt'], [160, 'nbsp'], [161, 'iexcl'], [162, 'cent'], [163, 'pound'], [164, 'curren'], [165, 'yen'], [166, 'brvbar'], [167, 'sect'], [168, 'uml'], [169, 'copy'], [170, 'ordf'], [171, 'laquo'], [172, 'not'], [173, 'shy'], [174, 'reg'], [175, 'macr'], [176, 'deg'], [177, 'plusmn'], [180, 'acute'], [181, 'micro'], [182, 'para'], [183, 'middot'], [184, 'cedil'], [186, 'ordm'], [187, 'raquo'], [191, 'iquest'], [192, 'Agrave'], [193, 'Aacute'], [194, 'Acirc'], [195, 'Atilde'], [196, 'Auml'], [197, 'Aring'], [198, 'AElig'], [199, 'Ccedil'], [200, 'Egrave'], [201, 'Eacute'], [202, 'Ecirc'], [203, 'Euml'], [204, 'Igrave'], [205, 'Iacute'], [206, 'Icirc'], [207, 'Iuml'], [208, 'ETH'], [209, 'Ntilde'], [210, 'Ograve'], [211, 'Oacute'], [212, 'Ocirc'], [213, 'Otilde'], [214, 'Ouml'], [215, 'times'], [216, 'Oslash'], [217, 'Ugrave'], [218, 'Uacute'], [219, 'Ucirc'], [220, 'Uuml'], _
					[221, 'Yacute'], [222, 'THORN'], [223, 'szlig'], [224, 'agrave'], [225, 'aacute'], [226, 'acirc'], [227, 'atilde'], [228, 'auml'], [229, 'aring'], [230, 'aelig'], [231, 'ccedil'], [232, 'egrave'], [233, 'eacute'], [234, 'ecirc'], [235, 'euml'], [236, 'igrave'], [237, 'iacute'], [238, 'icirc'], [239, 'iuml'], [240, 'eth'], [241, 'ntilde'], [242, 'ograve'], [243, 'oacute'], [244, 'ocirc'], [245, 'otilde'], [246, 'ouml'], [247, 'divide'], [248, 'oslash'], [249, 'ugrave'], [250, 'uacute'], [251, 'ucirc'], [252, 'uuml'], [253, 'yacute'], [254, 'thorn'], [255, 'yuml'], [338, 'OElig'], [339, 'oelig'], [352, 'Scaron'], [353, 'scaron'], [376, 'Yuml'], [402, 'fnof'], [710, 'circ'], [732, 'tilde'], [913, 'Alpha'], [914, 'Beta'], [915, 'Gamma'], [916, 'Delta'], [917, 'Epsilon'], [918, 'Zeta'], [919, 'Eta'], [920, 'Theta'], [921, 'Iota'], [922, 'Kappa'], [923, 'Lambda'], [924, 'Mu'], [925, 'Nu'], [926, 'Xi'], [927, 'Omicron'], [928, 'Pi'], [929, 'Rho'], _
					[931, 'Sigma'], [932, 'Tau'], [933, 'Upsilon'], [934, 'Phi'], [935, 'Chi'], [936, 'Psi'], [937, 'Omega'], [945, 'alpha'], [946, 'beta'], [947, 'gamma'], [948, 'delta'], [949, 'epsilon'], [950, 'zeta'], [951, 'eta'], [952, 'theta'], [953, 'iota'], [954, 'kappa'], [955, 'lambda'], [956, 'mu'], [957, 'nu'], [958, 'xi'], [959, 'omicron'], [960, 'pi'], [961, 'rho'], [962, 'sigmaf'], [963, 'sigma'], [964, 'tau'], [965, 'upsilon'], [966, 'phi'], [967, 'chi'], [968, 'psi'], [969, 'omega'], [977, 'thetasym'], [978, 'upsih'], [982, 'piv'], [8194, 'ensp'], [8195, 'emsp'], [8201, 'thinsp'], [8204, 'zwnj'], [8205, 'zwj'], [8206, 'lrm'], [8207, 'rlm'], [8211, 'ndash'], [8212, 'mdash'], [8216, 'lsquo'], [8217, 'rsquo'], [8218, 'sbquo'], [8220, 'ldquo'], [8221, 'rdquo'], [8222, 'bdquo'], [8224, 'dagger'], [8225, 'Dagger'], [8226, 'bull'], [8230, 'hellip'], [8240, 'permil'], [8242, 'prime'], [8243, 'Prime'], [8249, 'lsaquo'], [8250, 'rsaquo'], _
					[8254, 'oline'], [8260, 'frasl'], [8364, 'euro'], [8465, 'image'], [8472, 'weierp'], [8476, 'real'], [8482, 'trade'], [8501, 'alefsym'], [8592, 'larr'], [8593, 'uarr'], [8594, 'rarr'], [8595, 'darr'], [8596, 'harr'], [8629, 'crarr'], [8656, 'lArr'], [8657, 'uArr'], [8658, 'rArr'], [8659, 'dArr'], [8660, 'hArr'], [8704, 'forall'], [8706, 'part'], [8707, 'exist'], [8709, 'empty'], [8711, 'nabla'], [8712, 'isin'], [8713, 'notin'], [8715, 'ni'], [8719, 'prod'], [8721, 'sum'], [8722, 'minus'], [8727, 'lowast'], [8730, 'radic'], [8733, 'prop'], [8734, 'infin'], [8736, 'ang'], [8743, 'and'], [8744, 'or'], [8745, 'cap'], [8746, 'cup'], [8747, 'int'], [8764, 'sim'], [8773, 'cong'], [8776, 'asymp'], [8800, 'ne'], [8801, 'equiv'], [8804, 'le'], [8805, 'ge'], [8834, 'sub'], [8835, 'sup'], [8836, 'nsub'], [8838, 'sube'], [8839, 'supe'], [8853, 'oplus'], [8855, 'otimes'], [8869, 'perp'], [8901, 'sdot'], [8968, 'lceil'], [8969, 'rceil'], [8970, 'lfloor'], [8971, 'rfloor'], [9001, 'lang'], [9002, 'rang'], [9674, 'loz'], [9824, 'spades'], [9827, 'clubs'], [9829, 'hearts'], [9830, 'diams']]
			$g___aChrEnt = $aisEntities
		EndIf
		$g___oDicEntity = ObjCreate("Scripting.Dictionary")
		If @error Or Not IsObj($g___oDicEntity) Then Return SetError(1)
		For $i = 0 To UBound($g___aChrEnt) - 1
			$g___oDicEntity.Add($g___aChrEnt[$i][1], ChrW($g___aChrEnt[$i][0]))
		Next
		Return $g___oDicEntity
	EndFunc

	Func __HTML_RegexpReplace($sData, $Escape_Character_Head, $Escape_Character_Tail, $iHexLength, $isHexNumber = True)
		Local $Chr_or_WChar = ($iHexLength = 2 ? 'Chr' : 'ChrW')
		If $Escape_Character_Tail And $Escape_Character_Tail <> Default Then $Chr_or_WChar = 'ChrW'
		Local $sResult = Call('Execute', '"' & StringRegExpReplace(StringReplace($sData, '"', '""', 0, 1), '(?i)' & StringReplace($Escape_Character_Head, '\', '\\', 0, 1) & '([[:xdigit:]]{' & $iHexLength & '})' & $Escape_Character_Tail, '" & ' & $Chr_or_WChar & '(' & ($isHexNumber ? '0x' : '') & '${1}) & "') & '"')
		If $sResult == '' Then Return SetError(1, '', $sData)
		Return StringRegExpReplace($sResult, '\\([\\/"''\?:])', '\1')
	EndFunc

	Func _HTML_AbsoluteURL($sSource, $sURL, $sAdditional_Pattern = '', $sProtocol = '')
		If Not StringRegExp($sSource, '(?i)<\h*?base .*?href\h*?=') Then
			$sSource = '<base href="' & $sURL & '"/><script>var _b = document.getElementsByTagName("base")[0], _bH = "' & $sURL & '";if (_b && _b.href != _bH) _b.href = _bH;</script>' & @CRLF & $sSource
		EndIf
		;-------------------------------------------------------------------------------------------------------
		If $sAdditional_Pattern Then $sAdditional_Pattern &= '|'
		Local $basePattern = '(?i)(' & $sAdditional_Pattern & '(?:window\.location|\W(?:src|href)|v-bind:src|param name\h*?=\h*?["'']movie["'']\h+value)\h*?=\h*?["'']*?|attr\([''"]src[''"]\h*?,\h*?[''"])(?!https?:|javascript:|\&|\#)'
		;-------------------------------------------------------------------------------------------------------
		$sURL = StringRegExpReplace($sURL, '(?i)^(.*?)/[^/]+\.(?:php|html|aspx).*$', '$1')
		$sURL = StringRegExpReplace($sURL, '/$', '')
		;-------------------------------------------------------------------------------------------------------
		Local $aURL = StringRegExp($sURL, '(?i)^(https?://[^/]+)(/?)(.*)/?$', 3)
		If @error Then Return SetError(1, '', $sSource)
		;href='//' ----------------------------------------------------------
		$sSource = StringRegExpReplace($sSource, $basePattern & '//', '$1' & $sProtocol & '://')
		;href='/' ----------------------------------------------------------
		$sSource = StringRegExpReplace($sSource, $basePattern & '/', '$1' & $aURL[0] & '/')
		;href='./' ----------------------------------------------------------
		$sSource = StringRegExpReplace($sSource, $basePattern & '\./', '$1' & $sURL & '/')
		;href='' ----------------------------------------------------------
;~ 		$sSource = StringRegExpReplace($sSource, $basePattern & '([^/\."''])', '$1' & $sURL & '/$2')
		;href='../' -------------------------------------------------------------------
		Local $regSource = StringRegExp($sSource, $basePattern & '((?:\.\./)+)', 3), $memReg = '|', $sRegAttach
		If Not @error Then
			Local $baseURL = ''
			For $i = 0 To UBound($regSource) - 1 Step 2
				$sRegAttach = $regSource[$i] & $regSource[$i + 1]
				If StringInStr($memReg, '|' & $sRegAttach & '|', 0, 1) Then ContinueLoop
				$baseURL = $aURL[2]
				For $j = 1 To (StringLen($regSource[$i + 1]) / 3) + 1 ;số lần Back
					$baseURL = StringRegExpReplace($baseURL, '(?:/|^)[^/]+$', '')
				Next
				$sSource = StringRegExpReplace($sSource, '\Q' & $sRegAttach & '\E', $regSource[$i] & $aURL[0] & '/' & ($baseURL ? $baseURL & '/' : ''))
				$memReg &= $sRegAttach & '|'
			Next
		EndIf
		Return $sSource
	EndFunc

	Func _HTML_Execute($sHTML, $iElement = '', $iAttribute = '', $iSpecifiedValue = '', $iReturnHTML = False)
		If $sHTML == '' Then Return SetError(1, __HttpRequest_ErrNotify('_HTML_Execute', 'Tham số $sHTML đưa vào rỗng'), '')
		Local $sResult = '', $oFind = 0, $l___iError = 0
		Local $oHTML = ObjCreate("HTMLFILE")
		If @error Or Not IsObj($oHTML) Then Return SetError(2, __HttpRequest_ErrNotify('_HTML_Execute', 'Tạo HTMLFile Object thất bại'), $sHTML)
		If $iElement = Default Then $iElement = ''
		If $iAttribute = Default Then $iAttribute = ''
		If $iReturnHTML = Default Then $iReturnHTML = False
		With $oHTML
			;.open()
			.write($sHTML)
			If @error Then Return SetError(3, __HttpRequest_ErrNotify('_HTML_Execute', 'HTMLFile Object không thể truy vấn dữ liệu HTML đưa vào'), $sHTML)
			Select
				Case Not $iElement And Not $iAttribute
					Local $oBody = .body
					If @error Or Not IsObj($oBody) Then Return SetError(4, __HttpRequest_ErrNotify('_HTML_Execute', 'Không thể xử lý dữ liệu HTML đã nạp vào'), $sHTML)
					$sResult = ($iReturnHTML ? $oBody.innerHTML : $oBody.innerText)
				Case $iElement And Not $iAttribute
					__HttpRequest_ErrNotify('_HTML_Execute', 'Phải điền giá trị cho tham số $iAttribute')
					$l___iError = 5
				Case $iAttribute And Not $iSpecifiedValue
					__HttpRequest_ErrNotify('_HTML_Execute', 'Phải điền giá trị cho tham số $iSpecifiedValue')
					$l___iError = 6
				Case Else
					Local $oElements = ($iElement ? .getElementsByTagName($iElement) : .All)
					If Not @error And IsObj($oElements) Then
						For $oElement In $oElements
							Switch $iAttribute
								Case 'class'
									If $oElement.classname = $iSpecifiedValue Then $oFind = 1
								Case 'id'
									If $oElement.id = $iSpecifiedValue Then $oFind = 1
								Case 'name'
									If $oElement.name = $iSpecifiedValue Then $oFind = 1
								Case 'type'
									If $oElement.type = $iSpecifiedValue Then $oFind = 1
								Case 'href'
									If $oElement.href = $iSpecifiedValue Then $oFind = 1
							EndSwitch
							If $oFind = 1 Then
								$sResult = ($iReturnHTML ? $oElement.innerHTML : $oElement.innerText)
								ExitLoop
							EndIf
						Next
					Else
						$l___iError = 7
					EndIf
			EndSelect
			.close()
		EndWith
		$oHTML = ''
		If $l___iError Then Return SetError($l___iError, '', $sHTML)
		Return $sResult
	EndFunc
#EndRegion



#Region <INTERNAL FUNCTIONS>
	Func __Gzip_Uncompress($sBinaryData)
		If Not StringRegExp(BinaryMid($sBinaryData, 1, 1), '(?i)0x(1F|08|8B)') Then Return SetError(1, __HttpRequest_ErrNotify('__Gzip_Uncompress', 'Chuỗi binary này không phải định dạng của nén Gzip'), $sBinaryData)
		If Not $g___JsLibGunzip Then
			;Compact zlib, deflate, inflate, zip library in JavaScript: https://github.com/imaya/zlib.js/ - Thanks imaya
			$g___JsLibGunzip &= 'G7wAKGZ1bmN0aW8AbigpeyJ1c2UAIHN0cmljdCICOwW4IGsodCl7AHRocm93IHR9AHZhciBVPXZvAGlkIDAscz10CGhpcwdSdCh0LAhyKXsBRmUsaT0AdC5zcGxpdCgAIi4iKSxuPXMAOyEoaVswXWkAbiBuKSYmbi4AZXhlY1NjcmkkcHQLDSgiAUEiKwEBLCk7Zm9yKDsAaS5sZW5ndGgAJiYoZT1pLnMAaGlmdCgpKTsCKQUYfHxyPT09AFU/bj1uW2Vdij8BBDoBBD17fQMHAnICunIsRT0idQBuZGVmaW5lZAAiIT10eXBlbwBmIFVpbnQ4QUBycmF5JiaVDzESNhwQMzIYEERhdCBhVmlld4JnbmVQdyhFPwc6OgIcKQAoMjU2KSxyPRAwO3I8AAU7KysscikBfwKoPYB8cikAPj4+MTtlO2WhgAM9MSkwh713gb0ULGWDvmkAtyJudUBtYmVyIj2FcXI4P3I6gC2A2g4NZT8YZTp0hFgCKGk9LQIxwBA3JnM7bi0iLQIiaT1pAB04XgBhWzI1NSYoaYBedFtyXSldwzMgPXM+PjOCCnIrwD04KWk9KJEAEhCzgTRLFSsxwBWQBTKTBaozkwU0kwU1kwU2kwUCN4AFO3JldHVyAG4oNDI5NDk2QDcyOTVeaUEnMAFClGk9WzAsMTkAOTY5NTk4OTQALDM5OTM5MTkQNzg4LIBwNzUyBDQ3QAUxMjQ2MwA0MTM3LDE4OIA2MDU3NjE1gAoAMTU2MjE2ODUALDI2NTczOTIEMDOAAjQ5MjY4ADI3NCwyMDQ0IDUwODMyQBU3N4AyMTE1MjMwQBUANDcxNzc4NjS4LDE2gCTAHgANMQBlADYxMDIxLDM4ADg3NjA3MDQ3ACwyNDI4NDQ0ADA0OSw0OTg1ADM2NTQ4LDE3ADg5OTI3NjY2ACw0MDg5MDE2AjagAjIyMjcwNmAxMjE0LGAO4AQ4hDYxYBU0MzI14BUAMyw0MTA3NTgRYAAzLDIgETY3N9A2MzksoQM4IB4gIEA2ODQ3NzdAFCxANDI1MTEygBcyICwyMzIxoBk2MwA2LDMzNTYzM0Q0OCAgNjYxQRE2AjWgCjk1MzAyNwo1IBgzwAIxNTMxADcsOTk3MDczADA5NiwxMjgxQyAEQCYsMzU3QBg1lDMzoAo3oCk4OKAbACwxMDA2ODg4mDE0NWAFIRU3NiAMijOgLjEAGzI5LAAdQaAyMjQ0MyxAHTACOeAoMiwxMTE5BjAhBwArNjg2NTEANzIwNiwyODnEODDgMDI4LGAlYC+sNDVAIGADMiATMAArADcwNTAxNTc1u6AKoA82ACdAGKEHNkAToUAgMzczNSA4NIEdIcBBNTQzMAAMMjEQODEwNIBDLDU2ADU1MDcyNTMsJeAVNCA/NzOgCjQ4FcAlMWALLEAeOTQzADYzMDMsNjcxBaAOOYBAMTU5NDF7ABOBQDOANIAEgUBAJDMANDc4MTIsNzlANTgzNTUyACs0u2BFwCkyQEshHmBVN0ABwDA2MDE0OYAPwVSQN'
			$g___JsLibGunzip &= 'DE0NkAyLDOgSaegHIBMwDI5MEA3MsBBtDIzQEs54AvgIzdgFmY0ABxgOTYzoE9ANDb2OOAfYD4zQFfgIyBGYSG6MMBaNyBRQVOhJTCgPUtAQEEGMwAYMzcARzMYMDA04AMgBTY1NqdARyAgYQE4McBJNABfV0FHwElgMTIgHzlgGTjJYFk5NaBXLDRgOkBUQeBVMjIzODDADTaSOGAVNjZgYzg3AFqhoCAzNzA5wB40oEmHYFvAKqEuNjI1MIECX8A9QDmAB2BfgCE4gF4ybwAg4GIgGCAgM8BsQCswSSAgMjTgCzc1gEwxejbASTXgcEA0YGWgbDcBAGozNjI2NzAzgjIAYjIyNDk5gGsDIE8gNjUzNTk2MIgsOThAhDE0OKBJUDc0NzDAcDlADTXwNjkwM6AKYBVggmA9ijigXDFAAjYwNAAxKYFUNTIAbDMAKzU1QDQwNzk5OcAKMYYzgFigcTYsODcAj0uAAsAKOQAdNDOgRywjAGZgKzE4NeBkMTQhwDQ0NDY3QFc1OEY04ESAFTg1MgA2NhmBTDcw4IVAADksMeIz4HMzMzlgfsCBYDxqM6AvM1BJM/ALQAEza1AsUB4xoSc0oEGwCDB0OSyAGDEwCOBJ8Aow7XBIOZAbUQg1gD6gTFAqgjEwKzA1NCw3gBvEMzjwQSwyOfADIA5UNTDRPzKQBDTwKjGGNfBI8AYwNyw3gSsIMTg3YA8wODI27wAZYDrAFQAZMsBUUCZwNpY5QB9hKjngHDQ2wDVcNjJAQmA/4Cw1EDAwe+AOkAIzoBegMlEqcTE2vUAfLHBBcCHRJHA1OKA5RXAbM0ABOTE4QS4y3DQ4QFtALeAFMgA0QELhMCowNTM34RpQPDAJ3DE3AD9ABaAQOYE0sDvSNkAHNjeQCjLgROA/v0AYoTNATLBSAASwXDQABPcwINBA4Vc3EEcAMtBDsAePME8AEHAyIEsyMTUQEdY50DtAUzagFDLgKgAd3XArNyAfQAFAIzHQTHAf/DMxYQ8gYnBKAFOBKOAs2bAcNTVwHlAFOCBB8AUbADXBHTdQZ7A/OTc5rbBiNlELYFs4wEU3cCN5UCc5MGBTgCexKmBJONYsYFagbTeRLzNAMPBhOjIQEDXQGaAvEWo3OXA3MzYwQChAVzBXMsY4kBRAVzE5NhBncECxEAk0NzTQCyAQObApbDc1sAVAVzYgAeEqMV4wMAdgaiAgYEE3wEkwGYBOMjjwUJAlNjks4DgyOTMycCWRL4BB2DM1MSBoQHMy0W6RVFQxNsBwNOEeNoAtOHuAMgA1M0AaIBvQQYAtMzgzNjlwBsA9gEI3OLXxMDWQVTawPyAQNCEmrDA4kXcQJDMQPjTAZIw5OQAokA04LDfAZXPAbDAsMTUAVXBSoDUze8B4ERo5cEAQHIACcCgsj3F90CYADCEcMzIw0BUbQHzwPjfQVfAuOTY1lfASMKJUNvB2MjlQIPmQPTM1kHiQAlB4QFXgHt/AGqBtgHOwZsAPMBBNgEYzQHWwDjcx0H5wHzQxR8CDgH0QQzQwNhBFMtWAeDIgPTLwCzEQCsAuQ5AAIIYyNzI2QAkzgDg1NTk5MDJxG34woD4xRfEF8C5gH1AFMtY2IAsAcTKQKzUxA7CDjeAkNZAOEB4yLDUgOLUwRTmwHDGgGiEcMTAm/jMSGvE+MBPQJMBSUAFhQntQXYCBOEBRoDUwhNA+MqPwHZBaMTcxII40cE3PkCsQSbAK4Cw0MNCEUDD+N/ADoCDAQyF8YA4wGxBtPCwzADMwE2B38Bc1MH8wN+CAURowVTAQYEywizE9QIoykCuATJAJAJc4MHGADzUxMJAvEG+AcTJXoS+AMxCIM7AEMVCRNE9AX3BMwiZQeTE4ADE17bBaM/AMoC04MEkgZKBcabEkOGXwNDiQjbAsOWMAaiAwODM3kCqQLzAqMkCEOAAGOHAMMzM7'
			$g___JsLibGunzip &= '8DdwJzgwjqBcQAU0MLmQGzcxESCSZvFBLGCfDjjwgBFRgVgwNzQ548BPYGs0MjHQn2Bx0E54NTc0EJtRe5BC0Vsw9DA5UAc2sJkAMtCf4IQfMF8gFzFkEFEQFjkyOP9QkSCEkIyhmNAqsCDhciF8HyARcQlgD1AIgVoxN10ALGE9RT9uZXeDss+kyyhpKTppN8VqQTDhfXbgenHgochyAizQ3SxuLHMsYQAsaCxvLHUsZgwsY9DeI9osbD0wICxwPU510ccuUABPU0lUSVZFXwBJTkZJTklUWUXywm8gzm88Y3DFb4ApdFtvXT5sgN0kbD2hACksUQA8cA0AAXADARIDcj0xPOA8bCxlPRALpNM0Czm103IpIOdBy3HpMjuAaTw9bDspe+HTsdkGaWYo8QWw4mnDAYBoPW4sdT1hAAkIdTxpAAl1KWE9AGE8PDF8MSZoQCxoPj49MfIHZiQ9aUABNnzQDj1hADt1PHI7dSs9AHMpZVt1XT1mIWADbn0rK1DXPDxQPTEsc1EAfVPDWwBlLGwscF19QUAucHJvdG/h5i4AZ2V0TmFtZT2XlRVQFsMCIOH0Lm6wAbx9LAwD4eMPAwMDZKDljQsDR68CowJIfTsRHCBuLGg9W6Pbr7gAbj0wO248MjgAODtuKyspc3cAaXRjaCghMCkAe2Nhc2UgbjwAPTE0MzpoLnAAdXNoKFtuKzQAOCw4XSk7YnIQZWFrOwWIMjU1AQdELTE0NCs0MEgwLDkPTjc5CE4yQDU2KzAsNw8lOAI3CSU4MCsxOTIBCHVkZWZhdWx0ADprKCJpbnZhAGxpZCBsaXRlAHJhbDogIituACl9dmFyIG89AGZ1bmN0aW9uCCgpewUKIHQodAQpew3jMz09PXQAOnJldHVyblsAMjU3LHQtMyxUMF2DbjQLDjgADjRVBg41Cw45AA41Bg42qQoONjAADjYGDjcLDqoxAA43Bg44Cw4yAA6qOAYOOQsOMwAOOQYOlDEwiw40gA4xMAYPSnSA5jKGdDY1gQ4xrCwxhDpBBzRHBzZBB5ozSgc2RwdBSTE1SgemOEcHgUkxN0kHMsgdocFJMTksMscdMkcWkjcBSjIzSQczMEYWsjdBSjI3SgfHLDeBSlQzMUkHNMcdN8FKM2g1LDPHHTVIFsFKNJYzSgdHNDfBSjUxSQemNsgswUo1OUkHOMgdocFKNjcsNMcdOUgWycFKODNJBzExiDQBS6w5OYoHyEM4QUsxAGKpCBcxNsceOMFLMYBEUjUHHzE5hxc4QUwxaDYzLOgDMqgxABN0kC0xOTXqAzU35jGaOKEmMgAy5QcyNatXHjgBJ4ACYVARd2VuZ0R0aOF2dCl9AndyACxlLGk9W107AGZvcihyPTM7SHI8PeAHO3Kgk2UAPXQociksaVsAcl09ZVsyXTyAPDI0fGVbMQABRDE2AAEwXTvDeCAgaX0oKTvGfm0oCHQscgZ/dGhpcwouQQwsIgFqPTMyTDc2oDDBAmQ9YgJmHcMAY8MAQKRDBWlucAB1dD1FP25ldwAgVWludDhBckhyYXlgijp0AwRvjD0'
			$g___JsLibGunzip &= 'hAC4BCWs9U0MCQndBAihyfHyAGXsAfSwwKSkmJiiAci5pbmRleCABCcINYz0EAiksci4AYnVmZmVyU2l0emXFA2rAA6cCZwRUdHlwZgRrZgShAmEEcjRlc2gId+ADIwIpKaUDE2tEvE46YhBhySBEYj1AHChFP4ccOilCHSkoYiUrIgZqK/vgMkrBU0UJoFKBINcIAwhXJBHAOqEFROMScYMuQTWjAWyjAUMkzkdIRXIGcgBEJsBpbmZsYQB0ZSBtb2RlIiIpYElFJibFNDMyEeM0byk7QcNOPTAALFM9MTttLnAgcm90b3SgKC5nE8rFAAs7IUIUbzspQnthBnQ9VCgBAiyIMyk75McxJnQFNAXAPjBgGT4+Pj0xaaQqMDrBBnKjF+JGLI3EHGPAW2ILYixuwwMoYSxzoDNsQmEsYYA9VSxoPWkuxAE4bz1VBQ6CBgZUMCwAczw9ZSsxJiaMayhsIQDfb21woD0Ac2VkIGJsb2MAayBoZWFkZXJgOiBMRU6gI4AMckBbZSsrXXzDADy8PDjfBLV13wTWBE7mBHA9PX4oLQWALq8Jbg5jrwmjCQMPIHZlcghpZnlBCmUrYT4ucnQQWg5AOCDzMCBpAHMgYnJva2Vu1wEDAhJnLWZhHG5ABMUUAbAcaWYoYS09bwA9aC1uLEUpaQguc2UQRS5zdWICYZI8ZSxlK28pRCxuIAArPW/AAD0wbztlbHCQsgRvLRAtOylpQJArXT1Fww47IgdhPW41HWWWKKADVB19SJJTOt8IBClpgwJlKHt0OvwyfZU0Ly5Eji0usAwvDBooIQxhQAggACs9YTUhDGEoDGEvDCIMYz2OZfM2sQwSAWI9aVmegjGjPWwoTSxqOj8CMoIudSxmLGMsSGwscHUyNSnQQTdELHn4ADEsYtUANGApKzQsZ09GREZHUVQfKSxk0DB2MAB3tTAAQTAAbeAwkRRtUKsAbTxiOysrbSmAZ1tHW21dXbUFoxA6kBEhRSlzAmJQBypn9AQ7LQMwUmZ1PTh6KGexBk9PUSEpKCBwK3kpLCAHLGwmPaAAUARsO5Wydj1CcYNBdSksdrQ/MUI20h5BPTMrVEMyiCk7QZEXZFttUDd0PXeaFTfdAhAM2QIwTDt30A7IJDE4FAMxOjEFBjcLBikD9SR3PRHUB3Z9ZgAPRT9kQXcvMCxwKTrwAGwQaWNlKMEAKSxjyz0CGAJwNWRsKGAfjysB9Eh1bmtub3duwCBCVFlQRTN+gFnDFHpCJnEoKX0yWfIjYD1bMTYsYaPQrixAOCw3LDks8AAwAixgpjEsNCwxMiYs4KiQoSwxwAAsMSA1XSxHPcd4MTaRU15jKTqQKD1bIJW1kAM2YAQ4gARQsDGgrp8ABNAEkKngp5AENyygkf8goXCfwJ0QnGCaII/wljGVvjHAAbGR8Y8xjuBuLDQAhl2QLq8HbCk6bCAvWlsQtyw5AIAHLDEAMl4sEwBQpqAMAQo0cA0072AKAJZQADIDXQArpwUkhtB5KTp5ADNbsK2QA/sRA9ERMeAMsLAgrbAEwBIINSw5oA0yOSwxVjnADZDMM+CXNVACN5I2MAEwMqAPNTMAA5QwNLCqMPCwNDAAA0A2MTQ1LDghAzEoMjI4gAI2YQMyNDA1NzddADhvDmcp4DpnLHY9Zg6lDWEN6wENwQw2oBs3shfRF9EHNiwkGNACMrEcwQszXQwsQygcxA52KTp2DCx430HUQTI4OClBdDswLGY9eMU9dQQ8ZvBAdSl4W3WIXT11gus/ODqAAAng6T85ggA3OT83BDo4EidELEksTWEwLngpLFqvBqQGM4owkwZEIINJPVqVBghEPEmQBkQpWltQRF09NeIEaqAEWvfAAtByQecggD8hpzFIgYCRMa10LmYRfy5k4H6OdACgYGyAeHQuY/B+AnOl'
			$g___JsLibGunzip &= 'BW48cjspaHA8PWEmeX0fbxNvaSB8PXNbYSBFPDxAbixuKz049a9lAD1pJigxPDxyBCktAKkuZj1pPgA+PnIsdC5kPVRuLXEAY7FhfRCyVLgAY3Rpb24gcSgAdCxyKXtmb3IAKHZhciBlLGkALG49dC5mLHMRAChkLGEAFGlucBB1dCxoACRjLG8APWEubGVuZ3QAaCx1PXJbMF0ELGYADDFdO3M8AGYmJiEobzw9AGgpOylufD1hAFtoKytdPDxzACxzKz04O3JlAHR1cm4gczwoAGk9KGU9dVtuACYoMTw8ZiktADFdKT4+PjE2ACkmJmsoRXJyAQCJImludmFsaYBkIGNvZGUgA28AOiAiK2kpKSxBAJs9bj4+aQAIZAg9cy0BB2M9aCwANjU1MzUmZX0IZnVuA9hCKHQpQHt0aGlzLgK/PUR0LAIMYz0wAwhtCD1bXYMEcz0hMQB9bS5wcm90b4B0eXBlLmw9hSKLA44CjD0CH2IsaQMEhGE7ggdyPXQ7hZ2AbixzLGEsaICRAwAhgpEtMjU4OzIANTYhPT0obj0jALUARyx0KQCQaWYoKG48gAwpAJhpJgQmKAImYT1pLGURAy9lKCmGNCksZQBbaSsrXT1uOyBlbHNlIIE3aD0gcFtzPW4AMDddgCwwPGJbc12AItBoKz1UAiUsAQgAlokGOnIpAOVkW24BFvZDAAMAFmEGCwEEAAuhIkA7aC0tOynAI10iPQIlLWFdwj87OCI8Ay5kOynCMmQtND04g1djQAyFN30s3clZQ/9Zw4bjWG8BnkIibUBXKYSxP1dzwM8hV2m8K2hAXRMhv1WIr2WHVRwpe8GuQPjAlW5ldwAoRT9VaW50OEBBcnJheTqiACmBRE8tMzI3NjiHTifjAWGBQWhiOwBWRSkAZS5zZXQobi5Qc3ViYUEIKIIGLJfFXQBbBlR04GxyPSUDADt0PHI7Kyt0RSBHdOBYW3QrIgddA8EKIjJpLnB1c2hMKGXge+ENbitGByyMRSkADssOaSxpAwiXLA7ADOIKO0ENbltCDTBpK3RdpY9CDWE9aWIELG6LUUSIUcMmcmwsZaOhIRJpgaBkRi/B4ghjKzF8MMGkxgMboEAiA2LmDQA+Im51gG1iZXIiPT1hi8BvZiB0LnTAAoGsqHQpLLADeqEDK6CtgnoBeDwyP2U9YaYqc4QPLYMPKYMQclsgMl0vMirgjHwwFCk8ZbM/BQEraTpBRQE8PDE6Zaa3KoBuLEU/KHI9gEIWIEdC4C8pQjxhKTqIcj1h43diPXIrJe5x8ktBJaAhMIYgIcZCI1MAonpPbitrUCmhQjBQPT09aEQeKYTHRT4/wyvtT6ICQBalA2xpXGNlRFNEA2KUckBSZSGGCjtyPGXgRHIpBcGpaUADbj0odD0QaFtyXUWDO2k8Am4gBGkpb1t'
			$g___JsLibGunzip &= 'zK9HAr3RbaYOdciRIZLW2YUYJBAVhgAhKTmlmzQBidWZmZXI9b30rKkEwKsOpYAsD52Qfd+Y/gBSMNnIpgzbFH4VzojAAkyk6dAMJYocmleECOkMuYhQWPnJAJ3v0AyNVPRBjpQNEObQLdOh9LELYRka5C4QakgRIc3x8cgBnKPQDbdQuc5IZKQsEZwkEwRgPIktTC1k3A3djPHQ7q+g5AixBYEpVsCdVsCeqVWAnVRAnVfCJVcCJSFUsY6gELGzDAGMAO3N3aXRjaCiAaS51PWNbbOAdECxpLnalACgzMYMge7ABfHwxMzmiAAJ274lkIGZpbGVAIHNpZ25hwI1lgjoAii51KyIscQAOdsA7QE3HBXApe2MAYXNlIDg6YnIAZWFrO2RlZmEQdWx0OjaPdW5rEG5vd24wj21wcghlc3PRmW1ldGgEb2Sijy5wKSl9UVBYaS5oRQVlhAB8kRMMPDw4lgAxNqYAFDI0oA1IghJEYXRAZSgxZTMqcFZpRC5OdQRpLk2lADBwPCg0JiAG0A6wEUkDZAG2BDgsbCs9acguSSkwhyg4ggKzo8J1ES5vPTA7AAT1CRFgjnVbb5EzU3RyAGluZy5mcm9toENoYXJDUJ0oYKIgaS5uYW2woC5qwG9pbigiIvINAAQ8MTafBZ8FnwWZBUo9NnVnBTAFMiIFUC5pLgRCPcOidyhjLDCkLGwwby5CMB4oUxQDtw4fH2QgaGVhZABlciBjcmMxNiwiKQFaIBJjVDQtNHZd0APWADPQrxEZBgEydwABghkWATEQAfAZ9gBsAC00LTQ8NTEyhCpuoAphPW4pYC0F4T5tsAp7aW5kZRB4OmwsgzhTaXoQZTphfeELZGF0QGE9cj1zLrE2bIVwAGMgH0s9Zj0+DQfWDQoiAbgwLHcocuAsVSxVKbAQsLx2KRG1uENSQ0CBIGNoAGVja3N1bTogEDB4Iis1Ay50b4VDHiiAvCsiIC/SAeZmagESMUw9cFadFl8JEDw8MjRTCSg0MgA5NDk2NzI5NfwmcoQWIQoAEi4K4kIAOUeAEdDCnwN0aCmyCSKsK27ACNJJbYOHaTRKMGM9bH1yAUDBMDsBkUhwLHksYixn0XNDbSxkkGJ2cC6BSiJwsAB5PWdECTtwBDx54GJwKXYrPeBnW3BdLlEZtQGAOxpFkmZigjgod3YpLB9wBMYDcFqwjvYDLGQp3Cxk/wSCwELbYoE3CgRjoAf3BztiPWKeCFtjAG9uY2F0LmFwAHBseShbXSxiBCl9ZFtifSx0KAAiWmxpYi5HdYBuemlwIixCcBHDKgGIBGRlY2+jSnACDXgBZz8DptVnZXRNamXxkHM7A0Y9A1MCIhwsQY8B4QPpBmV0Tr9gRKACTwHPA88DQFFhzgMHQQHPA88DTXRpbWUB2wNHKX0pLmNhVGxsInYpQiRn4hQgLD0gsTwIFigVFGVkc4MCNxVlZIIbMAPDAy5tpwEovwJgAiLwOHGOIMOC/BCNOyBpPP4EBSkAIGlOsAArKykge2RlYwBvbXByZXNzZUBkICs9ICgJgEEAcnJheVtpXSAAPCAxNiA/ICIAMCIgOiAiIikIICsgEZwudG9TAHRyaW5nKDE2ACk7fTs='
			$g___JsLibGunzip = _B64Decode($g___JsLibGunzip, True, True)
			If @error Then Return SetError(2, __HttpRequest_ErrNotify('__Gzip_Uncompress', 'Khởi tạo thư viện Gzip thất bại'), $sBinaryData)
			$g___JsLibGunzip = BinaryToString($g___JsLibGunzip)
		EndIf
		Local $sRet = _JS_Execute('', 'var compressed=[' & StringTrimLeft(StringRegExpReplace($sBinaryData, '(\w{2})', ',0x$1'), 6) & '];' & $g___JsLibGunzip, 'decompressed')
		If @error Or $sRet = '' Then Return SetError(3, __HttpRequest_ErrNotify('__Gzip_Uncompress', 'Giải mã Gzip thất bại'), $sBinaryData)
		Return StringStripWS($sRet, 3)
	EndFunc

	Func __ArrayDuplicate($aArray, $iCase = True, $iCount = False, $iCheckDulplicateNumber = False, $iBase = 0)
		If Not IsArray($aArray) Or UBound($aArray) < $iBase Then Return SetError(1, 0, $aArray)
		Local $oDictionary = ObjCreate("Scripting.Dictionary")
		With $oDictionary
			.CompareMode = Number(Not $iCase)
			If $iCheckDulplicateNumber = False Then
				For $i = $iBase To UBound($aArray) - 1
					.Item($aArray[$i])
				Next
				$aArray = .Keys
				If $iCount Then _ArrayInsert($aArray, 0, $oDictionary.Count)
				$oDictionary = Null
				Return $aArray
			Else
				For $i = $iBase To UBound($aArray) - 1
					If .Exists($aArray[$i]) Then
						.Item($aArray[$i]) = .Item($aArray[$i]) + 1
					Else
						.Add($aArray[$i], 1)
					EndIf
				Next
				Local $aArray2 = [.Keys, .Items]
				If $iCount Then
					_ArrayInsert($aArray2[0], 0, $oDictionary.Count)
					_ArrayInsert($aArray2[1], 0, $oDictionary.Count)
				EndIf
				$oDictionary = Null
				Return $aArray2
			EndIf
		EndWith
	EndFunc

	Func __ObjectErrDetect()
		If $g___oErrorStop = 0 Then
			Local $sReport = StringReplace('<Error> COM Error (Line ' & $g___oError.scriptline & ') ' & $g___oError.source & ': ' & $g___oError.windescription & ' ' & $g___oError.description, @CRLF, ' ')
			_HttpRequest_ConsoleWrite(@CRLF & $sReport & @CRLF)
		EndIf
		$g___oErrorStop = 0
		Return SetError($g___oError.scriptline)
	EndFunc

	Func __HttpRequest_CancelReadWrite()
		$g___CancelReadWrite = Not $g___CancelReadWrite
	EndFunc

	Func __Data2Send_CheckEncode($sData2Send)
		Local $aPartData = StringRegExp($sData2Send, '(?:^|\&)(\w+)\h*?=\h*?([^\&]+)', 3)
		For $i = 1 To UBound($aPartData) - 1 Step 2
			If Not StringRegExp($aPartData[$i], '\%[0-9A-Z]') And StringRegExp($aPartData[$i], '[^\w\-\+\.\~]') Then
				;$sData2Send = StringReplace($sData2Send, $aPartData[$i - 1] & $aPartData[$i], $aPartData[$i - 1] & _URIEncode($aPartData[$i]), 1, 1)
				__HttpRequest_ErrNotify('__Data2Send_CheckEncode', 'Giá trị của Key "' & $aPartData[$i - 1] & '" trong POST data của _HttpRequest chưa Encode, điều đó có thể khiến request thất bại' & @CRLF, '', 'Warning')
			EndIf
		Next
	EndFunc

	Func __HttpRequest_CloseAll()
		ConsoleWrite(@CRLF)
		Local $aListSession = _HttpRequest_SessionList()
		If Not @error Then
			For $i = 0 To UBound($aListSession) - 1
				If $g___hRequest[$i] Then $g___hRequest[$i] = _WinHttpCloseHandle2($g___hRequest[$i])
				If $g___hWebSocket[$i] Then $g___hWebSocket[$i] = _WinHttpWebSocketClose2($g___hWebSocket[$i])
				If $g___hConnect[$i] Then $g___hConnect[$i] = _WinHttpCloseHandle2($g___hConnect[$i])
				_HttpRequest_ClearSession($aListSession[$i])
			Next
		EndIf
		;---------------------------------------------------------------------------
		If $g___hStatusCallback Then DllCallbackFree($g___hStatusCallback)
		If $dll_WinInet Then $dll_WinInet = DllClose($dll_WinInet)
		If $dll_Gdi32 Then $dll_Gdi32 = DllClose($dll_Gdi32)
		$dll_WinHttp = DllClose($dll_WinHttp)
		$dll_User32 = DllClose($dll_User32)
		$g___oDicEntity = Null
		$g___retData = Null
		$g___oError = Null
		;---------------------------------------------------------------------------
		If $g___CookieJarPath Then _HttpRequest_CookieJarUpdateToFile()
	EndFunc

	Func __HttpRequest_ErrNotify($__TrueValue = '', $__ErrorNote = '', $__FalseValue = '', $iTypeWarning = Default)
		If @Compiled Or $g___OldConsole = $__ErrorNote Then Return
		$g___OldConsole = $__ErrorNote
		If $g___ErrorNotify = True And $__ErrorNote Then
			If $iTypeWarning = Default Then $iTypeWarning = 'Error'
			_HttpRequest_ConsoleWrite('<' & $iTypeWarning & '> [#' & $g___LastSession & '] ' & $__TrueValue & ' : ' & $__ErrorNote & @CRLF)
		EndIf
		Return $__FalseValue
	EndFunc

	Func __HttpRequest_CheckUpdate($iCurrentVersion)
		;http://jsoneditoronline.org/?id
		If $CmdLine[0] > 0 And $CmdLine[1] = '--httprequest-update' Then
			TraySetState(2)
			Local $UpdateInfo = BinaryToString(InetRead('http://api.jsoneditoronline.org/v1/docs/39cf9a61c45c466880a2e4899bc293be'))
			Local $sVersionHR = StringRegExp($UpdateInfo, 'version=(\d+)', 1)
			If Not @error And Number($sVersionHR[0]) > $iCurrentVersion Then
				If MsgBox(64 + 4096 + 4, 'Thông báo', '_HttpRequest có bản cập nhật mới (ver.' & $sVersionHR[0] & '). Bạn có muốn tải về ngay ?') = 6 Then
					Local $LinkDownload = StringRegExp($UpdateInfo, '(?i)linkdownload=\[([^\]]*?)\]', 1)
					If @error Then MsgBox(16 + 4096, 'Thông báo', 'Có lỗi trong khi thực hiện Update')
					If $LinkDownload[0] = '' Then Exit
					ShellExecute($LinkDownload[0])
					MsgBox(16 + 4096, 'Thông báo', 'Vui lòng xem ChangeLog phiên bản ' & $sVersionHR[0] & ' trong tập tin Help để xem thông tin thay đổi cụ thể.')
				Else
					MsgBox(64 + 4096, 'Thông báo', 'Thông báo cập nhật sẽ hiển thị lại sau nửa tiếng')
				EndIf
			EndIf
			Exit
		Else
			If @Compiled Or ($CmdLine[0] > 0 And $CmdLine[1] = '--hh-multi-process') Then Return
			Local $TimeInit = Number(RegRead('HKCU\Software\AutoIt v3\HttpRequest\AutoUpdate', 'Timer'))
			If Not $TimeInit Or TimerDiff($TimeInit) > 30 * 60 * 1000 Then
				RegWrite('HKCU\Software\AutoIt v3\HttpRequest\AutoUpdate', 'Timer', 'REG_SZ', TimerInit())
				Run(FileGetShortName(@AutoItExe) & ' "' & @ScriptFullPath & '" --httprequest-update', @WorkingDir, @SW_HIDE)
			EndIf
		EndIf
	EndFunc

	Func _HttpRequest_DetectMIME_Ex($sFilePath)
		Local $aMimeFromData = DllCall("urlmon.dll", "long", "FindMimeFromData", "ptr", 0, 'wstr', $sFilePath, "ptr", 0, 'dword', 0, "ptr", 0, 'dword', 1, "ptr*", 0, 'dword', 0)
		If @error Then
			Return SetError(1, 0, 'application/octet-stream')
		Else
			Local $aStrlenW = DllCall("kernel32.dll", "int", "lstrlenW", "struct*", $aMimeFromData[7])
			If @error Then Return SetError(2, 0, 'application/octet-stream')
			$aMimeFromData = DllStructGetData(DllStructCreate("wchar[" & $aStrlenW[0] & "]", $aMimeFromData[7]), 1)
			If $aMimeFromData = 0 Then
				Return SetError(3, 0, 'application/octet-stream')
			EndIf
			Return $aMimeFromData
		EndIf
	EndFunc

	Func _HttpRequest_DetectMIME($sFileName_Or_FilePath)
		If $g___MIMEData = '' Then
			$g___MIMEData &= ';ai|1/postscript;aif|2/x-aiff;aifc|2/x-aiff;aiff|2/x-aiff;asc|3/plain;atom|1/atom+xml;au|2/basic;avi|5/x-msvideo;bcpio|'
			$g___MIMEData &= '4/bmp;cdf|1/x-netcdf;cgm|4/cgm;class|1/7/;cpio|1/x-bcpio;bin|1/7/;bmp|5/x-dv;dir|1/x-director;djv|4/vnd.djvu;djvu|'
			$g___MIMEData &= '1/x-cpio;cpt|1/mac-compactpro;csh|1/x-csh;css|3/css;dcr|1/x-director;dif|4/vnd.djvu;dll|1/7/;dmg||1/msword;dtd|'
			$g___MIMEData &= '3/x-setext;exe|1/7/;ez|1/andrew-inset;gif|4/gif;gram|2/midi;latex|1/x-latex;lha|1/7/;lzh|1/7/;m3u|2/mp4a-latm;m4b|'
			$g___MIMEData &= '3/calendar;ief|4/ief;ifb|3/calendar;iges|6/iges;igs|6/iges;jnlp|1/x-java-jnlp-file;jp2|1/x-sv4cpio;sv4crc|1/x-sv4crc;svg|'
			$g___MIMEData &= '3/vnd.wap.wmlscript;wmlsc|1/vnd.wap.wmlscriptc;wrl|6/vrml;xbm|4/svg+xml;swf|1/x-shockwave-flash;t|1/x-koan;skt|'
			$g___MIMEData &= '4/pict;pict|4/pict;png|4/png;pnm|4/x-portable-anymap;pnt|4/x-macpaint;pntg|2/x-pn-realaudio;ras|4/x-cmu-raster;rdf|'
			$g___MIMEData &= '4/x-macpaint;ppm|4/x-portable-pixmap;ppt|1/vnd.ms-powerpoint;ps|1/postscript;qt|1/rdf+xml;rgb|1/x-futuresplash;src|'
			$g___MIMEData &= '5/quicktime;qti|4/x-quicktime;qtif|4/x-quicktime;ra|2/x-pn-realaudio;ram|1/vnd.rn-realmedia;roff|1/x-troff;rtf|3/rtf;rtx|'
			$g___MIMEData &= '3/sgml;sh|1/x-sh;shar|1/x-shar;silo|6/mesh;sit|1/x-stuffit;skd|1/x-tcl;tex|1/x-tex;texi|1/x-texinfo;texinfo|1/x-texinfo;tif|'
			$g___MIMEData &= '4/tiff;tiff|4/tiff;tr|1/x-troff;tsv|3/tab-separated-values;txt|3/plain;ustar|1/smil;snd|2/basic;so|1/x-ustar;vcd|6/vrml;vxml|'
			$g___MIMEData &= '4/vnd.wap.wbmp;wbmxl|1/vnd.wap.wbxml;wml|3/vnd.wap.wml;wmlc|1/vnd.wap.wmlc;wmls|1/7/;spl|1/x-cdlink;vrml|'
			$g___MIMEData &= '4/x-xbitmap;xht|1/xhtml+xml;xhtml|1/xhtml+xml;xls|1/vnd.ms-excel;xml|1/voicexml+xml;wav|1/x-koan;skm|1/xml;xpm|'
			$g___MIMEData &= '1/xml-dtd;dv|5/x-dv;dvi|1/x-dvi;dxr|1/x-director;eps|1/postscript;etx|1/7/;dms|1/7/;doc|1/x-gtar;hdf|1/x-hdf;hqx|'
			$g___MIMEData &= '1/mac-binhex40;htm|3/html;html|3/html;ice|x-conference/x-cooltalk;ico|4/x-icon;ics|1/srgs;grxml|1/srgs+xml;gtar|'
			$g___MIMEData &= '4/jp2;jpe|4/jpeg;jpeg|4/jpeg;jpg|4/jpeg;js|1/x-javascript;kar|1/x-wais-source;sv4cpio|3/richtext;sgm|3/sgml;sgml|'
			$g___MIMEData &= '2/x-mpegurl;m4a|2/mp4a-latm;m4p|2/mp4a-latm;m4u|5/vnd.mpegurl;m4v|1/x-troff;tar|1/x-tar;tcl|2/x-wav;wbmp|'
			$g___MIMEData &= '5/x-m4v;mac|4/x-macpaint;man|1/x-troff-man;mathml|1/mathml+xml;me|1/xslt+xml;xul|1/vnd.mozilla.xul+xml;xwd|'
			$g___MIMEData &= '1/x-troff-me;mesh|6/mesh;mid|2/midi;midi|2/midi;mif|1/vnd.mif;mov|4/x-portable-graymap;pgn|1/x-chess-pgn;pic|'
			$g___MIMEData &= '5/quicktime;movie|5/x-sgi-movie;mp2|2/mpeg;mp3|2/mpeg;mp4|5/mp4;mpe|5/mpeg;mpeg|4/x-xwindowdump;xyz|'
			$g___MIMEData &= '5/mpeg;mpg|5/mpeg;mpga|2/mpeg;ms|1/x-troff-ms;msh|6/mesh;mxu|5/vnd.mpegurl;nc|1/x-koan;smi|1/smil;smil|'
			$g___MIMEData &= '1/x-netcdf;oda|1/oda;ogg|1/ogg;pbm|4/x-portable-bitmap;pct|4/pict;pdb|chemical/x-pdb;pdf|1/pdf;pgm|4/x-rgb;rm|'
			$g___MIMEData &= '1/x-koan;skp|4/x-xpixmap;xsl|1/xml;xslt|chemical/x-xyz;zip|1/zip;xlsx|1/vnd.openxmlformats-officedocument.spread'
			$g___MIMEData &= 'sheetml.sheet;doc|1/msword;dot|1/msword;docx|1/vnd.openxmlformats-officedocument.wordprocessingml.document;dotx|'
			$g___MIMEData &= '1/vnd.openxmlformats-officedocument.wordprocessingml.template;docm|1/vnd.ms-word.document.macroEnabled.12;dotm|'
			$g___MIMEData &= '1/vnd.ms-word.template.macroEnabled.12;xls|1/vnd.ms-excel;xlt|1/vnd.ms-excel;xla|1/vnd.ms-excel;xltx|1/vnd.openxml'
			$g___MIMEData &= 'formats-officedocument.spreadsheetml.template;xlsm|1/vnd.ms-excel.sheet.macroEnabled.12;xltm|1/vnd.ms-excel.template.'
			$g___MIMEData &= 'macroEnabled.12;xlam|1/vnd.ms-excel.addin.macroEnabled.12;xlsb|1/vnd.ms-excel.sheet.binary.macroEnabled.12;ppt|'
			$g___MIMEData &= '1/vnd.ms-powerpoint;pot|1/vnd.ms-powerpoint;pps|1/vnd.ms-powerpoint;ppa|1/vnd.ms-powerpoint;pptx|1/vnd.openxmlfor'
			$g___MIMEData &= 'mats-officedocument.presentationml.presentation;potx|1/vnd.openxmlformats-officedocument.presentationml.template;ppsx|'
			$g___MIMEData &= '1/vnd.openxmlformats-officedocument.presentationml.slideshow;ppam|1/vnd.ms-powerpoint.addin.macroEnabled.12;pptm|'
			$g___MIMEData &= '1/vnd.ms-powerpoint.presentation.macroEnabled.12;potm|1/vnd.ms-powerpoint.template.macroEnabled.12;ppsm|'
			$g___MIMEData &= '1/vnd.ms-powerpoint.slideshow.macroEnabled.12'
			;-----------------------------------------------------------------------------------------------------
			Local $aMshort = ['', 'application', 'audio', 'text', 'image', 'video', 'model', 'octet-stream']
			For $i = 1 To 7
				$g___MIMEData = StringReplace($g___MIMEData, $i & '/', $aMshort[$i] & '/', 0, 1)
			Next
		EndIf
		;-----------------------------------------------------------------------------------------------------
		Local $aArray = StringRegExp($g___MIMEData, "(?i)\Q;" & StringRegExpReplace($sFileName_Or_FilePath, "(.*?)\.(\w+)$", "$2") & "\E\|(.*?);", 1)
		If @error Then
			If FileExists($sFileName_Or_FilePath) Then
				Local $fOpen = FileOpen($sFileName_Or_FilePath)
				Switch FileRead($fOpen, 4)
					Case 'ÿØÿà'
						Return 'image/jpg'
					Case '‰PNG'
						Return 'image/png'
					Case 'BMN'
						Return 'image/bmp'
				EndSwitch
				FileClose($fOpen)
			EndIf
			Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_DetectMIME', 'Không thể tra MIME của loại tập tin này. MIME sẽ được trả về mặc định là: application/octet-stream'), 'application/octet-stream')
		Else
			Return $aArray[0]
		EndIf
	EndFunc

	Func __HttpRequest_StatusGetDataFromPointer($pInfo, $lInfo, $iReturnType = 'wchar')
		Return DllStructGetData(DllStructCreate($iReturnType & '[' & $lInfo & ']', $pInfo), 1)
	EndFunc

	Func __HttpRequest_StatusCallback($hInternet, $iContext, $iInternetStatus, $pStatusInfo, $iStatusInfoLen)
		#forceref $hInternet, $iContext, $iInternetStatus, $pStatusInfo, $iStatusInfoLen
		Switch $iInternetStatus
			Case 0x00000002 ;CALLBACK_STATUS_NAME_RESOLVED
				$g___ServerIP = __HttpRequest_StatusGetDataFromPointer($pStatusInfo, $iStatusInfoLen)
				Return
				;----------------------------------------------------------------------------------------------------
			Case 0x00004000 ;CALLBACK_STATUS_REDIRECT
				$g___LocationRedirect = DllStructGetData(DllStructCreate("wchar[" & $iStatusInfoLen & "]", $pStatusInfo), 1)
				$g___retData[$g___LastSession][0] &= __CookieJar_Insert(StringRegExpReplace($g___LocationRedirect, '(?i)^https?://([^/]+).+', '${1}', 1), _WinHttpQueryHeaders2($g___hRequest[$g___LastSession], 22)) & _
						@CRLF & 'Redirect → [' & $g___LocationRedirect & ']' & @CRLF
				_HttpRequest_ConsoleWrite('> [#' & $g___LastSession & '] Request đã redirect tới: ' & $g___LocationRedirect & @CRLF)
				Return
				;----------------------------------------------------------------------------------------------------
			Case 0x00010000 ;CALLBACK_STATUS_SECURE_FAILURE
				Local $sStatus = ''
				Local $aSSLError = [ _
						[__HttpRequest_StatusGetDataFromPointer($pStatusInfo, $iStatusInfoLen, 'dword')], _
						[0x00000001, 'CERT_REV_FAILED'], _
						[0x00000002, 'INVALID_CERT'], _
						[0x00000004, 'CERT_REVOKED'], _
						[0x00000008, 'INVALID_CA'], _
						[0x00000010, 'CERT_CN_INVALID'], _
						[0x00000020, 'CERT_DATE_INVALID'], _
						[0x00000040, 'CERT_WRONG_USAGE'], _
						[0x80000000, 'SECURITY_CHANNEL_ERROR']]
				For $i = 1 To 8
					If BitAND($aSSLError[0][0], $aSSLError[$i][0]) = $aSSLError[$i][0] Then $sStatus &= ' ' & $aSSLError[$i][1]
				Next
				_HttpRequest_ConsoleWrite('<Error> [#' & $g___LastSession & '] SLL Certificate:' & $sStatus & ' - Kiểm tra lại URL là http hay https' & @CRLF)
				Return
		EndSwitch
	EndFunc

	Func __HttpRequest_iReturnSplit($iReturn)
		Local $aRetMode[20]
		$aRetMode[11] = 4
		$aRetMode[8] = $g___LastSession
		;-------------------------------------------------
		For $iReturn In StringSplit($iReturn, '|', 2)
			If $iReturn == '' Then ContinueLoop
			Local $iLocalMode = StringRegExp($iReturn, '^\h*?([\+\-\*\.\_\~\^]*?)(\d+):?(\d{0,2})', 3)
			If Not @error Then
				$aRetMode[0] = Number($iLocalMode[1]) ;Number Return Mode
				If $iLocalMode[2] Then ;Query Header Mode
					$aRetMode[0] = 1
					$aRetMode[1] = Number($iLocalMode[2])
				EndIf
				If $iLocalMode[0] Then
					For $iLocalMode In StringSplit($iLocalMode[0], '', 2)
						Switch $iLocalMode
							Case '-' ;$iReturn = 1 => Return Cookies, $iReturn > 1 => Return Binary Data
								$aRetMode[2] = 1
							Case '*' ;force Disable Redirect
								$aRetMode[3] = 1
							Case '+' ;Complete URL for relative URL
								$aRetMode[4] = 1
							Case '~' ;force return ANSI
								$aRetMode[11] = 1
							Case '_' ; force return Raw Text
								$aRetMode[12] = 1
							Case '^' ; force WebSocket
								$aRetMode[13] = 1
							Case '.' ;chỉ gửi request đi và không làm gì tiếp cả
								$aRetMode[14] = 1
							Case Else
								Return SetError(1, __HttpRequest_ErrNotify('__HttpRequest_iReturnSplit', 'Không nhận ra dấu hiệu đã cài đặt'), '')
						EndSwitch
					Next
				EndIf
			Else
				;--------------------------------------------------------------------------------------------------------------------------------
				Local $iLocalOption = StringRegExp($iReturn, '^\h*?([\$\#\%])(.+)', 3)
				If Not @error Then
					For $i = 0 To UBound($iLocalOption) - 1 Step 2
						Switch $iLocalOption[$i]
							Case '%' ;proxy ; $aRetMode[5][6][7]
								Local $aProxy = StringRegExp($iLocalOption[$i + 1], '(?i)(https?://)?(?:(\w*):)?(?:(\w*)@)?((?:\d{1,3}\.){3}\d{1,3}:\d+)$', 3)
								If @error Then Return SetError(2, __HttpRequest_ErrNotify('__HttpRequest_iReturnSplit', 'Sai pattern của Proxy'), '')
								$aRetMode[5] = $aProxy[0] & $aProxy[3]
								$aRetMode[6] = $aProxy[1]
								$aRetMode[7] = $aProxy[2]
							Case '#' ;session ; $aRetMode[8]
								$g___hCookieLast = ''
								If Not StringIsDigit($iLocalOption[$i + 1]) Then Return SetError(4, __HttpRequest_ErrNotify('__HttpRequest_iReturnSplit', 'Sai pattern của Session'), '')
								If $iLocalOption[$i + 1] > $g___MaxSession_USE Then Return SetError(4, __HttpRequest_ErrNotify('__HttpRequest_iReturnSplit', 'Session vượt quá giới hạn. Max=' & $g___MaxSession_USE), '')
								$aRetMode[8] = Number($iLocalOption[$i + 1])
							Case '$' ;file path ; $aRetMode[9][10]
								Local $aPath = StringRegExp($iLocalOption[$i + 1], '(?i)^([A-Z]:[\\\/].*?)(?:\:(\d+))?($)', 3)
								If @error Then Return SetError(5, __HttpRequest_ErrNotify('__HttpRequest_iReturnSplit', 'Sai pattern của FilePath'), '')
								$aRetMode[0] = 3
								$aRetMode[9] = $aPath[0]
								$aRetMode[10] = Number($aPath[1])
						EndSwitch
					Next
				EndIf
			EndIf
		Next
		Return $aRetMode
	EndFunc

	Func __HttpRequest_URLSplit($sURL)
		Local $aResult[10] = [1, 80, '', '', '', '', '', 'http', '', '']
		;---------------------------------------------------
		Local $aURL1 = StringRegExp($sURL, '(?i)^\h*(?:(?:(https?)|(ftp)|(wss?)):/{2,})?(www\.)?(.*?)\h*$', 3)
		If @error Or Not $aURL1[4] Then Return SetError(1, __HttpRequest_ErrNotify('__HttpRequest_URLSplit', '$sURL sai định dạng chuẩn #1'), '')
		If $aURL1[1] Then ; Check ftp
			$aResult[0] = 3
			$aResult[1] = 0
		ElseIf $aURL1[2] Then
			If Not StringRegExp(@OSVersion, '^WIN_(10|81|8)$') Then Return SetError(2, __HttpRequest_ErrNotify('__HttpRequest_URLSplit', 'Websock chỉ áp dụng cho Win8 and Win10'), '')
			$aResult[8] = 1
			If $aURL1[2] = 'wss' Then
				$aResult[0] = 2
				$aResult[1] = 443
				$aResult[7] = 'https'
			EndIf
		ElseIf $aURL1[0] = 'https' Then ;Check https
			$aResult[0] = 2
			$aResult[1] = 443
			$aResult[7] = 'https'
		EndIf
		; Chưa xác định được protocol thì sẽ check $aURL3[0] bên dưới
		;---------------------------------------------------
		Local $aURL2 = StringRegExp($aURL1[4], '^(?:(\w+):(\w+)@)?(.+)$', 3) ;Tách user, pass, cred, URL
		If @error Then Return SetError(4, __HttpRequest_ErrNotify('__HttpRequest_URLSplit', '$sURL sai định dạng User/Pass trong URL'), '')
		$aResult[4] = $aURL2[0] ;User
		$aResult[5] = $aURL2[1] ;Pass
		;---------------------------------------------------
		Local $aURL3 = StringRegExp($aURL2[2], '^([^\/\:]+)(?::(\d+))?(/.*)?($)', 3) ;Tách Host, (Port) và URI
		If @error Then Return SetError(5, __HttpRequest_ErrNotify('__HttpRequest_URLSplit', '$sURL sai định dạng Host/Port trong URL'), '')
		If $aURL1[0] == '' And Not (StringRegExp($aURL3[0], '\.\w+$') Or $aURL3[0] = 'localhost') Then Return SetError(3, __HttpRequest_ErrNotify('__HttpRequest_URLSplit', '$sURL sai định dạng chuẩn #2'), '')
		$aResult[2] = StringRegExpReplace($aURL1[3] & $aURL3[0], '(\#[\w\-]+)$', '', 1) ;Host
		$aResult[3] = $aURL3[2] ;URI
		If $aURL3[1] Then $aResult[1] = Number($aURL3[1]) ; Check Port
		;---------------------------------------------------
		$aResult[9] = StringRegExpReplace($aResult[2], '.*?([\w\-]*?\.?[\w\-]+\.[\w\-]+)$', '$1') ;Domain
		;---------------------------------------------------
		Return $aResult
	EndFunc

	Func _HttpRequest_ParseCURL($iReturn)
		Local $iURL = '', $iHeaders = '', $iData, $iProxy = '', $iAuth = '', $iAuthBK = '', $iMethod = 'GET'
		Local $aURL = StringRegExp($iReturn, '(?i)(?![''"])\h+(?![''"])(https?://[^\s]+|localhost:?\d+?)(?:\h|$)', 1)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_ParseCURL', 'Không thể parse URL từ chuỗi CURL đã nạp vào'), '')
		$iURL = $aURL[0]
		;---------------------------------------------------------
		Local $aHeaders = StringRegExp($iReturn, '(?i) -H\h+([''"])(.+?(?!\\))\1', 3)
		If Not @error Then
			For $i = 1 To UBound($aHeaders) - 1 Step 2
				$iHeaders &= $aHeaders[$i] & ($i < UBound($aHeaders) - 1 ? '|' : '')
			Next
		EndIf
		;---------------------------------------------------------
		Local $aData = StringRegExp($iReturn, '(?:--data(?:-urlencode)?|-d)\h+([''"])(.+?(?!\\))\1(?:\h|$)', 3)
		If Not @error Then
			For $i = 1 To UBound($aData) - 1 Step 2
				$iData &= $aData[$i] & ($i < UBound($aData) - 1 ? '&' : '')
			Next
			$iMethod = 'POST'
		Else
			Local $aData = StringRegExp($iReturn, '(?:--form|-F)\h+([''"])(.+?(?!\\))\1(?:\h|$)', 3)
			If Not @error Then
				Local $iData[UBound($aData) / 2], $iCount = 0
				For $i = 1 To UBound($aData) - 1 Step 2
					$iData[$iCount] = $aData[$i]
					$iCount += 1
				Next
				$iMethod = 'POST'
			EndIf
		EndIf
		;---------------------------------------------------------
		Local $aMethod = StringRegExp($iReturn, ' -X\h+(GET|POST|PUT|HEAD|DELETE|CONNECT|OPTIONS|TRACE|PATCH)(?:\h|$)', 1)
		If Not @error Then $iMethod = $aMethod[0]
		;---------------------------------------------------------
		Local $aUserAgent = StringRegExp($iReturn, ' -A\h+([''"])(.+?(?!\\))\1(?:\h|$)', 1)
		If Not @error Then $iHeaders = ($iHeaders ? '|' : '') & 'User-Agent: ' & $aUserAgent[1]
		;---------------------------------------------------------
		Local $aAuth = StringRegExp($iReturn, ' -u\h+([''"])?(.+?(?!\\))\1?(?:\h|$)', 1)
		If Not @error Then $iAuth = $aAuth[1]
		;----------------------------------------------------------------------------------------------------------------------------------------------
		Local $aProxy = StringRegExp($iReturn, ' (?:--proxy|-x)\h+([''"])?(.+?(?!\\))\1?(?:\h|$)', 1)
		If Not @error Then $iProxy = $aProxy[1]
		;----------------------------------------------------------------------------------------------------------------------------------------------
		If $iAuth Then $iAuthBK = _HttpRequest_SetAuthorization($iAuth)
		Local $vData = _HttpRequest(2 & ($iProxy ? '|' & $iProxy : ''), $iURL, $iData, '', '', $iHeaders, $iMethod)
		Local $vError = @error, $vExtended = @extended
		If $iAuth Then _HttpRequest_SetAuthorization($iAuthBK)
		Return SetError($vError, $vExtended, $vData)
	EndFunc
#EndRegion



#Region < FTP UDF>
	Func __FTP_MakeQWord($iLoDWORD, $iHiDWORD)
		Local $tInt64 = DllStructCreate("uint64")
		Local $tDwords = DllStructCreate("dword;dword", DllStructGetPtr($tInt64))
		DllStructSetData($tDwords, 1, $iLoDWORD)
		DllStructSetData($tDwords, 2, $iHiDWORD)
		Return DllStructGetData($tInt64, 1)
	EndFunc

	Func _FTP_Open2($sAgent, $iAccessType, $sProxyName = '', $sProxyBypass = '', $iFlags = 0) ;$iAccessType = 1: No Proxy; 3: Proxy
		Local $ai_InternetOpen = DllCall($dll_WinInet, 'handle', 'InternetOpenW', 'wstr', $sAgent, 'dword', $iAccessType, 'wstr', $sProxyName, 'wstr', $sProxyBypass, 'dword', $iFlags)
		If @error Or $ai_InternetOpen[0] = 0 Then Return SetError(1)
		Return $ai_InternetOpen[0]
	EndFunc

	Func _FTP_Connect2($hInternetSession, $sServerName, $sUserName, $sPassword, $iServerPort = 0)
		Local $ai_InternetConnect = DllCall($dll_WinInet, 'hwnd', 'InternetConnectW', 'handle', $hInternetSession, 'wstr', $sServerName, 'ushort', $iServerPort, 'wstr', $sUserName, 'wstr', $sPassword, 'dword', 1, 'dword', 2, 'dword_ptr', 0)
		If @error Or $ai_InternetConnect[0] = 0 Then Return SetError(1)
		Return $ai_InternetConnect[0]
	EndFunc

	Func _FTP_CloseHandle2($hSession)
		DllCall($dll_WinInet, 'bool', 'InternetCloseHandle', 'handle', $hSession)
	EndFunc

	Func _FTP_FileReadEx($hFTPConnect, $sRemoteFile, $CallBackFunc_Progress = '', $iBytesPerLoop = $g___BytesPerLoop)
		Local $ai_FtpOpenfile = DllCall($dll_WinInet, 'handle', 'FtpOpenFileW', 'handle', $hFTPConnect, 'wstr', $sRemoteFile, 'dword', 0x80000000, 'dword', 2, 'dword_ptr', 0) ;2 = Binarry, 1 = ascii
		If @error Or $ai_FtpOpenfile[0] == 0 Then Return SetError(1, '', '')
		Local $tBuffer = DllStructCreate("byte[" & $iBytesPerLoop & "]")
		Local $vBinaryData = Binary(''), $vNowSizeBytes = 1, $vTotalSizeBytes = -1, $iCheckCallbackFunc = 0, $aCall
		;----------------------------------
		If $CallBackFunc_Progress <> '' Then
			$iCheckCallbackFunc = 1
			Local $ai_hSize = DllCall($dll_WinInet, 'dword', 'FtpGetFileSize', 'handle', $ai_FtpOpenfile[0], 'dword*', 0)
			If @error Or $ai_hSize[0] = 0 Then Return SetError(103, __HttpRequest_ErrNotify('_FTP_FileReadEx', 'Không thể lấy được kích cỡ tập tin'), 0)
			$vTotalSizeBytes = __FTP_MakeQWord($ai_hSize[0], $ai_hSize[2])
			If $vTotalSizeBytes > 2147483647 Then Return SetError(102, __HttpRequest_ErrNotify('_FTP_FileReadEx', 'Tập tin quá lớn'), 0)
		EndIf
		;----------------------------------
		For $i = 1 To 2147483647
			If $g___CancelReadWrite Then
				$g___CancelReadWrite = False
				Return SetError(997, __HttpRequest_ErrNotify('_FTP_FileReadEx', 'Đã huỷ request'), 0)
			EndIf
			$aCall = DllCall($dll_WinInet, 'bool', 'InternetReadFile', 'handle', $ai_FtpOpenfile[0], 'struct*', $tBuffer, 'dword', $iBytesPerLoop, 'dword*', 0)
			If @error Or $aCall[0] = 0 Or ($aCall[0] = 1 And $aCall[4] = 0) Then ExitLoop
			If $aCall[4] < $iBytesPerLoop Then
				$vBinaryData &= BinaryMid(DllStructGetData($tBuffer, 1), 1, $aCall[4])
			Else
				$vBinaryData &= DllStructGetData($tBuffer, 1)
			EndIf
			$vNowSizeBytes += $aCall[4]
			If $iCheckCallbackFunc Then $CallBackFunc_Progress($vNowSizeBytes, $vTotalSizeBytes)
		Next
		DllCall($dll_WinInet, 'bool', 'InternetCloseHandle', 'handle', $ai_FtpOpenfile[0])
		Return $vBinaryData
	EndFunc

	Func _FTP_FileWriteEx($hFTPConnect, $sRemoteFile, $iData, $CallBackFunc_Progress = '', $iBytesPerLoop = $g___BytesPerLoop)
		Local $ai_FtpOpenfile = DllCall($dll_WinInet, 'handle', 'FtpOpenFileW', 'handle', $hFTPConnect, 'wstr', $sRemoteFile, 'dword', 0x40000000, 'dword', 2, 'dword_ptr', 0) ;2 = Binarry, 1 = ascii
		If @error Or $ai_FtpOpenfile[0] = 0 Then Return SetError(1, '', '')
		If Not IsBinary($iData) Then $iData = StringToBinary($iData, 4)
		Local $vNowSizeBytes = 1, $vTotalSizeBytes = -1, $iCheckCallbackFunc = 0
		Local $iDataMid, $iDataMidLen, $tBuffer, $aCall
		;----------------------------------
		If $CallBackFunc_Progress <> '' Then
			$iCheckCallbackFunc = 1
			$vTotalSizeBytes = BinaryLen($iData)
			If $vTotalSizeBytes > 2147483647 Then Return SetError(101, __HttpRequest_ErrNotify('_FTP_FileWriteEx', 'Tập tin quá lớn'), 0)
		EndIf
		;----------------------------------
		For $i = 1 To 2147483647
			If $g___CancelReadWrite Then
				$g___CancelReadWrite = False
				Return SetError(996, __HttpRequest_ErrNotify('_FTP_FileWriteEx', 'Đã huỷ request'), 0)
			EndIf
			$iDataMid = BinaryMid($iData, $vNowSizeBytes, $iBytesPerLoop)
			$iDataMidLen = BinaryLen($iDataMid)
			If Not $iDataMidLen Then ExitLoop
			$tBuffer = DllStructCreate("byte[" & ($iDataMidLen + 1) & "]")
			DllStructSetData($tBuffer, 1, $iDataMid)
			$aCall = DllCall($dll_WinInet, 'bool', 'InternetWriteFile', 'handle', $ai_FtpOpenfile[0], 'struct*', $tBuffer, 'dword', $iDataMidLen, 'dword*', 0)
			If @error Or $aCall[0] = 0 Then ExitLoop
			$vNowSizeBytes += $iDataMidLen
			If $iCheckCallbackFunc Then $CallBackFunc_Progress($vNowSizeBytes, $vTotalSizeBytes)
		Next
		DllCall($dll_WinInet, 'bool', 'InternetCloseHandle', 'handle', $ai_FtpOpenfile[0])
		Return 1
	EndFunc

	Func _FTP_DirDelete2($hFTPConnect, $sRemoteDirPath)
		Local $ai_FTPDelDir = DllCall($dll_WinInet, 'bool', 'FtpRemoveDirectoryW', 'handle', $hFTPConnect, 'wstr', $sRemoteDirPath)
		If @error Or $ai_FTPDelDir[0] = 0 Then Return SetError(1, 0, 0)
		Return 1
	EndFunc

	Func _FTP_DirCreate2($hFTPConnect, $sRemoteDirPath)
		Local $ai_FTPMakeDir = DllCall($dll_WinInet, 'bool', 'FtpCreateDirectoryW', 'handle', $hFTPConnect, 'wstr', $sRemoteDirPath)
		If @error Or $ai_FTPMakeDir[0] = 0 Then Return SetError(1, 0, 0)
		Return 1
	EndFunc

	Func _FTP_DirSetCurrent2($hFTPConnect, $sRemoteDirPath)
		Local $ai_FTPSetCurrentDir = DllCall($dll_WinInet, 'bool', 'FtpSetCurrentDirectoryW', 'handle', $hFTPConnect, 'wstr', $sRemoteDirPath)
		If @error Or $ai_FTPSetCurrentDir[0] = 0 Then Return SetError(1, 0, 0)
		Return 1
	EndFunc

	Func _FTP_DirGetCurrent2($hFTPConnect)
		Local $ai_FTPGetCurrentDir = DllCall($dll_WinInet, 'bool', 'FtpGetCurrentDirectoryW', 'handle', $hFTPConnect, 'wstr', "", 'dword*', 260)
		If @error Or $ai_FTPGetCurrentDir[0] = 0 Then Return SetError(1, 0, 0)
		Return $ai_FTPGetCurrentDir[2]
	EndFunc

	Func _FTP_ListToArray2($hFTPConnect, $iReturnType = 0)
		Local $asFileArray[1][3], $aDirectoryArray[1][3] = [[0, 'Size', 'Type']]
		If $iReturnType < 0 Or $iReturnType > 2 Then Return SetError(1, 0, $asFileArray)
		Local $tWIN32_FIND_DATA = DllStructCreate("DWORD dwFileAttributes; dword ftCreationTime[2]; dword ftLastAccessTime[2]; dword ftLastWriteTime[2]; DWORD nFileSizeHigh; DWORD nFileSizeLow; dword dwReserved0; dword dwReserved1; WCHAR cFileName[260]; WCHAR cAlternateFileName[14];")
		Local $iLasterror
		Local $aCallFindFirst = DllCall($dll_WinInet, 'handle', 'FtpFindFirstFileW', 'handle', $hFTPConnect, 'wstr', "", 'struct*', $tWIN32_FIND_DATA, 'dword', 0x04000000, 'dword_ptr', 0)
		If @error Or Not $aCallFindFirst[0] Then Return SetError(2, 0, '')
		Local $iDirectoryIndex = 0, $sFileIndex = 0, $bIsDir, $aCallFindNext
		Do
			$bIsDir = BitAND(DllStructGetData($tWIN32_FIND_DATA, "dwFileAttributes"), $FILE_ATTRIBUTE_DIRECTORY) = $FILE_ATTRIBUTE_DIRECTORY
			If $bIsDir And($iReturnType <> 2) Then
				$iDirectoryIndex += 1
				If UBound($aDirectoryArray) < $iDirectoryIndex + 1 Then ReDim $aDirectoryArray[$iDirectoryIndex * 2][3]
				$aDirectoryArray[$iDirectoryIndex][0] = DllStructGetData($tWIN32_FIND_DATA, "cFileName")
				$aDirectoryArray[$iDirectoryIndex][1] = __FTP_MakeQWord(DllStructGetData($tWIN32_FIND_DATA, "nFileSizeLow"), DllStructGetData($tWIN32_FIND_DATA, "nFileSizeHigh"))
				$aDirectoryArray[$iDirectoryIndex][2] = 'Folder'
			ElseIf Not $bIsDir And $iReturnType <> 1 Then
				$sFileIndex += 1
				If UBound($asFileArray) < $sFileIndex + 1 Then ReDim $asFileArray[$sFileIndex * 2][3]
				$asFileArray[$sFileIndex][0] = DllStructGetData($tWIN32_FIND_DATA, "cFileName")
				$asFileArray[$sFileIndex][1] = __FTP_MakeQWord(DllStructGetData($tWIN32_FIND_DATA, "nFileSizeLow"), DllStructGetData($tWIN32_FIND_DATA, "nFileSizeHigh"))
				$asFileArray[$sFileIndex][2] = 'File'
			EndIf
			$aCallFindNext = DllCall($dll_WinInet, 'bool', 'InternetFindNextFileW', 'handle', $aCallFindFirst[0], 'struct*', $tWIN32_FIND_DATA)
			If @error Then Return SetError(3, DllCall($dll_WinInet, 'bool', 'InternetCloseHandle', 'handle', $aCallFindFirst[0]), '')
		Until Not $aCallFindNext[0]
		DllCall($dll_WinInet, 'bool', 'InternetCloseHandle', 'handle', $aCallFindFirst[0])
		$aDirectoryArray[0][0] = $iDirectoryIndex
		$asFileArray[0][0] = $sFileIndex
		Switch $iReturnType
			Case 0
				ReDim $aDirectoryArray[$aDirectoryArray[0][0] + $asFileArray[0][0] + 1][3]
				For $i = 1 To $sFileIndex
					For $j = 0 To 2
						$aDirectoryArray[$aDirectoryArray[0][0] + $i][$j] = $asFileArray[$i][$j]
					Next
				Next
				$aDirectoryArray[0][0] += $asFileArray[0][0]
				Return $aDirectoryArray
			Case 1
				ReDim $aDirectoryArray[$iDirectoryIndex + 1][3]
				Return $aDirectoryArray
			Case 2
				ReDim $asFileArray[$sFileIndex + 1][3]
				Return $asFileArray
		EndSwitch
	EndFunc

	Func _FtpRequest($aRetMode, $aURL, $sData2Send, $CallBackFunc_Progress)
		If Not $dll_WinInet Then
			$dll_WinInet = DllOpen('wininet.dll')
			If @error Then Return SetError(1)
		EndIf
		;------------------------------------------
		Local $iError = 0, $ReData, $iProxy = '', $iAccessType = 1, $iProxyBypass = ''
		If $aRetMode[5] Then
			$iProxy = $aRetMode[5]
			$iAccessType = 3
		ElseIf $g___hProxy[$g___LastSession][0] Then
			$iProxy = $g___hProxy[$g___LastSession][0]
			$iProxyBypass = $g___hProxy[$g___LastSession][2]
			$iAccessType = 3
		EndIf
		;------------------------------------------
		If Not $g___ftpOpen[$aRetMode[8]] Then $g___ftpOpen[$aRetMode[8]] = _FTP_Open2($g___UserAgent[$g___LastSession], $iAccessType, $iProxy, $iProxyBypass)
		$g___ftpConnect[$aRetMode[8]] = _FTP_Connect2($g___ftpOpen[$aRetMode[8]], $aURL[2], $aURL[4], $aURL[5], $aURL[1])
		;------------------------------------------
		Local $sFileName = '', $sDirPath = ''
		Switch $aURL[3]
			Case '/', ''
				$aRetMode[0] = 1
			Case Else
				Local $aRemotePath = StringSplit($aURL[3], '/')
				If StringRegExp($aRemotePath[$aRemotePath[0]], '^[^\.]+\.\w+$') Then
					$sFileName = $aRemotePath[$aRemotePath[0]]
				EndIf
				If $aRemotePath[0] > 1 Then
					If _FTP_DirSetCurrent2($g___ftpConnect[$aRetMode[8]], '/') = 1 Then
						For $i = 1 To $aRemotePath[0] - ($sFileName ? 1 : 0)
							If $aRemotePath[$i] == '' Then
								If $i = 1 Then
									ContinueLoop
								Else
									$iError = 4
									ExitLoop
								EndIf
							EndIf
							If _FTP_DirSetCurrent2($g___ftpConnect[$aRetMode[8]], $aRemotePath[$i]) = 0 Then
								If _FTP_DirCreate2($g___ftpConnect[$aRetMode[8]], $aRemotePath[$i]) = 0 Then
									$iError = 5
									ExitLoop
								Else
									If _FTP_DirSetCurrent2($g___ftpConnect[$aRetMode[8]], $aRemotePath[$i]) = 0 Then
										$iError = 6
										ExitLoop
									EndIf
								EndIf
							EndIf
						Next
					Else
						$iError = 7
					EndIf
				EndIf
		EndSwitch
		;------------------------------------------
		If $iError = 0 Then
			Switch $aRetMode[0]
				Case 0
					;Null
				Case 1
					$ReData = _FTP_ListToArray2($g___ftpConnect[$aRetMode[8]])
					If @error Then $iError = 8
				Case 2, 3
					If $sFileName Then
						If $sData2Send Then
							If StringRegExp($sData2Send, '(?i)^[A-Z]:\\') And FileExists($sData2Send) Then
								$sData2Send = _GetFileInfo($sData2Send, 0)
								If @error Then
									$iError = 9
								Else
									$sData2Send = $sData2Send[2]
								EndIf
							EndIf
							If $iError = 0 Then
								_FTP_FileWriteEx($g___ftpConnect[$aRetMode[8]], $sFileName, $sData2Send, $CallBackFunc_Progress)
								If @error Then $iError = 10
							EndIf
						Else
							$ReData = _FTP_FileReadEx($g___ftpConnect[$aRetMode[8]], $sFileName, $CallBackFunc_Progress)
							If @error Then $iError = 11
						EndIf
					EndIf
					If $iError = 0 And $aRetMode[0] = 2 Then
						$ReData = BinaryToString($ReData, 4)
					EndIf
			EndSwitch
		EndIf
		Return SetError($iError, '', $ReData)
	EndFunc
#EndRegion



#Region <WinHttp Websock - Thanks [Firefox - autoscipt.com]>
	Func _HttpRequest_WebSocketReceive($iReturnType = Default, $iSession = Default) ;$iReturnType = 0: ANSI, 1: UTF8, 2: Binary
		If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
		If $iReturnType = Default Then $iReturnType = 1
		If $g___hWebSocket[$iSession] = 0 Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_WebSocketReceive', 'WebSocket handle rỗng (Chưa được khởi tạo ?)'), '')
		Switch $iReturnType
			Case 0
				Return BinaryToString(_WinHttpWebSocketRead2($g___hWebSocket[$iSession]))
			Case 1
				Return BinaryToString(_WinHttpWebSocketRead2($g___hWebSocket[$iSession]), 4)
			Case 2
				Return _WinHttpWebSocketRead2($g___hWebSocket[$iSession])
		EndSwitch
	EndFunc

	Func _HttpRequest_WebSocketSend($sData2Send, $iSession = Default)
		If IsKeyword($iSession) Or $iSession == '' Then $iSession = $g___LastSession
		If $g___hWebSocket[$iSession] = 0 Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_WebSocketSend', 'WebSocket handle rỗng'), False)
		Local $iError = _WinHttpWebSocketSend2($g___hWebSocket[$iSession], $sData2Send)
		If @error Or $iError <> 0 Then Return SetError(2, __HttpRequest_ErrNotify('_HttpRequest_WebSocketSend', 'WebSocket gửi dữ liệu thất bại'), False)
		Return True
	EndFunc

	Func _WinHttpWebSocketRequest($sData2Send)
		$g___hWebSocket[$g___LastSession] = _WinHttpWebSocketCompleteUpgrade2($g___hRequest[$g___LastSession])
		If Not $g___hWebSocket[$g___LastSession] Then Return SetError(114, __HttpRequest_ErrNotify('_WinHttpWebSocketRequest', 'WebSocket mở thất bại', -1), '')
		_HttpRequest_ConsoleWrite('> [#' & $g___LastSession & '] WebSocket mở thành công' & @CRLF)
		;------------------------------------------------------------------------------------------------
		If $sData2Send Then
			Local $iError = _WinHttpWebSocketSend2($g___hWebSocket[$g___LastSession], $sData2Send)
			If @error Or $iError <> 0 Then Return SetError(115, __HttpRequest_ErrNotify('_WinHttpWebSocketRequest', 'WebSocket gửi dữ liệu thất bại', 101, 'Warning'), '')
			_HttpRequest_ConsoleWrite('> [#' & $g___LastSession & '] WebSocket gửi dữ liệu thành công' & @CRLF)
		EndIf
	EndFunc

	Func _WinHttpWebSocketRead2($hWebSocket, $iBufferLen = Default)
		If $iBufferLen = Default Then $iBufferLen = $g___BytesPerLoop
		Local $tBuffer = 0, $bRecv = Binary(""), $iError, $iBytesRead = 0, $iBufferType = 0
		Do
			_HttpRequest_ConsoleWrite('> [#' & $g___LastSession & '] WebSocket đang chờ dữ liệu gửi về...')
			$tBuffer = DllStructCreate("byte[" & $iBufferLen & "]")
			$iError = _WinHttpWebSocketReceive2($hWebSocket, $tBuffer, $iBytesRead, $iBufferType)
			If @error Or $iError <> 0 Then Return SetError(1, __HttpRequest_ErrNotify('_WinHttpWebSocketRead2', @CRLF & 'WebSocket không nhận được phản hồi'), '')
			$bRecv &= BinaryMid(DllStructGetData($tBuffer, 1), 1, $iBytesRead)
			$iBufferLen -= $iBytesRead
			$tBuffer = 0
			ConsoleWrite('...')
		Until $iBufferType <> 1 ;WEBSOCKET_BINARYFRAGMENT_BUFFERTYPE
		ConsoleWrite('OK' & @CRLF)
		Return $bRecv
	EndFunc

	Func _WinHttpWebSocketCompleteUpgrade2($hRequest, $pContext = 0)
		Local $aCall = DllCall($dll_WinHttp, "handle", "WinHttpWebSocketCompleteUpgrade", "handle", $hRequest, "dword_ptr", $pContext)
		If @error Then Return SetError(@error, @extended, -1)
		Return $aCall[0]
	EndFunc

	Func _WinHttpWebSocketSend2($hWebSocket, $vData, $iBufferType = 0) ;$iBufferType = WEBSOCKET_BINARYMESSAGE_BUFFERTYPE
		Local $tBuffer = 0, $iBufferLen = 0
		If Not IsBinary($vData) Then $vData = StringToBinary($vData, 4)
		$iBufferLen = BinaryLen($vData)
		If $iBufferLen > 0 Then
			$tBuffer = DllStructCreate("byte[" & $iBufferLen & "]")
			DllStructSetData($tBuffer, 1, $vData)
		EndIf
		Local $aCall = DllCall($dll_WinHttp, 'dword', "WinHttpWebSocketSend", "handle", $hWebSocket, "int", $iBufferType, "ptr", DllStructGetPtr($tBuffer), 'dword', $iBufferLen)
		If @error Then Return SetError(@error, @extended, -1)
		Return $aCall[0]
	EndFunc

	Func _WinHttpWebSocketReceive2($hWebSocket, $tBuffer, ByRef $iBytesRead, ByRef $iBufferType)
		Local $aCall = DllCall($dll_WinHttp, "handle", "WinHttpWebSocketReceive", "handle", $hWebSocket, "ptr", DllStructGetPtr($tBuffer), 'dword', DllStructGetSize($tBuffer), "dword*", $iBytesRead, "int*", $iBufferType)
		If @error Then Return SetError(@error, @extended, -1)
		$iBytesRead = $aCall[4]
		$iBufferType = $aCall[5]
		Return $aCall[0]
	EndFunc

	Func _WinHttpWebSocketClose2($hWebSocket, $iStatus = Default, $tReason = 0)
		If $iStatus = Default Then $iStatus = 1000 ;WEBSOCKER_SUCCESS_CLOSESTATUS
		Local $aCall = DllCall($dll_WinHttp, "handle", "WinHttpWebSocketClose", "handle", $hWebSocket, "ushort", $iStatus, "ptr", DllStructGetPtr($tReason), 'dword', DllStructGetSize($tReason))
		;If @error Then Return SetError(@error, @extended, 0)
		;Return $aCall[0]
	EndFunc

	Func _WinHttpWebSocketQueryCloseStatus2($hWebSocket, ByRef $iStatus, ByRef $iReasonLengthConsumed, $tCloseReasonBuffer = 0)
		Local $aCall = DllCall($dll_WinHttp, "handle", "WinHttpWebSocketQueryCloseStatus", "handle", $hWebSocket, "ushort*", $iStatus, "ptr", DllStructGetPtr($tCloseReasonBuffer), 'dword', DllStructGetSize($tCloseReasonBuffer), "DWORD*", $iReasonLengthConsumed)
		If @error Then Return SetError(@error, @extended, -1)
		$iStatus = $aCall[2]
		$iReasonLengthConsumed = $aCall[5]
		Return $aCall[0]
	EndFunc
#EndRegion



#Region <Set Binary Image To Ctrl + Simple Captcha GUI>
	Func _Image_GetDimension($sBinaryData_Or_FilePath, $Release_hBitmap = True, $isFilePath = False)
		_GDIPlus_Startup()
		If $isFilePath Or FileExists($sBinaryData_Or_FilePath) Then
			Local $___hBitmap = _GDIPlus_BitmapCreateFromFile($sBinaryData_Or_FilePath)
		Else
			Local $___hBitmap = _GDIPlus_BitmapCreateFromMemory(Binary($sBinaryData_Or_FilePath))
		EndIf
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('_Image_GetDimension', 'Tạo Bitmap thất bại'))
		Local $___w = _GDIPlus_ImageGetWidth($___hBitmap)
		Local $___h = _GDIPlus_ImageGetHeight($___hBitmap)
		If $Release_hBitmap Then
			_GDIPlus_BitmapDispose($___hBitmap)
			_GDIPlus_Shutdown()
			Local $aRet = [$___w, $___h]
		Else
			Local $aRet = [$___hBitmap, $___w, $___h]
		EndIf
		Return $aRet
	EndFunc

	Func _Image_SetGUI($sBinaryData_Or_FilePath_Or_URL_Or_arrayHBitmap, $idCtrl_Or_hWnd, $width_Image = Default, $height_Image = Default)
		_GDIPlus_Startup()
		If Not IsHWnd($idCtrl_Or_hWnd) Then
			$idCtrl_Or_hWnd = GUICtrlGetHandle($idCtrl_Or_hWnd)
			If @error Or $idCtrl_Or_hWnd = 0 Then Return SetError(1, __HttpRequest_ErrNotify('_Image_SetGUI', 'Không tìm thấy Handle của Control hoặc Cửa sổ đã gọi'))
		EndIf
		If BitAND(WinGetState($idCtrl_Or_hWnd), 2) = 0 Then Return SetError(2, __HttpRequest_ErrNotify('_Image_SetGUI', 'Hàm này phải đặt bên dưới hàm GUISetState(@SW_SHOW)'), '')
		If UBound($sBinaryData_Or_FilePath_Or_URL_Or_arrayHBitmap) <> 3 Then
			If StringRegExp($sBinaryData_Or_FilePath_Or_URL_Or_arrayHBitmap, '(?i)^https?://') Then
				$sBinaryData_Or_FilePath_Or_URL_Or_arrayHBitmap = _HttpRequest(3, $sBinaryData_Or_FilePath_Or_URL_Or_arrayHBitmap)
				If @error Then Return SetError(3, __HttpRequest_ErrNotify('_Image_SetGUI', 'Lấy dữ liệu ảnh từ URL thất bại'))
			EndIf
			Local $aHBitmap = _Image_GetDimension($sBinaryData_Or_FilePath_Or_URL_Or_arrayHBitmap, False, FileExists($sBinaryData_Or_FilePath_Or_URL_Or_arrayHBitmap))
			If @error Then Return SetError(4, __HttpRequest_ErrNotify('_Image_SetGUI', 'Tạo dữ liệu Bitmap thất bại'))
		Else
			Local $aHBitmap = $sBinaryData_Or_FilePath_Or_URL_Or_arrayHBitmap
		EndIf
		If $width_Image = Default Or $width_Image = 0 Then $width_Image = $aHBitmap[1]
		If $height_Image = Default Or $height_Image = 0 Then $height_Image = $aHBitmap[2]
		Local $___hGraphics = _GDIPlus_GraphicsCreateFromHWND($idCtrl_Or_hWnd)
		_GDIPlus_GraphicsDrawImageRectRect($___hGraphics, $aHBitmap[0], 0, 0, $aHBitmap[1], $aHBitmap[2], 0, 0, $width_Image, $height_Image)
		_GDIPlus_BitmapDispose($aHBitmap[0])
		_GDIPlus_GraphicsDispose($___hGraphics)
		_GDIPlus_Shutdown()
		Local $aRet = [$aHBitmap[1], $aHBitmap[2]]
		Return $aRet
	EndFunc

	Func _Image_SetSimpleCaptchaGUI($BinaryCaptcha_Or_FilePath_Or_URL, $___x = -1, $___y = -1, $___hParent = Default)
		If StringRegExp($BinaryCaptcha_Or_FilePath_Or_URL, '(?i)^https?://') Then
			$BinaryCaptcha_Or_FilePath_Or_URL = _HttpRequest(3, $BinaryCaptcha_Or_FilePath_Or_URL)
			If @error Then Return SetError(1, __HttpRequest_ErrNotify('_Image_SetSimpleCaptchaGUI', 'Lấy dữ liệu ảnh từ URL thất bại'), '')
		EndIf
		Local $aHBitmap = _Image_GetDimension($BinaryCaptcha_Or_FilePath_Or_URL, False)
		If @error Then Return SetError(2, __HttpRequest_ErrNotify('_Image_SetSimpleCaptchaGUI', 'Tạo dữ liệu Bitmap thất bại'), '')
		If $___hParent = Default Then $___hParent = 0
		Local $___w = $aHBitmap[1]
		Local $___h = $aHBitmap[2]
		Local $___hGUI_Captcha = GUICreate("Captcha Display", $___w + 4, $___h + 25, $___x, $___y, 0x80800000, 0x8 + IsHWnd($___hParent) * 0x40, $___hParent)
		Local $PicCtrl = GUICtrlCreateLabel('', 2, 2, $___w, $___h, 0x800000, 0x100000)
		Local $InputCtrl = GUICtrlCreateInput('', 2, $___h + 3, $___w - 30, 20)
		Local $OKCtrl = GUICtrlCreateButton('OK', $___w - 27, $___h + 3, 30, 20, 0x1)
		GUISetState(@SW_SHOW, $___hGUI_Captcha)
		_Image_SetGUI($aHBitmap, $PicCtrl)
		If @error Then Return SetError(3, __HttpRequest_ErrNotify('_Image_SetSimpleCaptchaGUI', 'Set dữ liệu ảnh lên GUI thất bại'), '')
		Local $CaptchaRs = ''
		While Sleep(30)
			Switch GUIGetMsg()
				Case $OKCtrl
					$CaptchaRs = GUICtrlRead($InputCtrl)
					ExitLoop
				Case -3
					ExitLoop
			EndSwitch
		WEnd
		GUIDelete($___hGUI_Captcha)
		Return $CaptchaRs
	EndFunc
#EndRegion



#Region IE External
	Func __IE_Init_GoogleBox($sUser, $sPassword, $sURL = Default, $vFuncCallback = '', $vDebug = False, $vTimeOut = Default)
		;	Local Const $mUserAgent = 'User-Agent: Mozilla/5.0 (Linux; U; Android 4.4.2; en-us; SCH-I535 Build/KOT49H) AppleWebKit/534.30 (KHTML, like Gecko) Version/4.0 Mobile Safari/534.30'
		$sUser = StringRegExpReplace($sUser, '(?i)@gmail.com[\.\w]*?$', '', 1)
		If $sURL = Default Then $sURL = ''
		;------------------------------------------------------------------------------------------------------
		Local $sHTML = _HttpRequest(2, 'https://accounts.google.com/ServiceLogin?hl=en&passive=true&continue=' & _URIEncode($sURL), '', '', '', 'User-Agent: ' & $g___defUserAgent)
		If @error Or $sHTML = '' Then Return SetError(-1, __HttpRequest_ErrNotify('_GoogleBox', 'Request đến trang đăng nhập thất bại'), '')
		$sHTML = StringRegExpReplace($sHTML, '(?i)(<input.*?\h+?id="Email".*?\h+?value=")(")', '$1' & $sUser & '$2', 1)
		$sHTML = StringReplace($sHTML, '<body', '<body onLoad="document.getElementById(''next'').click();"', 1, 1)
		;------------------------------------------------------------------------------------------------------
		Local $oIE = ObjCreate("Shell.Explorer.2")
		If Not IsObj($oIE) Then Return SetError(-2, __HttpRequest_ErrNotify('_GoogleBox', 'Không thể tạo IE Object'), '')
		;------------------------------------------------------------------------------------------------------
		If $vTimeOut = Default Then $vTimeOut = 20000
		If $vTimeOut < 10000 Then $vTimeOut = 10000
		Local $vReturn = '', $vError = 0, $_oid_Pass, $_old_Locate
		;------------------------------------------------------------------------------------------------------
		Local $GUI_EmbededGG = GUICreate("Google Box", 800, 600)
		GUICtrlCreateObj($oIE, 0, 0, 800, 600)
		If $vDebug Then GUISetState()
		;------------------------------------------------------------------------------------------------------
		With $oIE
			.navigate('about:blank')
			Local $sTimer = TimerInit()
			While .busy()
				If TimerDiff($sTimer) > $vTimeOut Then Return SetError(GUIDelete($GUI_EmbededGG) * -3, __HttpRequest_ErrNotify('_GoogleBox', 'TimeOut #1'), '')
				Sleep(40)
			WEnd
			;------------------------------------------------------------------------------------------------------
			.document.write($sHTML)
			.document.close()
			;------------------------------------------------------------------------------------------------------
			_HttpRequest_ConsoleWrite('> [Google Login] Đang cài đặt Tài khoản ...')
			For $i = 1 To 2
				$sTimer = TimerInit()
				Do
					If TimerDiff($sTimer) > $vTimeOut Then Return SetError(GUIDelete($GUI_EmbededGG) * -4, __HttpRequest_ErrNotify('_GoogleBox', 'TimeOut #2'), '')
					$_oid_Pass = .document.getElementById('Passwd')
					Sleep(40)
					If $i = 1 Then ConsoleWrite('..')
				Until IsObj($_oid_Pass)
				$_oid_Pass.value = $sPassword
				$_old_Locate = .locationName()
				If $i = 1 Then
					ConsoleWrite(' (' & Int(TimerDiff($sTimer)) & 'ms)' & @CRLF)
					.document.getElementById('signIn').click()
				EndIf
				;------------------------------------------------------------------------------------------------------
				_HttpRequest_ConsoleWrite('> [Google Login] ' & ($i = 1 ? 'Đang chuyển hướng tới địa chỉ đích ...' : 'Đang chờ giải Captcha'))
				$sTimer = TimerInit()
				Do
					If TimerDiff($sTimer) > $vTimeOut * $i Then
						ConsoleWrite(@CRLF)
						Return SetError(GUIDelete($GUI_EmbededGG) * -5, __HttpRequest_ErrNotify('_GoogleBox', 'TimeOut #3'), '')
					EndIf
					Sleep(40)
					ConsoleWrite('..')
				Until .locationName() <> $_old_Locate
				ConsoleWrite(' (' & Int(TimerDiff($sTimer)) & 'ms)' & @CRLF & @CRLF)
				;------------------------------------------------------------------------------------------------------
				If Not StringRegExp(.document.body.innerHtml, '(?i)<div class="?captcha-box"?>') Then ExitLoop
				_HttpRequest_ConsoleWrite('> [Google Login] Phát hiện Captcha' & @CRLF)
				GUISetState()
			Next
			;--------------------------------------------------------------
			If $vFuncCallback Then
				Local $aFuncCallback = StringSplit($vFuncCallback, '|')
				Local $sFuncCallbackName = $aFuncCallback[1]
				$aFuncCallback[0] = 'CallArgArray'
				$aFuncCallback[1] = $oIE
				$vReturn = Call($sFuncCallbackName, $aFuncCallback)
				$vError = @error
			EndIf
		EndWith
		;------------------------------------------------------------------------------------------------------
		If $vDebug Then
			While GUIGetMsg() <> -3
				Sleep(35)
			WEnd
		EndIf
		Return SetError($vError * GUIDelete($GUI_EmbededGG), '', $vReturn)
	EndFunc

	Func __IE_Init_RecaptchaBox($sURL, $vAdvancedMode, $hGUI, $___GUI_Offset, $Custom_RegExp_GetDataSiteKey, $vTimeOut)
		Local $oIE = ObjCreate("Shell.Explorer.2")
		If Not IsObj($oIE) Then Return SetError(1, __HttpRequest_ErrNotify('__IE_Init_RecaptchaBox', 'Không thể tạo IE Object'), '')
		Local $sReCaptchaResponse = '', $iError = 0, $sDataSiteKey = '', $isInvisible = 0
		;------------------------------------------------------------------------------------------------------
		GUICtrlSetDefBkColor(0x222222, $hGUI)
		GUICtrlSetDefColor(0xFFFFFF, $hGUI)
		GUISetFont(10, 600)
		GUICtrlCreateLabel('ReCaptcha Box', 2, 2, 377, 22, 0x201, 0x100000)
		Local $__idCloseButton = GUICtrlCreateLabel('X', 380, 2, 22, 22, 0x201)
		GUICtrlSetBkColor(-1, 0xFF0011)
		GUICtrlCreateObj($oIE, 2, 25, 400, 580)
		;------------------------------------------------------------------------------------------------------
		With $oIE
			;.navigate2($sURL, 2, Default, Default, 'User-Agent: ' & $g___UserAgent)
			.navigate($sURL)
			TrayTip('ReCaptcha', 'Đang tải thông tin Recaptcha...', 0)
			_HttpRequest_ConsoleWrite('> [reCAPTCHA] Đang tải trang ...')
			Local $sTimer = TimerInit()
			While .busy()
				If TimerDiff($sTimer) > $vTimeOut Then Return SetError(2, __HttpRequest_ErrNotify('__IE_Init_RecaptchaBox', 'TimeOut #1'), '')
				Sleep(100)
				ConsoleWrite('..')
			WEnd
			ConsoleWrite(' (' & Int(TimerDiff($sTimer)) & 'ms)' & @CRLF)
			;------------------------------------------------------------------------------------------------------
			$sTimer = TimerInit()
			;------------------------------------------------
			Local $sourceHtml = .document.body.innerHTML
			;------------------------------------------------
			If $Custom_RegExp_GetDataSiteKey Then
				$sDataSiteKey = StringRegExp($Custom_RegExp_GetDataSiteKey, '(?i)["'']?siteKey[''"]?\s*?:\s*?[''"](.*?)[''"]', 1)
				If Not @error And $sDataSiteKey[0] Then
					$sDataSiteKey = $sDataSiteKey[0]
					$isInvisible = (StringInStr($Custom_RegExp_GetDataSiteKey, 'invisible', 0, 1) > 0)
				Else
					$sDataSiteKey = ''
				EndIf
			EndIf
			;------------------------------------------------
			If $sDataSiteKey == '' Then
				Local $oiFrames = .document.GetElementsByTagName("iframe")
				If @error Then
					$iError = 1
				Else
					For $oiFrame In $oiFrames
						$sDataSiteKey = StringRegExp($oiFrame.src, '(?i)^https://www.google.com/recaptcha/api2/.*?\&?k=([^\&]+)', 1)
						If Not @error And $sDataSiteKey[0] Then
							$sDataSiteKey = $sDataSiteKey[0]
							$isInvisible = (StringInStr($oiFrame.src, 'size=invisible', 0, 1) > 0)
							ExitLoop
						Else
							$sDataSiteKey = ''
						EndIf
					Next
				EndIf
				If $iError = 1 Or $sDataSiteKey == '' Then
					If $sourceHtml = '' Then Return SetError(3, __HttpRequest_ErrNotify('__IE_Init_RecaptchaBox', 'Đọc html thất bại'), '')
					If $Custom_RegExp_GetDataSiteKey And $Custom_RegExp_GetDataSiteKey <> Default Then
						Local $sDataSiteKey = StringRegExp($sourceHtml, $Custom_RegExp_GetDataSiteKey, 1)
						If @error Then Return SetError(4, __HttpRequest_ErrNotify('__IE_Init_RecaptchaBox', 'Không thể lấy data-sitekey #1'), '')
					Else
						Local $sDataSiteKey = StringRegExp($sourceHtml, '(?im)data-sitekey\h*?=\h*?["'']([\w\-]{20,})["'']', 1)
						If @error Then $sDataSiteKey = StringRegExp($sourceHtml, '(?i)(?:\?|&amp;|&|;)k=([\w\-]{20,})(?:"|&amp;|&|;|''|$)', 1)
						If @error Then $sDataSiteKey = StringRegExp($sourceHtml, '(?i)reCAPTCHA[^=]+=\h*?[''"]([\w\-]{20,})[''"]', 1)
						If @error Then $sDataSiteKey = StringRegExp($sourceHtml, '(?i)["'']reCAPTCHA_?site_?key["'']\h*?:\h*?["'']([\w\-]{20,})["'']', 1)
						If @error Then Return SetError(5, __HttpRequest_ErrNotify('__IE_Init_RecaptchaBox', 'Không thể lấy data-sitekey #2 (Có thể sitekey nằm trong code js)'), '')
					EndIf
					$sDataSiteKey = $sDataSiteKey[0]
					$isInvisible = (StringRegExp($sourceHtml, '(?i)data-sitekey="[^"]+" .*?data-size="invisible"|data-size="invisible" .*?data-sitekey="[^"]+"') > 0)
				EndIf
			EndIf
			;------------------------------------------------------------------------------------------------------
			ConsoleWrite('> [reCAPTCHA] Data-SiteKey : ' & $sDataSiteKey & @CRLF)
			;------------------------------------------------------------------------------------------------------
			Local $hIE = ControlGetHandle($hGUI, '', '[Classnn:Internet Explorer_Server1]')
			ConsoleWrite('> [reCAPTCHA] Hwnd IE : ' & $hIE & @CRLF)
			;-----------------------------------------------------
			.document.write( _
					'<!DOCTYPE html><html>' & _
					'	<head>' & _
					'		<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE11,IE=edge,chrome=1">' & _
					'		<style>div.g-recaptcha {padding: 50% 0}</style>' & _
					'		<script src="https://www.google.com/recaptcha/api.js?hl=vi&onload=onloadCallback" async defer></script>' & _
					'		<script>var onloadCallback = function(){grecaptcha.execute()}</script>' & _
					'	</head>' & _
					'	<body bgcolor="#222222">' & _
					'		<div align="center" class="g-recaptcha" ' & ($isInvisible ? 'data-size="invisible"' : '') & ' data-theme="dark" data-sitekey="' & $sDataSiteKey & '"></div>' & _
					'	</body>' & _
					'</html>')
			.document.close()
			_HttpRequest_ConsoleWrite('> [reCAPTCHA] Đang khởi tạo ReCaptcha ...')
			While .busy()
				If TimerDiff($sTimer) > $vTimeOut Then Return SetError(2, __HttpRequest_ErrNotify('__IE_Init_RecaptchaBox', 'TimeOut #2'), '')
				Sleep(100)
				ConsoleWrite('..')
			WEnd
			ConsoleWrite(' (' & Int(TimerDiff($sTimer)) & 'ms)' & @CRLF)
			;------------------------------------------------------------------------------------------------------
			$sTimer = TimerInit()
			Local $oReCaptchaResponse
			Do
				If TimerDiff($sTimer) > $vTimeOut Then Return SetError(2 + TrayTip('', '', 1), __HttpRequest_ErrNotify('__IE_Init_RecaptchaBox', 'TimeOut #3'), '')
				Sleep(Random(500, 1500, 1))
				$oReCaptchaResponse = .document.getElementById("g-recaptcha-response")
			Until IsObj($oReCaptchaResponse)
			;------------------------------------------------------------------------------------------------------
			If $isInvisible And $oReCaptchaResponse.value Then
				$sReCaptchaResponse = $oReCaptchaResponse.value
				GUISetState(@SW_HIDE, $hGUI)
			Else
				GUISetState(@SW_SHOW, $hGUI)
				TrayTip('', '', 1)
				;------------------------------------------------------------------------------------------------------
				If Not $isInvisible Then
					If $hIE Then
						__IE_MouseClick($hIE, 80, 260 - 25)
					Else
						Local $aPosMouse = MouseGetPos()
						MouseClick('left', 80, 260, 1, 0)
						MouseMove($aPosMouse[0], $aPosMouse[1], 0)
						Sleep(Random(1000, 2000, 1))
					EndIf
				EndIf
				;------------------------------------------------------------------------------------------------------
				_GDIPlus_Startup()
				Local $vGDI_Startup_Error = @error
				Local $___aMouseInfo, $___aMouseInfo_Old, $___aPosCurMem, $vClickDrag = False, $iClickNextButton
				If Not $dll_Gdi32 Then
					$dll_Gdi32 = DllOpen('gdi32.dll')
					If @error Then $vGDI_Startup_Error = 1
				EndIf
				;------------------------------------------------------------------------------------------------------
				Do
					Sleep(10)
					Switch GUIGetMsg()
						Case $__idCloseButton
							Return SetError(6, __HttpRequest_ErrNotify('__IE_Init_RecaptchaBox', 'Đã huỷ việc giải Captcha'), '')

						Case -8, -10 ;$GUI_EVENT_PRIMARYUP, $GUI_EVENT_SECONDARYUP
							$vClickDrag = False

						Case -7, -9 ;$GUI_EVENT_PRIMARYDOWN, $GUI_EVENT_SECONDARYDOWN
							$___aMouseInfo_Old = GUIGetCursorInfo($hGUI)
							If @error Then ContinueLoop
							$vClickDrag = True
							$iTimerClickDrag = TimerInit()
							While __IE_IsMousePressed(1) Or __IE_IsMousePressed(2)
								If TimerDiff($iTimerClickDrag) > 150 Then
									ContinueCase
								EndIf
							WEnd

						Case -11 ; $GUI_EVENT_MOUSEMOVE
							If $vClickDrag And $vGDI_Startup_Error = 0 Then
								Select
									Case __IE_IsMousePressed(1)
										$___aPosCurMem = $___aMouseInfo_Old[0] & '|' & $___aMouseInfo_Old[1]
										$iClickNextButton = __IE_RecaptchaBox_GuiOnDrawLine($hGUI, $___GUI_Offset, 2, $___aPosCurMem, $___aMouseInfo, $___aMouseInfo_Old, 20, True)
										__IE_RecaptchaBox_CalculateRectClick($hGUI, $hIE, $___aPosCurMem, $iClickNextButton)
									Case __IE_IsMousePressed(2)
										$___aPosCurMem = $___aMouseInfo_Old[0] & '|' & $___aMouseInfo_Old[1]
										$iClickNextButton = __IE_RecaptchaBox_GuiOnDrawRect($hGUI, $___GUI_Offset, 3, $___aPosCurMem, $___aMouseInfo, $___aMouseInfo_Old)
										__IE_RecaptchaBox_CalculateRectClick($hGUI, $hIE, $___aPosCurMem, $iClickNextButton)
									Case Else
										$vClickDrag = False
								EndSelect
							EndIf
					EndSwitch
					;------------------------------------------------------------------------------------------------------
					$sReCaptchaResponse = $oReCaptchaResponse.value
				Until $sReCaptchaResponse
			EndIf
		EndWith
		;------------------------------------------------------------------------------------------------------
		If $vAdvancedMode Then
			Local $aResponse = [$sReCaptchaResponse, _IE_GetCookie($sURL), $sourceHtml, $oIE.document.body.innerHTML]
			Return $aResponse
		Else
			Return $sReCaptchaResponse
		EndIf
	EndFunc

	Func __IE_IsMousePressed($sHexKey)
		Local $aReturn = DllCall($dll_User32, "short", "GetAsyncKeyState", "int", $sHexKey)
		If @error Then Return False
		Return BitAND($aReturn[0], 0x8000) <> 0
	EndFunc

	Func __IE_MouseClick($hwnd, $x, $y)
		Local $lParam = $y * 65536 + $x
		DllCall($dll_User32, "bool", "PostMessage", "hwnd", $hwnd, "uint", 0x201, "wparam", 0x1, "lparam", $lParam)
		DllCall($dll_User32, "bool", "PostMessage", "hwnd", $hwnd, "uint", 0x202, "wparam", 0x0, "lparam", $lParam)
	EndFunc

	Func __IE_RecaptchaBox_GuiOnDrawRect($hGUI, $___GUI_Offset, $iMouseEvent, ByRef $___aPosCurMem, $___aMouseInfo, $___aMouseInfo_Old)
		Local $iMouseEventReverse = BitXOR($iMouseEvent, 1)
		Local $aAbsPos = $___aMouseInfo_Old, $iClickNextButton = 0
		Local $posGUI = WinGetPos($hGUI)
		;-----------------------------------------------------------------------------------
		Local $___GDI_DrawGUI = GUICreate("HH Draw GUI", $posGUI[2], $posGUI[3], $___GUI_Offset, $___GUI_Offset, 0x80000000, 0x40 + 0x8, $hGUI)
		WinSetTrans($___GDI_DrawGUI, '', 80)
		Local $___GDI_Rect = GUICtrlCreateLabel('', $aAbsPos[0], $aAbsPos[1], 0, 0, 0x800000)
		GUICtrlSetBkColor(-1, 0xff0011)
		GUISetState(@SW_SHOW, $___GDI_DrawGUI)
		GUISetCursor(0, 1, $___GDI_DrawGUI)
		;-----------------------------------------------------------------------------------
		Do
			$___aMouseInfo = GUIGetCursorInfo($hGUI)
			If @error Or Not IsArray($___aMouseInfo) Then ExitLoop
			If $___aMouseInfo[$iMouseEventReverse] = 1 Then
				$iClickNextButton = 1
				ExitLoop
			EndIf
			;-----------------------------------------------------
			If $___aMouseInfo[0] <> $___aMouseInfo_Old[0] Or $___aMouseInfo[1] <> $___aMouseInfo_Old[1] Then
				If $___aMouseInfo[1] > $aAbsPos[1] Then
					If $___aMouseInfo[0] > $aAbsPos[0] Then
						GUICtrlSetPos($___GDI_Rect, $aAbsPos[0], $aAbsPos[1], $___aMouseInfo[0] - $aAbsPos[0], $___aMouseInfo[1] - $aAbsPos[1])
					Else ;-------------------------------------------
						GUICtrlSetPos($___GDI_Rect, $___aMouseInfo[0], $aAbsPos[1], $aAbsPos[0] - $___aMouseInfo[0], $___aMouseInfo[1] - $aAbsPos[1])
					EndIf
				Else ;---------------------------------------------------------------------------------------------------------------------------------------
					If $___aMouseInfo[0] > $aAbsPos[0] Then
						GUICtrlSetPos($___GDI_Rect, $aAbsPos[0], $___aMouseInfo[1], $___aMouseInfo[0] - $aAbsPos[0], $aAbsPos[1] - $___aMouseInfo[1])
					Else ;-------------------------------------------
						GUICtrlSetPos($___GDI_Rect, $___aMouseInfo[0], $___aMouseInfo[1], $aAbsPos[0] - $___aMouseInfo[0], $aAbsPos[1] - $___aMouseInfo[1])
					EndIf
				EndIf
				;-----------------------------------------------------------------------------------
				$___aMouseInfo_Old = $___aMouseInfo
			EndIf
		Until $___aMouseInfo[$iMouseEvent] = 0
		;-----------------------------------------------------------------------------------
		Local $x0 = $aAbsPos[0], $y0 = $aAbsPos[1], $x1 = $___aMouseInfo[0], $y1 = $___aMouseInfo[1], $wSelect = Abs($x1 - $x0), $hSelect = Abs($y1 - $y0)
		Select
			Case $x1 < $x0 And $y1 > $y0
				$x0 = $x1
			Case $x1 < $x0 And $y1 < $y0
				$x0 = $x1
				$y0 = $y1
			Case $x1 > $x0 And $y1 < $y0
				$y0 = $y1
		EndSelect
		;-----------------------------------------------------------------------------------
		Local $xPart = 4, $yPart = 4
		If Mod($wSelect, $xPart) Then $wSelect += ($xPart - Mod($wSelect, $xPart))
		If Mod($hSelect, $yPart) Then $hSelect += ($yPart - Mod($hSelect, $yPart))
		;-----------------------------------------------------------------------------------
		For $x = 0 To $wSelect Step $wSelect / $xPart
			For $y = 0 To $hSelect Step $hSelect / $yPart
				$___aPosCurMem &= '|' & ($x + $x0) & '|' & ($y + $y0)
			Next
		Next
		$___aPosCurMem &= '|' & $x1 & '|' & $y1
		;-------------------------------------------------------------------------------------------------------------------------------------------
		GUICtrlDelete($___GDI_Rect)
		GUIDelete($___GDI_DrawGUI)
		;-------------------------------------------------------------------------------------------------------------------------------------------
		Return $iClickNextButton
	EndFunc

	Func __IE_RecaptchaBox_GuiOnDrawLine($hGUI, $___GUI_Offset, $iMouseEvent, ByRef $___aPosCurMem, $___aMouseInfo, $___aMouseInfo_Old, $iSizePen, $iEasyModeGUI = True)
		Local $iMouseEventReverse = BitXOR($iMouseEvent, 1)
		Local $iClickNextButton = 0
		Local $posGUI = WinGetPos($hGUI)
		;------------------------------------------------------------------------------------------------------
		If $iEasyModeGUI Then
			Local $___GDI_DrawGUI = GUICreate("HH Draw GUI", $posGUI[2], $posGUI[3], $___GUI_Offset, $___GUI_Offset, 0x80000000, 0x80000 + 0x40 + 0x8, $hGUI)
			GUISetBkColor(0x123456, $___GDI_DrawGUI)
			DllCall($dll_User32, "bool", "SetLayeredWindowAttributes", "hwnd", $___GDI_DrawGUI, "INT", 0x563412, "byte", 255, "dword", 0x3)
			GUISetState(@SW_SHOW, $___GDI_DrawGUI)
		Else ;-----------------------------------------------------------
			Local $___WinAPI_hDDC = _WinAPI_GetDC($hGUI)
			Local $___WinAPI_hCDC = _WinAPI_CreateCompatibleDC($___WinAPI_hDDC)
			Local $___GDI_hCloneGUI = _WinAPI_CreateCompatibleBitmap($___WinAPI_hDDC, $posGUI[2], $posGUI[3])
			_WinAPI_SelectObject($___WinAPI_hCDC, $___GDI_hCloneGUI)
			_WinAPI_BitBlt($___WinAPI_hCDC, 0, 0, $posGUI[2], $posGUI[3], $___WinAPI_hDDC, 0, 0, 0x00CC0020) ;$__SCREENCAPTURECONSTANT_SRCCOPY
			_WinAPI_ReleaseDC($hGUI, $___WinAPI_hDDC)
			_WinAPI_DeleteDC($___WinAPI_hCDC)
			;----------------------------------------------------
			Local $___GDI_DrawGUI = GUICreate("HH Draw GUI", $posGUI[2], $posGUI[3], $___GUI_Offset, $___GUI_Offset, 0x80000000, 0x40 + 0x8, $hGUI)
			GUISetState(@SW_SHOW, $___GDI_DrawGUI)
			;----------------------------------------------------
			Local $___GDI_hGraphic = _GDIPlus_GraphicsCreateFromHWND($___GDI_DrawGUI)
			Local $___GDI_hBitmap = _GDIPlus_BitmapCreateFromHBITMAP($___GDI_hCloneGUI)
			_GDIPlus_GraphicsDrawImage($___GDI_hGraphic, $___GDI_hBitmap, 0, 0)
			_GDIPlus_BitmapDispose($___GDI_hBitmap)
			_WinAPI_DeleteObject($___GDI_hCloneGUI)
			_GDIPlus_GraphicsDispose($___GDI_hGraphic)
		EndIf
		GUISetCursor(0, 1, $___GDI_DrawGUI)
		;-----------------------------------------------------------------------------------------------------
		Local $___WinAPI_hWDC = _WinAPI_GetWindowDC($___GDI_DrawGUI)
		Local $___WinAPI_hPen = _WinAPI_CreatePen(0, $iSizePen, 0x1100FF)
		Local $___WinAPI_oSelect = _WinAPI_SelectObject($___WinAPI_hWDC, $___WinAPI_hPen)
		;-----------------------------------------------------------------------------------------------------
		Local $___WinAPI_hPen2 = _WinAPI_CreatePen(0, $iSizePen * 2, 0)
		Local $___WinAPI_oSelect2 = _WinAPI_SelectObject($___WinAPI_hWDC, $___WinAPI_hPen2)
		_WinAPI_DrawLine($___WinAPI_hWDC, $___aMouseInfo_Old[0], $___aMouseInfo_Old[1], $___aMouseInfo_Old[0], $___aMouseInfo_Old[1])
		_WinAPI_SelectObject($___WinAPI_hWDC, $___WinAPI_oSelect2)
		_WinAPI_DeleteObject($___WinAPI_hPen2)
		;------------------------------------------------------------------------------------------------------
		Local $VectorX, $VectorY, $A, $B
		Do
			$___aMouseInfo = GUIGetCursorInfo($___GDI_DrawGUI)
			If @error Or Not IsArray($___aMouseInfo) Then ExitLoop
			If $___aMouseInfo[$iMouseEventReverse] = 1 Then
				$iClickNextButton = 1
				ExitLoop
			EndIf
			;-----------------------------------------------------
			$VectorX = $___aMouseInfo[0] - $___aMouseInfo_Old[0]
			$VectorY = $___aMouseInfo[1] - $___aMouseInfo_Old[1]
			If Abs($VectorX) > 5 Or Abs($VectorY) > 5 Then
				$A = $VectorY / $VectorX
				$B = $___aMouseInfo[1] - $___aMouseInfo[0] * $A
				If Abs($VectorY) > 50 Then
					For $k = $___aMouseInfo_Old[1] To $___aMouseInfo[1] Step 10 * ($___aMouseInfo_Old[1] > $___aMouseInfo[1] ? -1 : 1)
						If $VectorX = 0 Then
							$___aPosCurMem &= '|' & $___aMouseInfo[0] & '|' & $k
						Else
							$___aPosCurMem &= '|' & ($k - $B) / $A & '|' & $k
						EndIf
					Next
				ElseIf Abs($VectorX) > 50 Then
					For $k = $___aMouseInfo_Old[0] To $___aMouseInfo[0] Step 10 * ($___aMouseInfo_Old[0] > $___aMouseInfo[0] ? -1 : 1)
						If $VectorY = 0 Then
							$___aPosCurMem &= '|' & $k & '|' & $___aMouseInfo[1]
						Else
							$___aPosCurMem &= '|' & $k & '|' & $A * $k + $B
						EndIf
					Next
				Else
					$___aPosCurMem &= '|' & $___aMouseInfo[0] & '|' & $___aMouseInfo[1]
				EndIf
				_WinAPI_DrawLine($___WinAPI_hWDC, $___aMouseInfo[0], $___aMouseInfo[1], $___aMouseInfo_Old[0], $___aMouseInfo_Old[1])
				$___aMouseInfo_Old = $___aMouseInfo
			EndIf
		Until $___aMouseInfo[$iMouseEvent] = 0
		;-----------------------------------------------------------------------------------------------------
		_WinAPI_SelectObject($___WinAPI_hWDC, $___WinAPI_oSelect)
		_WinAPI_DeleteObject($___WinAPI_hPen)
		_WinAPI_ReleaseDC(0, $___WinAPI_hWDC)
		GUIDelete($___GDI_DrawGUI)
		Return $iClickNextButton
	EndFunc

	Func __IE_RecaptchaBox_CalculateRectClick($hGUI, $hIE, $___asPosCurMem, $iClickNextButton, $iDefaultReCaptDimensions = 4, $___offsetClick = 7, $___speedClick = 0)
		Local $aCaptcha_Measure = __IE_ReCaptchaBox_Measure($hGUI)
		If @error Or Not IsArray($aCaptcha_Measure) Then
			$aCaptcha_Measure = StringSplit($iDefaultReCaptDimensions = 3 ? '17,163,118,118,4,3,3,367' : '19,163,88,88,2,4,4,363', ',', 2)
		EndIf
		;-----------------------------------------------------------------------------------------------------
		Local $___aClickCurMem[8][8], $eCoordX, $eCoordY
		Local $__X_offsetClick = $aCaptcha_Measure[0] + $aCaptcha_Measure[2] / 2 - $___offsetClick
		Local $__Y_offsetClick = $aCaptcha_Measure[1] + $aCaptcha_Measure[3] / 2 - $___offsetClick
		$___asPosCurMem = StringSplit($___asPosCurMem, '|')
		If $___asPosCurMem[0] < 2 Then Return SetError(1)
		For $i = 1 To $___asPosCurMem[0] Step 2
			$eCoordX = Floor(($___asPosCurMem[$i + 0] - $aCaptcha_Measure[0]) / $aCaptcha_Measure[2])
			$eCoordY = Floor(($___asPosCurMem[$i + 1] - $aCaptcha_Measure[1]) / $aCaptcha_Measure[3])
			If $eCoordX < 0 Or $eCoordX >= $aCaptcha_Measure[5] Or $eCoordY < 0 Or $eCoordY >= $aCaptcha_Measure[6] Or $___aClickCurMem[$eCoordX][$eCoordY] Then ContinueLoop
			If $hIE Then
				__IE_MouseClick($hIE, $__X_offsetClick + $aCaptcha_Measure[2] * $eCoordX, $__Y_offsetClick + $aCaptcha_Measure[3] * $eCoordY - 25) ;-25 là do vị trí IE Obj so với GUI
			Else
				MouseClick('left', $__X_offsetClick + $aCaptcha_Measure[2] * $eCoordX, $__Y_offsetClick + $aCaptcha_Measure[3] * $eCoordY, 1, $___speedClick)
			EndIf
			$___aClickCurMem[$eCoordX][$eCoordY] = 1
		Next
		If $iClickNextButton Then
			Sleep(250)
			If $hIE Then
				__IE_MouseClick($hIE, 330, 540)
			Else
				MouseClick('left', 330, 560, 1, 0)
				MouseMove($___asPosCurMem[$___asPosCurMem[0] - 1], $___asPosCurMem[$___asPosCurMem[0]], 0)
			EndIf
		EndIf
	EndFunc

	Func __IE_ReCaptchaBox_Measure($hGUI)
		Local $hDC = _WinAPI_GetWindowDC($hGUI)
		If @error Then Return SetError(1, _WinAPI_ReleaseDC($hGUI, $hDC), 0)
		Local $iW = 404, $x = 10, $y = 200, $iStep = 0
		Local $iXCaptchaPiece = 0, $iYCaptchaPiece = 0, $iWCaptchaPiece = 0, $iHCaptchaPiece = 0, $iWCaptchaPic = 0, $iNumCaptchaPieceW = 0, $iNumCaptchaPieceH = 0
		For $x = 0 To $iW Step 2
			Select
				Case $iStep = 0
					If $x > 30 Then
						Return SetError(2, 0, 0)
					ElseIf __IE_MemoryReadPixel($x, $y, $hDC) == '0x7F7F7F' Then
						$iStep = 1
					EndIf
				Case $iStep = 1 And __IE_MemoryReadPixel($x, $y, $hDC) == '0xFFFFFF'
					$iStep = 2
				Case $iStep = 2 And __IE_MemoryReadPixel($x, $y, $hDC) <> '0xFFFFFF'
					$iStep = 3
					$iXCaptchaPiece = $x
					$x += 50
				Case $iStep = 3
					If __IE_MemoryReadPixel($x, $y, $hDC) == '0xFFFFFF' Then
						For $vertY = 180 To 220 Step 2
							If __IE_MemoryReadPixel($x, $vertY, $hDC) <> '0xFFFFFF' Then ExitLoop
						Next
						If $vertY = 222 Then
							$iWCaptchaPiece = $x - $iXCaptchaPiece
							$iStep = 4
						EndIf
					EndIf
				Case $iStep = 4
					For $y = 180 To 0 Step -2
						If __IE_MemoryReadPixel($x, $y, $hDC) == '0x4A90E2' Then
							$iYCaptchaPiece = $y + 7
							ExitLoop 2
						EndIf
					Next
			EndSelect
		Next
		_WinAPI_ReleaseDC($hGUI, $hDC)
		$iWCaptchaPic = $iW - ($iXCaptchaPiece - 1) * 2
		$iNumCaptchaPieceW = Floor($iWCaptchaPic / $iWCaptchaPiece)
		$iNumCaptchaPieceH = ($iNumCaptchaPieceW = 2 ? 4 : $iNumCaptchaPieceW)
		$iDistCaptchaPiece = Floor(($iWCaptchaPic - $iNumCaptchaPieceW * $iWCaptchaPiece) / $iNumCaptchaPieceW)
		$iHCaptchaPiece = Floor($iWCaptchaPic / $iNumCaptchaPieceH) - $iDistCaptchaPiece
		Local $aRet = [$iXCaptchaPiece, $iYCaptchaPiece, $iWCaptchaPiece, $iHCaptchaPiece, $iDistCaptchaPiece, $iNumCaptchaPieceW, $iNumCaptchaPieceH, $iWCaptchaPic]
		Return $aRet
	EndFunc

	Func __IE_MemoryReadPixel($__x, $__y, $hDC)
		Return BinaryMid(Binary(DllCall($dll_Gdi32, "int", "GetPixel", "int", $hDC, "int", $__x, "int", $__y)[0]), 1, 3)
	EndFunc

	;------------------------------------------------------------------------------------------------------

	; $ClearID = 1: History Only ; 2: Cookies Only ; 8: Temporary Internet Files Only ; 16: Form Data Only ; 32: Password History Only ; 255: Everything
	Func _IE_ClearMyTracks($vClearID = Default)
		If $vClearID = Default Then $vClearID = 2
		RunWait(@ComSpec & " /C " & "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess " & $vClearID, "", @SW_HIDE)
	EndFunc

	Func _IE_CheckCompatible($vCheckMode = True)
		;https://blogs.msdn.microsoft.com/patricka/2015/01/12/controlling-webbrowser-control-compatibility/
		Local $_Reg_BROWSER_EMULATION = '\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION'
		Local $_Reg_HKCU_BROWSER_EMULATION = 'HKCU\SOFTWARE' & $_Reg_BROWSER_EMULATION
		Local $_Reg_HKLM_BROWSER_EMULATION = 'HKLM\SOFTWARE' & $_Reg_BROWSER_EMULATION
		Local $_Reg_HKLMx64_BROWSER_EMULATION = 'HKLM\SOFTWARE\WOW6432Node' & $_Reg_BROWSER_EMULATION
		Local $_IE_Mode, $_AutoItExe = StringRegExp(@AutoItExe, '(?i)\\([^\\]+.exe)$', 1)[0]
		Local $_IE_Version = StringRegExp(FileGetVersion(@ProgramFilesDir & "\Internet Explorer\iexplore.exe"), '^\d+', 1)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('_IE_CheckCompatible', 'Không lấy được version của IE'), False)
		$_IE_Version = Number($_IE_Version[0])
		Switch $_IE_Version
			Case 8, 9
				$_IE_Mode = $_IE_Version * 1111
				_HttpRequest_ConsoleWrite('! IE' & $_IE_Version & ' có thể không tương thích với HTML mới sau này gây ra không thể tải trang bình hường (blank page)' & @CRLF)
			Case 10, 11
				$_IE_Mode = $_IE_Version * 1000 + 1
			Case Else
				_HttpRequest_ConsoleWrite( _
						'!!! Phiên bản Internet Explorer hiện tại trên máy bạn đã quá cũ (IE' & $_IE_Version & ').' & @CRLF & _
						'!!! Điều này có thể khiến một số trang có ReCaptcha không thể hiển thị được.' & @CRLF & _
						'!!! Nếu không nhúng ReCaptcha được, máy cần cài Win7 trở lên và IE version 10 hoặc 11.' & @CRLF)
				Return SetError(2, '', False)
		EndSwitch
		If $vCheckMode Then
			If RegRead($_Reg_HKCU_BROWSER_EMULATION, $_AutoItExe) <> $_IE_Mode Then RegWrite($_Reg_HKCU_BROWSER_EMULATION, $_AutoItExe, 'REG_DWORD', $_IE_Mode)
			If RegRead($_Reg_HKLM_BROWSER_EMULATION, $_AutoItExe) <> $_IE_Mode Then RegWrite($_Reg_HKLM_BROWSER_EMULATION, $_AutoItExe, 'REG_DWORD', $_IE_Mode)
			If @AutoItX64 And RegRead($_Reg_HKLMx64_BROWSER_EMULATION, $_AutoItExe) <> $_IE_Mode Then RegWrite($_Reg_HKLMx64_BROWSER_EMULATION, $_AutoItExe, 'REG_DWORD', $_IE_Mode)
		Else
			If RegRead($_Reg_HKCU_BROWSER_EMULATION, $_AutoItExe) <> $_IE_Mode Then RegDelete($_Reg_HKCU_BROWSER_EMULATION, $_AutoItExe)
			If RegRead($_Reg_HKLM_BROWSER_EMULATION, $_AutoItExe) <> $_IE_Mode Then RegDelete($_Reg_HKLM_BROWSER_EMULATION, $_AutoItExe)
			If @AutoItX64 And RegRead($_Reg_HKLMx64_BROWSER_EMULATION, $_AutoItExe) <> $_IE_Mode Then RegDelete($_Reg_HKLMx64_BROWSER_EMULATION, $_AutoItExe)
		EndIf
		Return True
	EndFunc

	;$vFuncCallback support hàm có tối đa 4 tham số
	Func _IE_GoogleBox($sUser, $sPassword, $sURL = Default, $vFuncCallback = '', $vDebug = False, $vTimeOut = Default, $vCheckCompatible = False)
		If $vCheckCompatible Then
			_IE_CheckCompatible(True)
			If @error Then Return SetError(-1, '', '')
		EndIf
		Local $sRet = __IE_Init_GoogleBox($sUser, $sPassword, $sURL, $vFuncCallback, $vDebug, $vTimeOut)
		Local $vErr = @error
		If $vCheckCompatible Then _IE_CheckCompatible(False)
		Return SetError($vErr, '', $sRet)
	EndFunc

	Func _IE_RecaptchaBox($sURL, $vAdvancedMode = Default, $iX_GUI = Default, $iY_GUI = Default, $vTimeOut = Default, $Custom_RegExp_GetDataSiteKey = Default, $vCheckCompatible = True)
		If $vTimeOut = Default Or $vTimeOut < 30000 Then $vTimeOut = 30000
		If $vAdvancedMode = Default Then $vAdvancedMode = False
		If $iX_GUI = Default Then $iX_GUI = (@DesktopWidth - 404) / 2
		If $iY_GUI = Default Then $iY_GUI = (@DesktopHeight - 607) / 2 - 50
		;------------------------------------------------------------------------------------------------------
		If $vCheckCompatible Then
			_IE_CheckCompatible(True)
			If @error Then Return SetError(-1, '', '')
		EndIf
		;------------------------------------------------------------------------------------------------------
		Local $___oldCursorMode = [Opt('MouseCoordMode', 0), Opt('MouseClickDelay', 0), Opt('MouseClickDownDelay', 0)]
		;------------------------------------------------------------------------------------------------------
		Local $ie_GUI_EmbededCaptcha = GUICreate("Recaptcha Box", 404, 607, $iX_GUI, $iY_GUI, 0x80000000, 0x8)
		Local $ie_GUI_SampleGetOffset = GUICreate("Recaptcha Box Sample", 0, 0, 0, 0, 0x80000000, 0x8 + 0x40, $ie_GUI_EmbededCaptcha)
		Local $___GUI_Offset = WinGetPos($ie_GUI_EmbededCaptcha)[0] - WinGetPos($ie_GUI_SampleGetOffset)[0]
		GUIDelete($ie_GUI_SampleGetOffset)
		;------------------------------------------------------------------------------------------------------
		Local $sRet = __IE_Init_RecaptchaBox($sURL, $vAdvancedMode, $ie_GUI_EmbededCaptcha, $___GUI_Offset, $Custom_RegExp_GetDataSiteKey, $vTimeOut)
		Local $vErr = @error
		;------------------------------------------------------------------------------------------------------
		$ie_GUI_EmbededCaptcha = GUIDelete($ie_GUI_EmbededCaptcha)
		ConsoleWrite(@CRLF)
		;------------------------------------------------------------------------------------------------------
		Opt('MouseCoordMode', $___oldCursorMode[0])
		Opt('MouseClickDelay', $___oldCursorMode[1])
		Opt('MouseClickDownDelay', $___oldCursorMode[2])
		_GDIPlus_Shutdown()
		If $vCheckCompatible Then _IE_CheckCompatible(False)
		Return SetError($vErr, '', $sRet)
	EndFunc

	Func _IE_NavigateEx($oIE, $sURL, $sCookie = '', $sUserAgent = '', $sProxy = '', $sProxyBypass = '', $iIEFlags = Default, $sIEPostData = '', $iIEHeaders = '', $iTimeout = 30000)
		If Not IsObj($oIE) Then Return SetError(1, __HttpRequest_ErrNotify('_IE_NavigateEx', 'IE Object rỗng'), False)
		;--------------------------------------------
		If $sCookie Then
			_IE_SetCookie($sURL, $sCookie)
			If @error Then Return SetError(2, False)
		EndIf
		;--------------------------------------------
		If $sProxy Then
			_IE_SetProxy($sProxy, $sProxyBypass)
			If @error Then Return SetError(3, False)
		EndIf
		;--------------------------------------------
		If $sUserAgent Then
			_IE_SetUserAgent($sUserAgent)
			If @error Then Return SetError(4, False)
		EndIf
		;--------------------------------------------
		$oIE.navigate2($sURL, $iIEFlags, Default, $sIEPostData, $iIEHeaders)
		If $iTimeout > -1 Then
			ConsoleWrite(@CRLF & '> _IE_NavigateEx' & @CRLF)
			_IE_LoadWait($oIE, $iTimeout)
		EndIf
		Return True
	EndFunc

	Func _IE_LoadWait($__oIE, $__iTimeout = 0)
		Local $__iTimerInit1 = TimerInit()
		ConsoleWrite('> _IE_WaitLoad ...')
		With $__oIE
			While .busy()
				If $__iTimeout And TimerDiff($__iTimerInit1) > $__iTimeout Then Return SetError(1, __HttpRequest_ErrNotify('_IE_LoadWait', 'LoadWait TimeOut #1'), False)
				ConsoleWrite('.')
				Sleep(50)
			WEnd
			$__iTimerInit1 = TimerDiff($__iTimerInit1)
			;--------------------------------------------------------------------------------
			Local $__iTimerInit2 = TimerInit()
			If IsObj(.document) Then
				While Not (String(.document.readyState) = "complete" Or .document.readyState = 4)
					If $__iTimeout And TimerDiff($__iTimerInit2) > $__iTimeout Then Return SetError(2, __HttpRequest_ErrNotify('_IE_LoadWait', 'LoadWait TimeOut #2'), False)
					ConsoleWrite('.')
					Sleep(50)
				WEnd
			Else
				While Not (String(.readyState) = "complete" Or .readyState = 4)
					If $__iTimeout And TimerDiff($__iTimerInit2) > $__iTimeout Then Return SetError(3, __HttpRequest_ErrNotify('_IE_LoadWait', 'LoadWait TimeOut #3'), False)
					ConsoleWrite('.')
					Sleep(50)
				WEnd
			EndIf
		EndWith
		$__iTimerInit2 = TimerDiff($__iTimerInit2)
		;--------------------------------------------------------------------------------
		ConsoleWrite(' (' & Round(($__iTimerInit1 + $__iTimerInit2) / 1000, 2) & 's)' & @CRLF)
	EndFunc

	Func _IE_CheckObjType($__oIE, $sType)
		If Not IsObj($__oIE) Then Return False
		Local $sName = String(ObjName($__oIE))
		Switch $sType
			Case "browserdom"
				If _IE_CheckObjType($__oIE, "documentcontainer") Then
					Return True
				ElseIf _IE_CheckObjType($__oIE, "document") Then
					Return True
				Else
					If _IE_CheckObjType($__oIE.document, "document") Then Return True
				EndIf
			Case "browser"
				If $sName = "IWebBrowser2" Or $sName = "IWebBrowser" Or $sName = "WebBrowser" Then Return True
			Case "window"
				If $sName = "HTMLWindow2" Then Return True
			Case "documentContainer"
				If _IE_CheckObjType($__oIE, "window") Or _IE_CheckObjType($__oIE, "browser") Then Return True
			Case "document"
				If $sName = "HTMLDocument" Then Return True
			Case "table"
				If $sName = "HTMLTable" Then Return True
			Case "form"
				If $sName = "HTMLFormElement" Then Return True
			Case "forminputelement"
				If ($sName = "HTMLInputElement") Or ($sName = "HTMLSelectElement") Or ($sName = "HTMLTextAreaElement") Then Return True
			Case "elementcollection"
				If ($sName = "HTMLElementCollection") Then Return True
			Case "formselectelement"
				If $sName = "HTMLSelectElement" Then Return True
			Case Else
				Return False
		EndSwitch
		Return False
	EndFunc

	Func _IE_GetCookie($sURL, $iBufferSize = 2048)
		If Not $dll_WinInet Then
			$dll_WinInet = DllOpen('wininet.dll')
			If @error Then Return SetError(2, __HttpRequest_ErrNotify('_IE_GetCookieEx', 'Không thể mở wininet.dll'), '')
			DllOpen('wininet.dll')
		EndIf
		Local $tSize = DllStructCreate("dword")
		DllStructSetData($tSize, 1, $iBufferSize)
		Local $tCookieData = DllStructCreate("wchar[" & $iBufferSize & "]")
		Local $avResult = DllCall($dll_WinInet, "int", "InternetGetCookieExW", 'wstr', $sURL, 'wstr', Null, "ptr", DllStructGetPtr($tCookieData), "ptr", DllStructGetPtr($tSize), "dword", 0x2000, "ptr", 0)
		If @error Then Return SetError(1, 0, "")
		If Not $avResult[0] Then Return SetError(1, DllStructGetData($tSize, 1), "")
		Return DllStructGetData($tCookieData, 1)
	EndFunc

	Func _IE_SetCookie($sURL, $iCookieData)
		;https://blogs.msdn.microsoft.com/ieinternals/2009/08/20/internet-explorer-cookie-internals-faq/
		If $iCookieData = '' Then Return SetError(1, __HttpRequest_ErrNotify('_IE_SetCookie', 'Không thể set Cookie vì tham số CookieData là rỗng'), '')
		If Not $dll_WinInet Then
			$dll_WinInet = DllOpen('wininet.dll')
			If @error Then Return SetError(2, __HttpRequest_ErrNotify('_IE_SetCookie', 'Không thể mở wininet.dll'), '')
			DllOpen('wininet.dll')
		EndIf
		Local $avResult, $cError = 0
		$iCookieData = StringSplit($iCookieData, ';')
		For $i = 1 To $iCookieData[0]
			If StringIsSpace($iCookieData[$i]) Then ContinueLoop
			If Not StringRegExp($iCookieData[$i], '^\h*?[^=]+\h*?=\h*?') Then ContinueLoop
			$avResult = DllCall($dll_WinInet, "int", "InternetSetCookieW", 'wstr', $sURL, "ptr", 0, 'wstr', $iCookieData[$i])
			If @error Then
				__HttpRequest_ErrNotify('_IE_SetCookie', 'Không thể nạp Cookie "' & $iCookieData[$i] & '" vào IE')
				$cError += 1
			EndIf
		Next
		If $cError Then Return SetError(1, '', False)
		Return True
	EndFunc

	Func _IE_SetProxy($sProxy, $sProxyBypass = "")
		Local $tBuff = DllStructCreate("dword;ptr;ptr")
		DllStructSetData($tBuff, 1, 3)
		Local $tProxy = DllStructCreate("char[" & (StringLen($sProxy) + 1) & "]")
		DllStructSetData($tProxy, 1, $sProxy)
		DllStructSetData($tBuff, 2, DllStructGetPtr($tProxy))
		Local $tProxyBypass = DllStructCreate("char[" & (StringLen($sProxyBypass) + 1) & "]")
		DllStructSetData($tProxyBypass, 1, $sProxyBypass)
		DllStructSetData($tBuff, 3, DllStructGetPtr($tProxyBypass))
		Local $avResult = DllCall("urlmon.dll", "long", "UrlMkSetSessionOption", "dword", 38, "ptr", DllStructGetPtr($tBuff), "dword", DllStructGetSize($tBuff), "dword", 0)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('_IE_SetProxy', 'Set Proxy cho IE thất bại'), False)
		Return True
	EndFunc

	Func _IE_SetUserAgent($sUserAgent)
		If Not StringRegExp($sUserAgent, '(?im)^User-Agent\s*?:') Then $sUserAgent = 'User-Agent: ' & $sUserAgent
		Local $sUserAgentLen = StringLen($sUserAgent)
		Local $tBuff = DllStructCreate("char[" & $sUserAgentLen & "]")
		DllStructSetData($tBuff, 1, $sUserAgent)
		Local $avResult = DllCall("urlmon.dll", "long", "UrlMkSetSessionOption", "dword", 0x10000001, "ptr", DllStructGetPtr($tBuff), "dword", $sUserAgentLen, "dword", 0)
		If @error Then Return SetError(1, __HttpRequest_ErrNotify('_IE_SetUserAgent', 'Set User-Agent cho IE thất bại'), False)
		Return True
	EndFunc
#EndRegion



#Region <CookieJar + CookieGlobal>
	Func _HttpRequest_CookieJarSet($sCookieJarFilePath)
		If $sCookieJarFilePath = '' Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_CookieJarSet', 'Đường dẫn tập tin lưu Cookie không tồn tại'), False)
		If Not StringRegExp($sCookieJarFilePath, '^\h*?\w{1,2}:\\') Then $sCookieJarFilePath = @ScriptDir & (StringLeft($sCookieJarFilePath, 1) = '\' ? '' : '\') & $sCookieJarFilePath
		If $sCookieJarFilePath <> $g___CookieJarPath Then
			$g___CookieJarPath = $sCookieJarFilePath
			_HttpRequest_CookieJarUpdateToFile()
		EndIf
		;-------------------------------------------------------------------------------------------
		If Not FileExists($g___CookieJarPath) Then FileOpen($g___CookieJarPath, 2 + 8 + 32)
		$g___CookieJarINI($g___CookieJarPath) = FileRead($g___CookieJarPath)
		If @error Or Not $g___CookieJarINI($g___CookieJarPath) Then $g___CookieJarINI($g___CookieJarPath) = ''
		Return True
	EndFunc

	Func _HttpRequest_CookieJarSearch($sURL)
		If $g___CookieJarPath = '' Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_CookieJarSearch', 'Vui lòng cài đặt _HttpRequest_CookieJarSet trước khi sử dụng hàm này'), '')
		If Not $sURL Or IsKeyword($sURL) Or $sURL == -1 Then Return $g___CookieJarINI($g___CookieJarPath)
		Local $aDomain = StringRegExp($g___CookieJarINI($g___CookieJarPath), '(?m)^\[([^\]]+)\]$', 3)
		If @error Then Return SetError(1, 0, '')
		Local $sCookie = ''
		For $i = 0 To UBound($aDomain) - 1
			If StringRegExp($sURL, '(?i)^https?:\/.*?' & $aDomain[$i] & '(?:\/|$)') Then
				$sCookie &= __CookieJar_Read($aDomain[$i])
			EndIf
		Next
		Return StringReplace($sCookie, @CRLF, '; ', 0, 1)
	EndFunc

	Func _HttpRequest_CookieJarDelete($iSection = '', $iKey = '')
		If $g___CookieJarPath == '' Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_CookieJarDelete', 'Vui lòng cài đặt _HttpRequest_CookieJarSet trước khi sử dụng hàm này'), '')
		__CookieJar_Delete($iSection, $iKey)
	EndFunc

	Func _HttpRequest_CookieJarUpdateToFile()
		If $g___CookieJarPath = '' Then Return SetError(1, __HttpRequest_ErrNotify('_HttpRequest_CookieJarUpdateToFile', 'Vui lòng cài đặt _HttpRequest_CookieJarSet trước khi sử dụng hàm này'), False)
		If Not $g___CookieJarINI($g___CookieJarPath) Then Return False
		Local $hFileOpen = FileOpen($g___CookieJarPath, 2 + 8 + 32)
		FileWrite($hFileOpen, $g___CookieJarINI($g___CookieJarPath))
		$hFileOpen = FileClose($hFileOpen)
		Return True
	EndFunc

	;-------------------------------------------------------------------------------------

	Func __CookieJar_Insert($sDomain, $iHeaders)
		If Not $g___CookieJarPath Or Not $iHeaders Then Return $iHeaders
		Local $aCookie = StringRegExp($iHeaders, '(?im)^Set-Cookie\h*:\h*([^=]+)=(?!deleted;)([^;]+)(?:.*?;\h*?domain=([^;\r\n]+))?()', 3)
		If @error Or Mod(UBound($aCookie), 4) Then Return SetError(1, '', $iHeaders)
		For $i = 0 To UBound($aCookie) - 1 Step 4
			If $aCookie[$i + 2] == '' Then $aCookie[$i + 2] = $sDomain
			__CookieJar_Write($aCookie[$i + 2], $aCookie[$i], $aCookie[$i + 1]) ;$aCookie[$i + 2] nhớ thêm proxy vào
		Next
		Return $iHeaders
	EndFunc

	Func __CookieJar_Read($iSection, $iKey = '', $vDefault = '')
		Local $sRegion = StringRegExp($g___CookieJarINI($g___CookieJarPath), '(?ims)^\Q[' & $iSection & ']\E$\R?(.*?)(?:\R?^\[[^\]]+\]$|\R?\z)', 1)
		If @error Then Return SetError(1, '', $vDefault)
		If $iKey == '' Then Return $sRegion[0]
		Local $sKeyValue = StringRegExp($sRegion[0], '(?im)^\Q' & $iKey & '\E=(.*)$', 1)
		If @error Then Return SetError(2, '', $vDefault)
		Return $sKeyValue[0]
	EndFunc

	Func __CookieJar_Write($iSection, $iKey, $iValue)
		If $iKey == '' Then Return SetError(1, '', False)
		Local $vKeyValueOld = __CookieJar_Read($iSection, $iKey, False)
		Switch @error
			Case 0 ;Đã có Section lẫn Key
				$g___CookieJarINI($g___CookieJarPath) = StringRegExpReplace($g___CookieJarINI($g___CookieJarPath), '(?ims)^(\Q[' & $iSection & ']\E$.*?\R^\Q' & $iKey & '=\E)\Q' & $vKeyValueOld & '\E$', '${1}' & $iValue, 1)
			Case 1 ;Section chưa được tạo
				$g___CookieJarINI($g___CookieJarPath) = '[' & $iSection & ']' & @CRLF & $iKey & '=' & $iValue & @CRLF & @CRLF & $g___CookieJarINI($g___CookieJarPath)
			Case 2 ;Có Section và không có Key
				$g___CookieJarINI($g___CookieJarPath) = StringRegExpReplace($g___CookieJarINI($g___CookieJarPath), '(?im)^(\Q[' & $iSection & ']\E)$', '${1}' & @CRLF & $iKey & '=' & $iValue)
		EndSwitch
		If @error Then Return SetError(2, '', False)
		Return True
	EndFunc

	Func __CookieJar_Delete($iSection, $iKey = '')
		If $iKey == '' Then
			$g___CookieJarINI($g___CookieJarPath) = StringRegExpReplace($g___CookieJarINI($g___CookieJarPath), '(?ims)^\Q[' & $iSection & ']\E$.*?\R(^\[[^\]]+\]$|\R?\z)', '${1}', 1)
		Else
			$g___CookieJarINI($g___CookieJarPath) = StringRegExpReplace($g___CookieJarINI($g___CookieJarPath), '(?ims)^(\Q[' & $iSection & ']\E$.*?)\R^\Q' & $iKey & '=\E.*?$', '${1}', 1)
		EndIf
		If @error Then Return SetError(1, '', False)
		Return True
	EndFunc

	;------------------------------------------------------------------------

	Func __CookieGlobal_Insert($sDomain, $sCookie)
		If Not $sCookie Then Return
		Local $aCookie = StringRegExp($sCookie, '(?<=^|;)\h*([^=]+)=\h*([^;]+)(?:;|$)', 3)
		If @error Or Mod(UBound($aCookie), 2) Then Return SetError(1, '', '')
		For $i = 0 To UBound($aCookie) - 1 Step 2
			If $aCookie[$i + 1] = 'deleted' Then
				__CookieGlobal_Delete($sDomain, $aCookie[$i])
			Else
				__CookieGlobal_Write($sDomain, $aCookie[$i], $aCookie[$i + 1])
			EndIf
		Next
	EndFunc

	Func __CookieGlobal_Search($sURL)
		Local $aDomain = StringRegExp($g___hCookie[$g___LastSession], '(?m)^\[([^\]]+)\]$', 3)
		If @error Then Return SetError(1, 0, '')
		Local $sCookie = ''
		For $i = 0 To UBound($aDomain) - 1
			If StringRegExp($sURL, '(?i)^https?:\/.*?' & $aDomain[$i] & '[^\/]*?(?:\/|$)') Then
				$sCookie &= __CookieGlobal_Read($aDomain[$i])
			EndIf
		Next
		Return StringReplace($sCookie, @CRLF, '; ', 0, 1)
	EndFunc

	Func __CookieGlobal_Read($iSection, $iKey = '', $vDefault = '')
		Local $sRegion = StringRegExp($g___hCookie[$g___LastSession], '(?ims)^\Q[' & $iSection & ']\E$\R?(.*?)(?:\R?^\[[^\]]+\]$|\R?\z)', 1)
		If @error Then Return SetError(1, '', $vDefault)
		If $iKey == '' Then Return $sRegion[0]
		Local $sKeyValue = StringRegExp($sRegion[0], '(?im)^\Q' & $iKey & '\E=(.*)$', 1)
		If @error Then Return SetError(2, '', $vDefault)
		Return $sKeyValue[0]
	EndFunc

	Func __CookieGlobal_Write($iSection, $iKey, $iValue)
		If $iKey == '' Then Return SetError(1, '', False)
		Local $vKeyValueOld = __CookieGlobal_Read($iSection, $iKey, False)
		Switch @error
			Case 0 ;Đã có Section lẫn Key
				$g___hCookie[$g___LastSession] = StringRegExpReplace($g___hCookie[$g___LastSession], '(?ims)^(\Q[' & $iSection & ']\E$.*?\R^\Q' & $iKey & '=\E)\Q' & $vKeyValueOld & '\E$', '${1}' & $iValue, 1)
			Case 1 ;Section chưa được tạo
				$g___hCookie[$g___LastSession] = '[' & $iSection & ']' & @CRLF & $iKey & '=' & $iValue & @CRLF & @CRLF & $g___hCookie[$g___LastSession]
			Case 2 ;Có Section và không có Key
				$g___hCookie[$g___LastSession] = StringRegExpReplace($g___hCookie[$g___LastSession], '(?im)^(\Q[' & $iSection & ']\E)$', '${1}' & @CRLF & $iKey & '=' & $iValue)
		EndSwitch
		If @error Then Return SetError(2, '', False)
		Return True
	EndFunc

	Func __CookieGlobal_Delete($iSection, $iKey = '')
		If $iKey == '' Then
			$g___hCookie[$g___LastSession] = StringRegExpReplace($g___hCookie[$g___LastSession], '(?ims)^\Q[' & $iSection & ']\E$.*?\R(^\[[^\]]+\]$|\R?\z)', '${1}', 1)
		Else
			$g___hCookie[$g___LastSession] = StringRegExpReplace($g___hCookie[$g___LastSession], '(?ims)^(\Q[' & $iSection & ']\E$.*?)\R^\Q' & $iKey & '=\E.*?$', '${1}', 1)
		EndIf
		If @error Then Return SetError(1, '', False)
		Return True
	EndFunc
#EndRegion



#Region Đang test
	Func _HttpRequest_FillDataForm($sURL, $sElement1, $sValue1, $sElement2 = '', $sValue2 = '', $sElement3 = '', $sValue3 = '', $sElement4 = '', $sValue4 = '', $sElement5 = '', $sValue5 = '', $sElement6 = '', $sValue6 = '', $sElement7 = '', $sValue7 = '', $sElement8 = '', $sValue8 = '', $sElement9 = '', $sValue9 = '', $sElement10 = '', $sValue10 = '', $sElement11 = '', $sValue11 = '', $sElement12 = '', $sValue12 = '')
		Local $sHTML = _HttpRequest(2, $sURL)
		If @error Or $sHTML == '' Then Return SetError(1, '', $sHTML)
		Local $sCookie = _GetCookie()
		If Not StringRegExp($sHTML, '(?i)<\h*?base .*?href\h*?=') Then
			$sHTML = '<base href="' & $sURL & '"/><script>var _b = document.getElementsByTagName("base")[0], _bH = "' & $sURL & '";if (_b && _b.href != _bH) _b.href = _bH;</script>' & $sHTML
		EndIf
		Local $aInput = StringRegExp($sHTML, '(?i)<\h*?(?:input|button) [^>]+>', 3)
		If @error Then Return SetError(2, '', $sHTML)
		Local $sElement, $iTypeInput
		For $j = 1 To Ceiling((@NumParams - 1) / 2)
			$sElement = Eval('sElement' & $j)
			If $sElement == '' Then ContinueLoop
			For $i = 0 To UBound($aInput) - 1
				If StringInStr($aInput[$i], $sElement, 0, 1) Then
					$iTypeInput = StringRegExp($aInput[$i], '(?i)type\h*?=\h*?["''](checkbox|button|submit)["'']', 1)
					If @error Then
						$iTypeInput = 'value'
					ElseIf $iTypeInput[0] = 'checkbox' Then
						$iTypeInput = 'checked'
					ElseIf $iTypeInput[0] = 'button' Or $iTypeInput[0] = 'submit' Then
;~ 					If Not StringRegExp($sHTML, '<script .*?src=["''][^"'']*?jquery[^"'']*?\.js">') Then $sHTML = '<script src="https://sso.garena.com/js/jquery-1.10.2.min.js"></script>' & @CRLF & $sHTML
;~ 					$sHTML &= @CRLF & '<script>function Jquery_xPath(STR_XPATH) {var xresult = document.evaluate(STR_XPATH, document, null, XPathResult.ANY_TYPE, null); var xnodes = [];var xres; while (xres = xresult.iterateNext()) {xnodes.push(xres);} return xnodes;}; $(Jquery_xPath(''//input[@' & $sElement & ']'')).click()</script>' ;$sElement phải ở dạng xPath
					Else
						$iTypeInput = 'value'
					EndIf
					$sHTML = StringReplace($sHTML, $aInput[$i], StringRegExpReplace($aInput[$i], '(?i)<\h*?(input|button) ', '<\1 ' & $iTypeInput & '="' & Eval('sValue' & $j) & '" ', 1), 1, 1)
				EndIf
			Next
		Next
		Local $aRet = [$sHTML, $sCookie]
		Return $aRet
	EndFunc
#EndRegion
