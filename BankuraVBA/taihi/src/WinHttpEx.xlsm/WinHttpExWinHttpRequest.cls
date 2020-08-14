VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpExWinHttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WinHttp.WinHttpRequestクラスのラップ・拡張クラス
'* [詳  細] WinHttp.WinHttpRequestクラスのWrapperとしての機能を提供する他、
'*          WinHttp.WinHttpRequestクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WinHttp.WinHttpRequestクラスについて。
' <https://docs.microsoft.com/en-us/windows/win32/winhttp/winhttprequest>
' HTTPプロトコルを介して他のHTTPサーバーに要求を送信する機能を提供します。
'
' [イベント]
' Event OnError(ErrorNumber As Long, ErrorDescription As String)
' Event OnResponseDataAvailable(Data() As Byte)
' Event OnResponseFinished()
' Event OnResponseStart(Status As Long, ContentType As String)
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WinHttpRequest
Private mOrigin As Object
Private mCharset As String

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* Option_ プロパティ
'*
'* [補  足]
'* ・Optionプロパティの代替。
'* ・Sets or retrieves a WinHTTP option value.
'*-----------------------------------------------------------------------------
Property Get Option_(whrOption As WinHttpRequestOption)
    Option_ = mOrigin.Option
End Property
Property Let Option_(whrOption As WinHttpRequestOption, arg)
    mOrigin.Option(whrOption) = arg
End Property

'*-----------------------------------------------------------------------------
'* ResponseBody プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get response body as a safearray of UI1
'*-----------------------------------------------------------------------------
Property Get ResponseBody() As Variant
    If IsObject(mOrigin.ResponseBody) Then
        Set ResponseBody = mOrigin.ResponseBody
    Else
        ResponseBody = mOrigin.ResponseBody
    End If
End Property

'*-----------------------------------------------------------------------------
'* ResponseStream プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get response body as a stream
'*-----------------------------------------------------------------------------
Property Get ResponseStream() As Variant
    If IsObject(mOrigin.ResponseStream) Then
        Set ResponseStream = mOrigin.ResponseStream
    Else
        ResponseStream = mOrigin.ResponseStream
    End If
End Property

'*-----------------------------------------------------------------------------
'* ResponseText プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get response body as a string
'*-----------------------------------------------------------------------------
Property Get ResponseText() As String
    ResponseText = mOrigin.ResponseText
End Property

'*-----------------------------------------------------------------------------
'* Status プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get HTTP status code
'*-----------------------------------------------------------------------------
Property Get Status() As Long
    Status = mOrigin.Status
End Property

'*-----------------------------------------------------------------------------
'* StatusText プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get HTTP status text
'*-----------------------------------------------------------------------------
Property Get StatusText() As String
    StatusText = mOrigin.StatusText
End Property

'*-----------------------------------------------------------------------------
'* <拡張>CharSet プロパティ
'*
'* [補  足]
'* ・Request/Responseの文字コードを指定する。
'*-----------------------------------------------------------------------------
Property Get CharSet() As String
    CharSet = mCharset
End Property
Property Let CharSet(arg As String)
    mCharset = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("WinHttp.WinHttpRequest.5.1")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Abort メソッド
'* [詳  細] Abort のラッパーメソッド。
'*          Abort an asynchronous operation in progress
'*
'******************************************************************************
Public Sub Abort()
    Call mOrigin.Abort
End Sub
    
'******************************************************************************
'* [概  要] GetAllResponseHeaders メソッド
'* [詳  細] GetAllResponseHeaders のラッパーメソッド。
'*          Get all HTTP response headers
'*
'* @return ResponseHeaders
'*         Receives the resulting header information.
'******************************************************************************
Public Function GetAllResponseHeaders() As String
    GetAllResponseHeaders = mOrigin.GetAllResponseHeaders
End Function
    
'******************************************************************************
'* [概  要] GetResponseHeader メソッド
'* [詳  細] GetResponseHeader のラッパーメソッド。
'*          Get HTTP response header
'*
'* @param Header ResponseHeader名
'*               Specifies the case-insensitive header name.
'* @return HTTPResponseHeader
'*         Receives the resulting header information.
'******************************************************************************
Public Function GetResponseHeader(Header As String) As String
    GetResponseHeader = mOrigin.GetResponseHeader(Header)
End Function
    
    
'******************************************************************************
'* [概  要] OpenConn メソッド
'* [詳  細] Open のラッパーメソッド。
'*          Open HTTP connection
'*
'* @param Method メソッド
'*               Specifies the HTTP verb used for the Open method, such as
'*               "GET" or "PUT". Always use uppercase as some servers ignore
'*               lowercase HTTP verbs.
'* @param Url    URL
'*               Specifies the name of the resource. This must be an absolute URL.
'* @param Async  省略可能。同期・非同期の指定。省略時は同期。
'*               Indicates whether to open in asynchronous mode.
'*                 False - Opens the HTTP connection in synchronous mode.
'*                         A call to Send does not return until WinHTTP has
'*                         completely received the response.
'*                 True  - Opens the HTTP connection in asynchronous mode.
'******************************************************************************
Public Sub OpenConn(method As String, Url As String, Optional Async As Variant = False)
    Call mOrigin.Open(method, Url, Async)
End Sub

'******************************************************************************
'* [概  要] Send メソッド
'* [詳  細] Send のラッパーメソッド。
'*          Send HTTP request
'*
'* @param Body  省略可能。Data to be sent to the server.
'******************************************************************************
Public Sub Send(Optional Body)
    If IsMissing(Body) Then
       Call mOrigin.Send
    Else
       Call mOrigin.Send(Body)
    End If
End Sub

'******************************************************************************
'* [概  要] SetAutoLogonPolicy メソッド
'* [詳  細] SetAutoLogonPolicy のラッパーメソッド。
'*          Specify if credentials should be sent automatically
'*
'* @param AutoLogonPolicy Specifies the current automatic logon policy.
'******************************************************************************
Public Sub SetAutoLogonPolicy(AutoLogonPolicy As WinHttpRequestAutoLogonPolicy)
    Call mOrigin.SetAutoLogonPolicy(AutoLogonPolicy)
End Sub
    
'******************************************************************************
'* [概  要] SetClientCertificate メソッド
'* [詳  細] SetClientCertificate のラッパーメソッド。
'*          Specify a client certificate
'*
'* @param ClientCertificate Specifies the location, certificate store, and s
'*                          ubject of a client certificate.
'******************************************************************************
Public Sub SetClientCertificate(ClientCertificate As String)
    Call mOrigin.SetClientCertificate(ClientCertificate)
End Sub

'******************************************************************************
'* [概  要] SetCredentials メソッド
'* [詳  細] SetCredentials のラッパーメソッド。
'*          Specify authentication credentials
'*
'* @param UserName ユーザ名
'*                 Specifies the user name for authentication.
'* @param Password パスワード
'*                 Specifies the password for authentication.
'*                 This parameter is ignored if UserName is NULL or missing.
'* @param Flags    フラグ
'*                 Specifies when IWinHttpRequest uses credentials.
'*
'******************************************************************************
Public Sub SetCredentials(UserName As String, Password As String, Flags As HTTPREQUEST_SETCREDENTIALS_FLAGS)
    Call mOrigin.SetCredentials(UserName, Password, Flags)
End Sub
    
'******************************************************************************
'* [概  要] SetProxy メソッド
'* [詳  細] SetProxy のラッパーメソッド。
'*          Specify proxy configuration
'*
'* @param ProxySetting The flags that control this method.
'* @param ProxyServer  Set to a proxy server string when ProxySetting equals
'*                     HTTPREQUEST_PROXYSETTING_PROXY.
'* @param BypassList   Set to a domain bypass list string when ProxySetting
'*                     equals HTTPREQUEST_PROXYSETTING_PROXY.
'******************************************************************************
Public Sub SetProxy(ProxySetting As HTTPREQUEST_PROXY_SETTING, Optional ProxyServer, Optional BypassList)
    If IsMissing(ProxyServer) Then
        If IsMissing(BypassList) Then
            Call mOrigin.SetProxy(ProxySetting)
        Else
            Call mOrigin.SetProxy(ProxySetting, , BypassList)
        End If
    ElseIf IsMissing(BypassList) Then
        Call mOrigin.SetProxy(ProxySetting, ProxyServer)
    Else
        Call mOrigin.SetProxy(ProxySetting, ProxyServer, BypassList)
    End If
End Sub

'******************************************************************************
'* [概  要] SetRequestHeader メソッド
'* [詳  細] SetRequestHeader のラッパーメソッド。
'*          Add HTTP request header
'*
'* @param Header Header名
'*               Specifies the name of the header to be set, for example,
'*               "depth". This parameter should not contain a colon and must be
'*               the actual text of the HTTP header.
'* @param Value  項目値
'*               Specifies the value of the header, for example, "infinity".
'*
'******************************************************************************
Public Sub SetRequestHeader(Header As String, Value As String)
    Call mOrigin.SetRequestHeader(Header, Value)
End Sub
    
'******************************************************************************
'* [概  要] SetTimeouts メソッド
'* [詳  細] SetTimeouts のラッパーメソッド。
'*          Specify timeout settings (in milliseconds)
'*
'* @param ResolveTimeout リゾルヴタイムアウト値
'*                       Time-out value applied when resolving a host name
'*                       (such as www.microsoft.com) to an IP address (such
'*                       as 192.168.131.199), in milliseconds.
'*                       The default value is zero, meaning no time-out (infinite).
'*                       If DNS timeout is specified using NAME_RESOLUTION_TIMEOUT,
'*                       there is an overhead of one thread per request.
'* @param ConnectTimeout 接続タイムアウト値
'*                       Time-out value applied when establishing a communication
'*                       socket with the target server, in milliseconds.
'*                       The default value is 60,000 (60 seconds).
'* @param SendTimeout    送信タイムアウト値
'*                       Time-out value applied when sending an individual packet
'*                       of request data on the communication socket to the target
'*                       server, in milliseconds. A large request sent to an HTTP
'*                       server are normally be broken up into multiple packets;
'*                       the send time-out applies to sending each packet
'*                       individually. The default value is 30,000 (30 seconds).
'* @param ReceiveTimeout 受信タイムアウト値
'*                       Time-out value applied when receiving a packet of
'*                       response data from the target server, in milliseconds.
'*                       Large responses are be broken up into multiple packets;
'*                       the receive time-out applies to fetching each packet of
'*                       data off the socket.
'*                       The default value is 30,000 (30 seconds).
'******************************************************************************
Public Sub SetTimeouts(ResolveTimeout As Long, ConnectTimeout As Long, SendTimeout As Long, ReceiveTimeout As Long)
    Call mOrigin.SetTimeouts(ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout)
End Sub

'******************************************************************************
'* [概  要] WaitForResponse メソッド
'* [詳  細] WaitForResponse のラッパーメソッド。
'*          Wait for asynchronous send to complete, with optional timeout (in seconds)
'*
'* @param Timeout 省略可能。タイムアウト値。
'*                Time-out value, in seconds. Default time-out is infinite.
'*                To explicitly set time-out to infinite, use the value -1.
'* @return Receives one of the following values.
'*          True  - A response has been received.
'*          False - The specified time-out period was exceeded.
'******************************************************************************
Public Function WaitForResponse(Optional Timeout) As Boolean
    If IsMissing(Timeout) Then
       WaitForResponse = mOrigin.WaitForResponse()
    Else
       WaitForResponse = mOrigin.WaitForResponse(Timeout)
    End If
End Function


'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] ReNew メソッド
'* [詳  細] WinHttpRequestオブジェクトを再生成する。
'*
'******************************************************************************
Public Sub ReNew()
    Set mOrigin = Nothing
    Set mOrigin = CreateObject("WinHttp.WinHttpRequest.5.1")
End Sub

'******************************************************************************
'* [概  要] Post メソッド
'* [詳  細] Post送信を行う。
'*
'* @param ResponseDto Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function Post(req As RequestDto) As ResponseDto
    req.method = "POST"
    Set Post = Transmit(req)
End Function

'******************************************************************************
'* [概  要] PutReq メソッド
'* [詳  細] Put送信を行う。
'*
'* @param ResponseDto Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function PutReq(req As RequestDto) As ResponseDto
    req.method = "PUT"
    Set PutReq = Transmit(req)
End Function

'******************************************************************************
'* [概  要] Patch メソッド
'* [詳  細] Patch送信を行う。
'*
'* @param ResponseDto Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function Patch(req As RequestDto) As ResponseDto
    req.method = "PATCH"
    Set Patch = Transmit(req)
End Function

'******************************************************************************
'* [概  要] GetReq メソッド
'* [詳  細] Get送信を行う。
'*
'* @param ResponseDto Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function GetReq(req As RequestDto) As ResponseDto
    req.method = "GET"
    Set GetReq = Transmit(req)
End Function

'******************************************************************************
'* [概  要] Delete メソッド
'* [詳  細] Delete送信を行う。
'*
'* @param ResponseDto Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function Delete(req As RequestDto) As ResponseDto
    req.method = "DELETE"
    Set Delete = Transmit(req)
End Function

'******************************************************************************
'* [概  要] Head メソッド
'* [詳  細] Head送信を行う。
'*
'* @param ResponseDto Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function Head(req As RequestDto) As ResponseDto
    req.method = "HEAD"
    Set Head = Transmit(req)
End Function

'******************************************************************************
'* [概  要] OptionsReq メソッド
'* [詳  細] Options送信を行う。
'*
'* @param ResponseDto Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function OptionsReq(req As RequestDto) As ResponseDto
    req.method = "OPTIONS"
    Set OptionsReq = Transmit(req)
End Function

'******************************************************************************
'* [概  要] Transmit メソッド
'* [詳  細] 簡単なHttp通信を行う。
'*
'* @param ResponseDto Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function Transmit(req As RequestDto, Optional saveOriginalBody As Boolean = False) As ResponseDto
    Me.OpenConn req.method, req.Url, False
    SetRequestHeaders req
    If req.method = "GET" Or req.method = "DELETE" Or req.method = "HEAD" Or req.method = "OPTIONS" Then
        Me.Send
    Else
        Me.Send req.Body
    End If
    Set Transmit = SetResponseDto(saveOriginalBody)
End Function

'******************************************************************************
'* [概  要] SetRequestHeaders メソッド
'* [詳  細] RequestDtoのHeader情報を設定する。
'*
'* @param Request情報
'*
'******************************************************************************
Public Sub SetRequestHeaders(req As RequestDto)
    Dim v
    For Each v In req.Headers.keys
        Me.SetRequestHeader CStr(v), req.Headers(v)
    Next
End Sub

'******************************************************************************
'* [概  要] SetResponseDto メソッド
'* [詳  細] Request送信結果を ResponseDto に設定する。
'*
'* @param  saveOriginalBody 省略可能。BodyをText化せずそのまま保持するか指定。
'* @return Response情報
'*
'******************************************************************************
Public Function SetResponseDto(Optional saveOriginalBody As Boolean = False) As ResponseDto
    Dim Status As String:     Status = Me.Status
    Dim statusTxt As String:  statusTxt = Me.StatusText
    Dim resHeaders As String: resHeaders = Me.GetAllResponseHeaders
    Dim encode As String: encode = IIf(mCharset <> "", mCharset, "utf-8")
    
    Dim resBody As Variant
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adReadAll = -1
    
    If saveOriginalBody Then
        resBody = Me.ResponseBody
    Else
        With CreateObject("ADODB.Stream")
            .Type = adTypeBinary
            .Open
            .Write Me.ResponseBody
            .Position = 0
            .Type = adTypeText
            .CharSet = encode
            resBody = .ReadText(adReadAll)
            .Close
        End With
    End If
  
    ' レスポンス設定
    Dim res As ResponseDto: Set res = New ResponseDto
    res.StatusCd = Status
    res.StatusText = statusTxt
    res.Body = resBody
    res.Headers = resHeaders
    
    Set SetResponseDto = res
End Function

'******************************************************************************
'* [概  要] SendRequest メソッド
'* [詳  細] Httpリクエストを送信する。
'*
'* @param method         メソッド（GET,POST,PUT,DELETE,HEAD,OPTIONS,PATCH）
'* @param url            アクセスするURL（例：http://localhost:8080/api/v2/test）
'* @param reqHdrParams() Headerパラメータ(2次元配列)
'*                       第2インデックス … 0:name, 1:value
'* @param reqBody        Body
'* @param lTimeout       タイムアウト時間（ミリ秒）　※任意
'* @param resEncode      レスポンスエンコード　※任意
'* @return Response情報
'*
'******************************************************************************
Public Function SendRequest(method As String, Url As String, reqHdrParams() As String, reqBody As String, _
                            Optional lTimeout As Long = 0, _
                            Optional resEncode As String = "utf-8") As ResponseDto
    ' HTTP リクエスト
    Me.OpenConn method, Url, False

    ' リクエストヘッダ設定
    Dim i As Long
    If GetDimension(reqHdrParams) <> 0 Then
        For i = 0 To UBound(reqHdrParams, 1)
            Me.SetRequestHeader reqHdrParams(i, 0), reqHdrParams(i, 1)
        Next i
    End If

    ' タイムアウト設定
    If lTimeout > 0 Then
        Me.SetTimeouts lTimeout, lTimeout, lTimeout, lTimeout
    End If
    
    ' リクエスト送信
    If method = "GET" Or method = "DELETE" Or method = "HEAD" Or method = "OPTIONS" Then
        Me.Send
    Else
        Me.Send reqBody
    End If

    Set SendRequest = SetResponseDto
End Function

'******************************************************************************
'* [概  要] DownloadFile メソッド
'* [詳  細] ファイルをダウンロードする。
'*
'* @param  req Request情報
'* @param  saveFilePath ダウンロードファイル保存先フルパス
'* @return Response情報
'*
'******************************************************************************
Public Function DownloadFile(req As RequestDto, saveFilePath As String) As ResponseDto
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2

    Set DownloadFile = Transmit(req, True)
    
    If Me.Status >= 200 And Me.Status < 400 Then
        With CreateObject("ADODB.Stream")
            .Type = adTypeBinary
            .Open
            .Write Me.ResponseBody
            .SaveToFile saveFilePath, adSaveCreateOverWrite
            .Close
        End With
    End If
End Function

'******************************************************************************
'* [概  要] SetEnableRedirects メソッド
'* [詳  細] リダイレクト有効/無効を設定する。
'*
'* @param enable リダイレクト有効/無効（True:有効, False:無効）
'*
'******************************************************************************
Public Sub SetEnableRedirects(enable As Boolean)
    Me.Option_(WinHttpRequestOption_EnableRedirects) = enable
End Sub


'******************************************************************************
'* [概  要] 次元取得処理。
'* [詳  細] 指定した配列が何次元の配列か判定する。
'*
'* @param targetArray 対象となる配列
'* @return 次元数
'******************************************************************************
Private Function GetDimension(targetArray As Variant) As Long
    Dim dimention As Long
    Dim tmp As Long
    
    dimention = 1
    On Error Resume Next
    While Err.Number = 0
        tmp = UBound(targetArray, dimention)
        dimention = dimention + 1
    Wend
    GetDimension = dimention - 2
End Function

'******************************************************************************
'* [概  要] URLエンコード処理。
'* [詳  細] URLエンコードを行う。
'*          特にScriptControlも関数も使用せず、自前でエンコードしているので
'*          どのExcelバージョン・32bit/64bitでも使用可能。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Private Function EncodeUrl(target As String) As String
    Dim buf() As Byte, s As String, i As Long
    
    With CreateObject("ADODB.Stream")
        .Mode = 3 'adModeReadWrite
        .Open
        .Type = 2 'adTypeText
        .CharSet = "UTF-8"
        .WriteText target
        
        .Position = 0
        .Type = 1 'adTypeBinary
        .Position = 3 'BOM飛ばし
        buf = .Read
        .Close
    End With

    For i = 0 To UBound(buf)
        Dim flg As Boolean: flg = False
        Select Case buf(i)
            Case 45, 46, 95, 126 '-._~
                flg = True
            Case 48 To 57 '0-9
                flg = True
            Case 65 To 90 'A-Z
                flg = True
            Case 97 To 122 'a-z
                flg = True
        End Select
        If flg Then
            s = s & Chr(buf(i))
        Else
            s = s & "%" & Hex(buf(i))
        End If
    Next
    EncodeUrl = s
End Function
