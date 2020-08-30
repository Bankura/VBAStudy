VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] HttpRequestクラス
'* [詳  細] Request情報のデータクラス。
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mMethod As String
Private mUrl As String
Private mBody As String
'Private mHeaders As Dictionary
'Private mParams As Dictionary
Private mHeaders As Object
Private mParams As Object

'******************************************************************************
'* プロパティ定義
'******************************************************************************
' Method
Public Property Get method() As String
    method = mMethod
End Property
Public Property Let method(arg As String)
    mMethod = arg
End Property

' URL
Public Property Get Url() As String
    If mParams.Count > 0 Then
        Url = mUrl & "?" & GetRequestParamsString
    Else
        Url = mUrl
    End If
End Property
Public Property Let Url(arg As String)
    mUrl = arg
End Property

' Body
Public Property Get Body() As String
    Body = mBody
End Property
Public Property Let Body(arg As String)
    mBody = arg
End Property

' Headers
Public Property Get Headers() As Object
    Set Headers = mHeaders
End Property
Public Property Set Headers(arg As Object)
    Set mHeaders = arg
End Property

' Params
Public Property Get Params() As Object
    Set Headers = mParams
End Property
Public Property Set Params(arg As Object)
    Set mParams = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mHeaders = CreateObject("Scripting.Dictionary")
    Set mParams = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    Set mHeaders = Nothing
    Set mParams = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] SetRequestHeader メソッド
'* [詳  細] リクエストヘッダを設定する。
'*
'* @param headerName  Header名
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetRequestHeader(headerName As String, Optional headerValue As String = "")
    If headerValue = "" Then
        mHeaders.Remove headerName
    Else
        mHeaders.Add headerName, headerValue
    End If
End Sub

'******************************************************************************
'* [概  要] SetRequestParam メソッド
'* [詳  細] クエリパラメータを設定する。
'*
'* @param paramName  パラメータ名
'* @param paramValue 項目値
'*
'******************************************************************************
Public Sub SetRequestParam(paramName As String, Optional paramValue As String = "")
    If paramValue = "" Then
        mParams.Remove paramName
    Else
        mParams.Add paramName, paramValue
    End If
End Sub

'******************************************************************************
'* [概  要] SetRequestParam メソッド
'* [詳  細] クエリパラメータを設定する。
'*
'* @param paramName  パラメータ名
'* @param paramValue 項目値
'*
'******************************************************************************
Public Function GetRequestParamsString() As String
    Dim tmp As String, i As Long: i = 1
    Dim vKey
    Dim varKey: varKey = mParams.keys
    For Each vKey In varKey
        If i = 1 Then
            tmp = vKey & "=" & EncodeUrl(mParams(vKey))
        Else
            tmp = tmp & "&" & vKey & "=" & EncodeUrl(mParams(vKey))
        End If
        i = i + 1
    Next
    GetRequestParamsString = tmp
End Function

'******************************************************************************
'* [概  要] SetReqHeaderContentType メソッド
'* [詳  細] リクエストヘッダ"Content-Type"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderContentType(headerValue As String)
    SetRequestHeader "Content-Type", headerValue
End Sub

'******************************************************************************
'* [概  要] SetReqHeaderConnection メソッド
'* [詳  細] リクエストヘッダ"Connection"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderConnection(headerValue As String)
    SetRequestHeader "Connection", headerValue
End Sub

'******************************************************************************
'* [概  要] SetReqHeaderCookie メソッド
'* [詳  細] リクエストヘッダ"Cookie"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderCookie(headerValue As String)
    SetRequestHeader "Cookie", headerValue
End Sub

'******************************************************************************
'* [概  要] SetReqHeaderAccept メソッド
'* [詳  細] リクエストヘッダ"Accept"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderAccept(headerValue As String)
    SetRequestHeader "Accept", headerValue
End Sub

'******************************************************************************
'* [概  要] SetReqHeaderUserAgent メソッド
'* [詳  細] リクエストヘッダ"User-Agent"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderUserAgent(headerValue As String)
    SetRequestHeader "User-Agent", headerValue
End Sub

'******************************************************************************
'* [概  要] SetReqHeaderReferer メソッド
'* [詳  細] リクエストヘッダ"Referer"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderReferer(headerValue As String)
    SetRequestHeader "Referer", headerValue
End Sub

'******************************************************************************
'* [概  要] SetReqHeaderAcceptLanguage メソッド
'* [詳  細] リクエストヘッダ"Accept-Language"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderAcceptLanguage(headerValue As String)
    SetRequestHeader "Accept-Language", headerValue
End Sub

'******************************************************************************
'* [概  要] SetReqHeaderAcceptEncoding メソッド
'* [詳  細] リクエストヘッダ"Accept-Encoding"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderAcceptEncoding(headerValue As String)
    SetRequestHeader "Accept-Encoding", headerValue
End Sub

'******************************************************************************
'* [概  要] SetReqHeaderHost メソッド
'* [詳  細] リクエストヘッダ"Host"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderHost(headerValue As String)
    SetRequestHeader "Host", headerValue
End Sub

'******************************************************************************
'* [概  要] SetReqHeaderCacheControl メソッド
'* [詳  細] リクエストヘッダ"Cache-Control"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderCacheControl(headerValue As String)
    SetRequestHeader "Cache-Control", headerValue
End Sub

'******************************************************************************
'* [概  要] SetReqHeaderAuthorization メソッド
'* [詳  細] リクエストヘッダ"Authorization"を設定する。
'*
'* @param headerValue 項目値
'*
'******************************************************************************
Public Sub SetReqHeaderAuthorization(headerValue As String)
    SetRequestHeader "Authorization", headerValue
End Sub

'******************************************************************************
'* [概  要] URLエンコード処理。
'* [詳  細] URLエンコードを行う。
'*          特にScriptControlも関数も使用せず、自前でエンコードしているので
'*          どのExcelバージョン・32bit/64bitでも使用可能。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Private Function EncodeUrl(Target As String) As String
    Dim buf() As Byte, s As String, i As Long
    
    With CreateObject("ADODB.Stream")
        .Mode = 3 'adModeReadWrite
        .Open
        .Type = 2 'adTypeText
        .charSet = "UTF-8"
        .WriteText Target
        
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
