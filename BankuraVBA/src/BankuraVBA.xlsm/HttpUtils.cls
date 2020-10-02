VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Http通信ユーティリティクラス
'* [詳  細] Http通信に関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As LongPtr) As Long

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mWinHttp As Object

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
End Sub
Private Sub Class_Terminate()
    Set mWinHttp = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************


'******************************************************************************
' [関数名] DownloadFileByUrl
' [説　明] 指定したURLからファイルをダウンロードし、指定した保存先（ファイル）
'          に保存する。
' [引　数] strFileUrl ダウンロード対象のURL
'          strSavePath 保存先ファイル名（フルパス）
' [戻り値] Boolean 処理結果（True:正常 False：異常）
'******************************************************************************
Public Function DownloadFileByUrl(strFileUrl As String, strSavePath As String) As Boolean

    Dim lngCacheDelRs As Long
    Dim lngDlRs As Long

    ' キャッシュクリア
    lngCacheDelRs = DeleteUrlCacheEntry(strFileUrl)
    
    ' URLからファイルをダウンロードして保存
    lngDlRs = URLDownloadToFile(0, strFileUrl, strSavePath, 0, 0)

    If lngDlRs <> 0 Then
        DownloadFileByUrl = False
        Exit Function
    End If
    DownloadFileByUrl = True
End Function

'******************************************************************************
'* [概  要] ReNewWinHttp メソッド
'* [詳  細] WinHttpRequestオブジェクトを再生成する。
'*
'******************************************************************************
Private Sub ReNewWinHttp()
    Set mWinHttp = Nothing
    Set mWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
End Sub

'******************************************************************************
'* [概  要] Post メソッド
'* [詳  細] Post送信を行う。
'*
'* @param HttpRequest Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function Post(req As HttpRequest) As HttpResponse
    req.method = "POST"
    Set Post = Transmit(req)
End Function

'******************************************************************************
'* [概  要] PutReq メソッド
'* [詳  細] Put送信を行う。
'*
'* @param HttpRequest Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function PutReq(req As HttpRequest) As HttpResponse
    req.method = "PUT"
    Set PutReq = Transmit(req)
End Function

'******************************************************************************
'* [概  要] Patch メソッド
'* [詳  細] Patch送信を行う。
'*
'* @param HttpRequest Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function Patch(req As HttpRequest) As HttpResponse
    req.method = "PATCH"
    Set Patch = Transmit(req)
End Function

'******************************************************************************
'* [概  要] GetReq メソッド
'* [詳  細] Get送信を行う。
'*
'* @param HttpRequest Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function GetReq(req As HttpRequest) As HttpResponse
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
Public Function Delete(req As HttpRequest) As HttpResponse
    req.method = "DELETE"
    Set Delete = Transmit(req)
End Function

'******************************************************************************
'* [概  要] Head メソッド
'* [詳  細] Head送信を行う。
'*
'* @param HttpRequest Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function Head(req As HttpRequest) As HttpResponse
    req.method = "HEAD"
    Set Head = Transmit(req)
End Function

'******************************************************************************
'* [概  要] OptionsReq メソッド
'* [詳  細] Options送信を行う。
'*
'* @param HttpRequest Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function OptionsReq(req As HttpRequest) As HttpResponse
    req.method = "OPTIONS"
    Set OptionsReq = Transmit(req)
End Function

'******************************************************************************
'* [概  要] Transmit メソッド
'* [詳  細] 簡単なHttp通信を行う。
'*
'* @param HttpRequest Request情報
'* @return Response情報
'*
'******************************************************************************
Public Function Transmit(req As HttpRequest, Optional saveOriginalBody As Boolean = False) As HttpResponse
    Call ReNewWinHttp
    mWinHttp.Open req.method, req.Url, False
    Call SetRequestHeaders(req)
    If req.method = "GET" Or req.method = "DELETE" Or req.method = "HEAD" Or req.method = "OPTIONS" Then
        mWinHttp.Send
    Else
        mWinHttp.Send req.Body
    End If
    Set Transmit = SetHttpResponse(, saveOriginalBody)
End Function

'******************************************************************************
'* [概  要] SetRequestHeaders メソッド
'* [詳  細] HttpRequestのHeader情報を設定する。
'*
'* @param Request情報
'*
'******************************************************************************
Public Sub SetRequestHeaders(req As HttpRequest)
    Dim v
    For Each v In req.Headers.keys
        mWinHttp.SetRequestHeader CStr(v), req.Headers(v)
    Next
End Sub

'******************************************************************************
'* [概  要] SetHttpResponse メソッド
'* [詳  細] Request送信結果を HttpResponse に設定する。
'*
'* @param  encode 省略可能。Responseの文字コード。
'* @param  saveOriginalBody 省略可能。BodyをText化せずそのまま保持するか指定。
'* @return Response情報
'*
'******************************************************************************
Public Function SetHttpResponse(Optional encode As String = "utf-8", Optional saveOriginalBody As Boolean = False) As HttpResponse
    Dim Status As String:     Status = mWinHttp.Status
    Dim statusTxt As String:  statusTxt = mWinHttp.StatusText
    Dim resHeaders As String: resHeaders = mWinHttp.GetAllResponseHeaders
    
    Dim resBody As Variant
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adReadAll = -1
    
    If saveOriginalBody Then
        resBody = mWinHttp.ResponseBody
    Else
        With CreateObject("ADODB.Stream")
            .Type = adTypeBinary
            .Open
            .Write mWinHttp.ResponseBody
            .Position = 0
            .Type = adTypeText
            .charSet = encode
            resBody = .ReadText(adReadAll)
            .Close
        End With
    End If
  
    ' レスポンス設定
    Dim res As HttpResponse: Set res = New HttpResponse
    res.StatusCd = Status
    res.StatusText = statusTxt
    res.Body = resBody
    res.Headers = resHeaders
    
    Set SetHttpResponse = res
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
                            Optional resEncode As String = "utf-8") As HttpResponse
    Call ReNewWinHttp
    
    ' HTTP リクエスト
    mWinHttp.Open method, Url, False

    ' リクエストヘッダ設定
    Dim i As Long
    If ArrayUtils.GetDimension(reqHdrParams) <> 0 Then
        For i = 0 To UBound(reqHdrParams, 1)
            mWinHttp.SetRequestHeader reqHdrParams(i, 0), reqHdrParams(i, 1)
        Next i
    End If

    ' タイムアウト設定
    If lTimeout > 0 Then
        mWinHttp.SetTimeouts lTimeout, lTimeout, lTimeout, lTimeout
    End If
    
    ' リクエスト送信
    If method = "GET" Or method = "DELETE" Or method = "HEAD" Or method = "OPTIONS" Then
        mWinHttp.Send
    Else
        mWinHttp.Send reqBody
    End If

    Set SendRequest = SetHttpResponse(resEncode)
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
Public Function DownloadFile(req As HttpRequest, saveFilePath As String) As HttpResponse
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2

    Set DownloadFile = Transmit(req, True)
    
    If mWinHttp.Status >= 200 And mWinHttp.Status < 400 Then
        With CreateObject("ADODB.Stream")
            .Type = adTypeBinary
            .Open
            .Write mWinHttp.ResponseBody
            .SaveToFile saveFilePath, adSaveCreateOverWrite
            .Close
        End With
    End If
End Function

