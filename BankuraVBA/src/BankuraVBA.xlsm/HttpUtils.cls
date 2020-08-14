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
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mWinApi As WinAPI

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mWinApi = Base.GetWinAPI()
End Sub
Private Sub Class_Terminate()
    Set mWinApi = Nothing
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
    lngCacheDelRs = mWinApi.DeleteUrlCacheEntry(strFileUrl)
    
    ' URLからファイルをダウンロードして保存
    lngDlRs = mWinApi.URLDownloadToFile(0, strFileUrl, strSavePath, 0, 0)

    If lngDlRs <> 0 Then
        DownloadFileByUrl = False
        Exit Function
    End If
    DownloadFileByUrl = True
End Function

'******************************************************************************
'* [概  要] Httpリクエストを送信し、レスポンスを取得・返却する。
'* [詳  細] 「WinHttp.WinHttpRequest.5.1」を使用して、Httpリクエストを送信
'*          する。
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
Public Function SendRequest(Method As String, Url As String, reqHdrParams() As String, reqBody As String, _
                            Optional lTimeout As Long = 0, _
                            Optional resEncode As String = "utf-8") As HttpResponse
    Dim winHttp As Variant
    Set winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' HTTP リクエスト
    winHttp.Open Method, Url, False
    
    ' リクエストヘッダ設定
    Dim i As Long
    If ArrayUtils.GetDimension(reqHdrParams) <> 0 Then
        For i = 0 To UBound(reqHdrParams, 1)
            winHttp.SetRequestHeader reqHdrParams(i, 0), reqHdrParams(i, 1)
        Next i
    End If

    ' タイムアウト設定
    If lTimeout > 0 Then
        winHttp.SetTimeouts lTimeout, lTimeout, lTimeout, lTimeout
    End If
    
    ' リクエスト送信
    If Method = "GET" Or Method = "DELETE" Or Method = "HEAD" Or Method = "OPTIONS" Then
        winHttp.Send
    Else
        winHttp.Send reqBody
    End If
    
    ' ステータスとレスポンスボディの取得
    Dim status As String:     status = winHttp.status
    Dim statusTxt As String:  statusTxt = winHttp.StatusText
    Dim resHeaders As String: resHeaders = winHttp.GetAllResponseHeaders
    Dim resBody As String
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adReadAll = -1
    With CreateObject("ADODB.Stream")
        .Type = adTypeBinary
        .Open
        .Write winHttp.ResponseBody
        .Position = 0
        .Type = adTypeText
        .charset = resEncode
        resBody = .ReadText(adReadAll)
        .Close
    End With
  
    ' レスポンス設定
    Dim res As HttpResponse: Set res = New HttpResponse
    res.StatusCd = status
    res.StatusText = statusTxt
    res.Body = resBody
    res.Headers = resHeaders
    
    Set SendRequest = res

End Function

'******************************************************************************
'* [概  要] URLエンコード処理（UTF-8）。
'* [詳  細] URLエンコードを行う。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Public Function EncodeUrl(Target As String) As String
    Dim xlBit As String
    xlBit = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    '32Bit版Excelの場合はScriptControl版使用
    If xlBit = "x86" Then
        EncodeUrl = EncodeUrlFor32(Target)
    Else
        ' 64Bit版でExcel2010以前の場合自前実装版を使用
        Dim ver As Integer: ver = CInt(Application.Version)
        If ver <= 14 Then
            EncodeUrl = EncodeUrlMyP(Target)
            'PowerShell版は使用しない
            'EncodeURL = EncodeURLByPs(target)
        ' 64Bit版でExcel2013以降の場合はExcel関数使用
        Else
            EncodeUrl = EncodeUrlForXl2013OrLater(Target)
        End If
    End If

End Function

'******************************************************************************
'* [概  要] URLエンコード処理。
'* [詳  細] URLエンコードを行う。Excel2010以前の64bit用。PowerShellを使用する。
'*
'* @param target 対象となる文字列
'* @param pEncode encode指定（shift-jis等）
'* @return エンコード後文字列
'******************************************************************************
Private Function EncodeURLByPs(Target As String, Optional pEncode As String = "utf-8") As String
    Dim cmd
    Dim i As Long, c As Long
    Dim tmp As String, flg As Boolean
    flg = False
    
    For i = 1 To Len(Target)
        tmp = Mid(Target, i, 1)
        c = Asc(tmp)
        Select Case tmp
            ' 予約文字の場合
            Case ":", "/", "?", "#", "[", "]", "@", "!", "$", "&", "'", "(", ")", "*", "+", ",", ";", "="
                flg = True
                Exit For
        End Select
        If c < 33 Or c > 127 Then
                flg = True
                Exit For
        End If
    Next
    ' TODO：文字が多いと途中で切れる
    If flg Then
        Target = Replace(Target, "`", "``")
        Target = Replace(Target, """", "`""""""")
        cmd = "PowerShell -Command ""[void]([Reflection.Assembly]::LoadWithPartialName(""""""System.Web""""""));[Web.HttpUtility]::UrlEncode(""""""" & Target & """"""", [Text.Encoding]::GetEncoding(""""""" & pEncode & """""""))"""
        With CreateObject("WScript.Shell").Exec(cmd)
            Do While (.status = 0)
                DoEvents
            Loop
            EncodeURLByPs = .stdOut.ReadLine
        End With
        Exit Function
    End If
    EncodeURLByPs = Target
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
Private Function EncodeUrlMyP(Target As String) As String
    Dim buf() As Byte, s As String, i As Long
    
    With CreateObject("ADODB.Stream")
        .Mode = 3 'adModeReadWrite
        .Open
        .Type = 2 'adTypeText
        .charset = "UTF-8"
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
    EncodeUrlMyP = s
End Function

'******************************************************************************
'* [概  要] URLエンコード処理。
'* [詳  細] URLエンコードを行う。Excel2013以降で使用可能。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Private Function EncodeUrlForXl2013OrLater(Target As String) As String
    EncodeUrlForXl2013OrLater = Application.WorksheetFunction.EncodeUrl(Target)
End Function

'******************************************************************************
'* [概  要] URLエンコード処理。
'* [詳  細] URLエンコードを行う。64bitでは使えない。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Private Function EncodeUrlFor32(Target As String) As String
    Dim sc As Object

    Set sc = CreateObject("ScriptControl")
    sc.Language = "Jscript"
    EncodeUrlFor32 = sc.CodeObject.encodeURIComponent(Target)
    Set sc = Nothing

End Function



