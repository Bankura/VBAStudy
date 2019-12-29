Attribute VB_Name = "Main"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Httpリクエスト送信メイン機能
'* [詳  細] Httpリクエスト送信メイン処理を実装する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

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
Public Function SendRequest(method As String, url As String, reqHdrParams() As String, reqBody As String, _
                            Optional lTimeout As Long = 0, _
                            Optional resEncode As String = "utf-8") As Response
    Dim winHttp As Variant
    Set winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' HTTP リクエスト
    winHttp.Open method, url, False
    
    ' リクエストヘッダ設定
    Dim i As Long
    If GetDimension(reqHdrParams) <> 0 Then
        For i = 0 To UBound(reqHdrParams, 1)
            winHttp.setRequestHeader reqHdrParams(i, 0), reqHdrParams(i, 1)
        Next i
    End If

    ' タイムアウト設定
    If lTimeout > 0 Then
        winHttp.SetTimeouts lTimeout, lTimeout, lTimeout, lTimeout
    End If
    
    ' リクエスト送信
    If method = "GET" Or method = "DELETE" Or method = "HEAD" Or method = "OPTIONS" Then
        winHttp.send
    Else
        winHttp.send reqBody
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
        .Charset = resEncode
        resBody = .ReadText(adReadAll)
        .Close
    End With
  
    ' レスポンス設定
    Dim res As Response: Set res = New Response
    res.StatusCd = status
    res.StatusText = statusTxt
    res.Body = resBody
    res.Headers = resHeaders
    
    Set SendRequest = res

End Function

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
'* [概  要] URLエンコード処理（UTF-8）。
'* [詳  細] URLエンコードを行う。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Public Function EncodeURL(target As String) As String
    Dim xlBit As String
    xlBit = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    '32Bit版Excelの場合はScriptControl版使用
    If xlBit = "x86" Then
        EncodeURL = EncodeUrlFor32(target)
    Else
        ' 64Bit版でExcel2010以前の場合自前実装版を使用
        Dim ver As Integer: ver = CInt(Application.Version)
        If ver <= 14 Then
            EncodeURL = EncodeUrlMyP(target)
            'PowerShell版は使用しない
            'EncodeURL = EncodeURLByPs(target)
        ' 64Bit版でExcel2013以降の場合はExcel関数使用
        Else
            EncodeURL = EncodeUrlForXl2013OrLater(target)
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
Private Function EncodeURLByPs(target As String, Optional pEncode As String = "utf-8") As String
    Dim cmd
    Dim i As Long, c As Long
    Dim tmp As String, flg As Boolean
    flg = False
    
    For i = 1 To Len(target)
        tmp = Mid(target, i, 1)
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
        target = Replace(target, "`", "``")
        target = Replace(target, """", "`""""""")
        cmd = "PowerShell -Command ""[void]([Reflection.Assembly]::LoadWithPartialName(""""""System.Web""""""));[Web.HttpUtility]::UrlEncode(""""""" & target & """"""", [Text.Encoding]::GetEncoding(""""""" & pEncode & """""""))"""
        With CreateObject("WScript.Shell").Exec(cmd)
            Do While (.status = 0)
                DoEvents
            Loop
            EncodeURLByPs = .StdOut.ReadLine
        End With
        Exit Function
    End If
    EncodeURLByPs = target
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
Private Function EncodeUrlMyP(target As String) As String
    Dim buf() As Byte, s As String, i As Long
    
    With CreateObject("ADODB.Stream")
        .Mode = 3 'adModeReadWrite
        .Open
        .Type = 2 'adTypeText
        .Charset = "UTF-8"
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
    EncodeUrlMyP = s
End Function

'******************************************************************************
'* [概  要] URLエンコード処理。
'* [詳  細] URLエンコードを行う。Excel2013以降で使用可能。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Private Function EncodeUrlForXl2013OrLater(target As String) As String
    EncodeUrlForXl2013OrLater = Application.WorksheetFunction.EncodeURL(target)
End Function

'******************************************************************************
'* [概  要] URLエンコード処理。
'* [詳  細] URLエンコードを行う。64bitでは使えない。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Private Function EncodeUrlFor32(target As String) As String
    Dim sc As Object

    Set sc = CreateObject("ScriptControl")
    sc.Language = "Jscript"
    EncodeUrlFor32 = sc.CodeObject.encodeURIComponent(target)
    Set sc = Nothing

End Function

