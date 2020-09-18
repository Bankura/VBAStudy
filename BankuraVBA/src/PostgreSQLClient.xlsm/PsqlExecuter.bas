Attribute VB_Name = "PsqlExecuter"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] psql実行メイン機能
'* [詳  細] Httpリクエスト送信メイン処理を実装する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* [概  要] psqlコマンドを介してSQLを実行する。
'* [詳  細] psqlコマンドを介して、SQLを実行し、実行結果を返却する。
'*          ローカル環境にpsqlがインストールされている必要がある。
'*          psqlはpostgreSQLインストール時に合わせてインストールされる。
'*
'* @param psql     psqlコマンドのフルパス
'* @param host     アクセスするDBのホスト（例：localhost）
'* @param port     アクセスするDBのポート番号
'* @param dbName   DB名
'* @param userName DBユーザ名
'* @param password DBパスワード
'* @param sql      SQL文
'* @param clEncode Clientエンコード　※任意
'* @return Result情報
'*
'******************************************************************************
Public Function ExecPsql(psql As String, host As String, port As String, _
                         dbName As String, userName As String, password As String, _
                         sql As String, Optional clEncode As String = "SJIS") As Result
    Dim cmd As String, oExec
    
    ' コマンド組立て
    cmd = psql & " -h " & host & " -p " & port & " -d " & dbName & " -U " & userName & " -c """ & sql & """ -A"
    Debug.Print cmd
    
    ' Wscript.Shellオブジェクト生成
    With CreateObject("Wscript.Shell")
        ' パスワードを環境変数に設定
        .Environment("Process").Item("PGPASSWORD") = password
    
        ' Client(psql)のエンコード設定：デフォルト「SJIS」
        .Environment("Process").Item("PGCLIENTENCODING") = clEncode
    
        ' コマンド実行
        Set oExec = .Exec(cmd)
    End With
    

    ' 実行結果設定
    Dim res As Result: Set res = New Result
    res.ExitCd = oExec.ExitCode
    If oExec.ExitCode <> 0 Then
        ' コマンド失敗時処理
        Dim errTxt As String: errTxt = oExec.StdErr.ReadAll
        res.StdErrTxt = errTxt
        Set ExecPsql = res
        Exit Function
    End If

    ' 正常時処理
    ' データ量計測
    Dim lRowMax As Long: lRowMax = 0
    Dim lColMax As Long: lColMax = 0
    
    Dim stdOutTxt As String
    While Not oExec.stdOut.AtEndOfStream
        Dim strTmp As String: strTmp = oExec.stdOut.ReadLine
        If strTmp <> "" Then
            If lRowMax = 0 Then
                Dim vTmpCols As Variant: vTmpCols = Split(strTmp, "|")
                lColMax = UBound(vTmpCols) - LBound(vTmpCols) + 1
                stdOutTxt = strTmp
            Else
                stdOutTxt = stdOutTxt & vbCrLf & strTmp
            End If
            lRowMax = lRowMax + 1
        End If
    Wend
    'データ設定
    Dim vArray()
    ReDim vArray(0 To lRowMax - 1, 0 To lColMax - 1)
    Dim i As Long, j As Long, stdOut As Variant, cols As Variant
    stdOut = Split(stdOutTxt, vbCrLf)
    For i = LBound(vArray, 1) To UBound(vArray, 1)
        Debug.Print "Line: " & stdOut(i)
        If stdOut(i) <> "" Then
            If InStr(stdOut(i), "|") = 0 Then
                vArray(i, 0) = stdOut(i)
                For j = LBound(vArray, 2) + 1 To UBound(vArray, 2)
                    vArray(i, j) = ""
                Next j
            Else
                cols = Split(stdOut(i), "|")
                For j = LBound(cols) To UBound(cols)
                    vArray(i, j) = cols(j)
                Next j
            End If

        End If
    Next i
    res.StdOutList = vArray
    res.RowMax = lRowMax
    res.ColMax = lColMax
    
    Set oExec = Nothing
    Set ExecPsql = res

End Function

'******************************************************************************
'* [概  要] 開始処理
'* [詳  細] 画面描画の停止等、処理性能に影響のある設定を変更する。
'*
'******************************************************************************
Public Sub CommonStr()
    '各種設定変更
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

'******************************************************************************
'* [概  要] 終了処理
'* [詳  細] 開始処理で行った設定を解除する。
'*
'******************************************************************************
Public Sub CommonEnd()
    '各種設定変更
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

'******************************************************************************
'* [概  要] シート存在チェック
'* [詳  細] シートが存在するかを判定する。
'*
'* @param strName ブック名
'* @param wb ブックオブジェクト
'* @return 処理結果（True:存在する False：存在しない）
'******************************************************************************
Public Function CheckSheet(strName As String, Optional wb As Workbook) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    'シートの検索
    For Each ws In wb.Worksheets
        If ws.Name = strName Then
            flg = True
            Exit For
        End If
    Next ws
    
    CheckSheet = flg
End Function
