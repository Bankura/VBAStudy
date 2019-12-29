Attribute VB_Name = "LogUtil"
Option Explicit

'==============================================================================
'ログ出力用ユーティリティ関数モジュール
'==============================================================================


'******************************************************************************
'ログ出力機能共通定数
'******************************************************************************
Public Const CMN_LOG_LEVEL_ERROR = "ERROR"
Public Const CMN_LOG_LEVEL_INFO = "INFO"
Public Const CMN_LOG_LEVEL_DEBUG = "DEBUG"

'******************************************************************************
'ログ出力機能共通変数
'******************************************************************************
' ログ出力フラグ
'  0 … 出力しない
'  1 … 「DEBUG」以上出力
'  2 … 「INFO」以上を出力
'  3 … 「ERROR」のみ出力
Public cmn_outLogFlg As Integer

' イミディエイトウィンドウ出力可否フラグ
' True … 出力する
' False … 出力しない
Public cmn_iwFlg As Boolean

' ログ出力先ディレクトリ
Public cmn_strLogDirPath As String

' ログファイル名
Public cmn_strLogFileName As String

' ステータスバー表示可否フラグ
' True … 出力する
' False … 出力しない
Public cmn_sbFlg As Boolean

'******************************************************************************
' [関数名] ログ出力初期設定
' [説　明] ログを出力するための初期設定を行う関数。
' [引　数] outLogFlg      ログ出力フラグ
'          iwFlg          イミディエイトウィンドウ出力
'                         可否フラグ
'          strLogDirPath  ログ出力先ディレクトリ
'          strLogFileName ログファイル名
'          lrotateFileSize ログローテートファイルサイズ
'******************************************************************************
Public Sub InitLogSetting(outLogFlg As Integer, iwFlg As Boolean, _
                          strLogDirPath As String, strLogFileName As String, _
                          Optional divPersonFlg As Boolean = False, _
                          Optional lrotateFileSize As Long = 10485760)
    cmn_outLogFlg = outLogFlg
    cmn_iwFlg = iwFlg
    cmn_strLogDirPath = AddPathSeparator(strLogDirPath)
    
    If divPersonFlg Then
        cmn_strLogFileName = Mid(strLogFileName, 1, InStrRev(strLogFileName, ".") - 1) & "_" & _
                             CStr(CreateObject("WScript.Network").UserName) & _
                             Mid(strLogFileName, InStrRev(strLogFileName, "."), _
                                 Len(strLogFileName) - InStrRev(strLogFileName, ".") + 1)
    Else
        cmn_strLogFileName = strLogFileName
    End If
    
    ' ログサイズが指定したサイズを超えると切り替えを行う（デフォルト10MB）
    If Dir(cmn_strLogDirPath & cmn_strLogFileName) <> "" Then
        If FileLen(cmn_strLogDirPath & cmn_strLogFileName) > lrotateFileSize Then
            Name cmn_strLogDirPath & cmn_strLogFileName _
                As cmn_strLogDirPath & cmn_strLogFileName & "." & Format(Now, "yyyymmdd-hhmm")
        End If
    End If
End Sub

'******************************************************************************
' [関数名] ログ出力関数
' [説　明] ログを出力する関数。
' [引　数] strLogText ログに出力する内容
'******************************************************************************
Public Sub OutLog(strLogText As String, strLogLevel As String)
    On Error GoTo ErrorHandler
    Dim lngLogLevel As Long
    Dim lngFileNum As Long
    Dim strLogFile As String
    Dim strOutMsg As String
    
    If cmn_outLogFlg = 0 And cmn_iwFlg = False Then
        Exit Sub
    End If

    Select Case strLogLevel
        Case CMN_LOG_LEVEL_ERROR
            lngLogLevel = 3
        Case CMN_LOG_LEVEL_INFO
            lngLogLevel = 2
        Case CMN_LOG_LEVEL_DEBUG
            lngLogLevel = 1
        Case Else
            lngLogLevel = 0
    End Select
    
    strOutMsg = GetNowWithMSec() & " [" & strLogLevel & "] " & strLogText
    If lngLogLevel >= cmn_outLogFlg Then
        strLogFile = cmn_strLogDirPath & cmn_strLogFileName
        lngFileNum = FreeFile()
        Open strLogFile For Append As #lngFileNum
        Print #lngFileNum, strOutMsg
        Close #lngFileNum
    End If
    If cmn_iwFlg Then
        Debug.Print strOutMsg
    End If
    Exit Sub
ErrorHandler:
    Debug.Print GetNowWithMSec() & " [FATAL] 【OutLog】エラーが発生：Number=" & Err.Number & " Description=" & Err.Description
    Debug.Print GetNowWithMSec() & " 【CAUTION】ログ出力不可。イミディエイトウィンドウ出力に切り替えます。"
    cmn_iwFlg = True
    cmn_outLogFlg = 9
    Debug.Print strOutMsg
End Sub

'******************************************************************************
' [関数名] OutStatusBar ステータスバー表示関数
' [説　明] ステータスバーを表示する関数。
' [引　数] varText ステータスバーに表示する内容
'******************************************************************************
Public Sub OutStatusBar(varText As Variant)
    If cmn_sbFlg Then
        Application.StatusBar = varText
    End If
End Sub


'******************************************************************************
' [関数名] GetNowWithMSec
' [説　明] 現在時刻の年月日時分秒ミリ秒を「YYYY/MM/DD HH:NN:SS.000」形式の
'          文字列として取得する。
' [引　数] なし
' [戻り値] String 「yyyy/mm/dd hh:nn:ss.000」形式現在時刻文字列
'******************************************************************************
Private Function GetNowWithMSec() As String
    Dim dblTimer As Double
    Dim hour As Integer
    Dim minute As Integer
    Dim second As Integer
    Dim mSec As Double
    Dim rtn As String
    
    dblTimer = CDbl(Timer)
    hour = dblTimer \ 3600
    minute = (dblTimer Mod 3600) \ 60
    second = dblTimer Mod 60
    mSec = Fix((dblTimer - Fix(dblTimer)) * 1000)
    
    rtn = Format(Now, "yyyy/mm/dd") & " " & Format(hour, "00") & ":" & _
          Format(minute, "00") & ":" & Format(second, "00") & "." & Format(mSec, "000")
          
    GetNowWithMSec = rtn
End Function




