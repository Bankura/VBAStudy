VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Loggerクラス
'* [詳  細] ログ出力用のクラス。
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum定義
'******************************************************************************
Public Enum LogLevelEnum
    lvNone
    lvFatal
    lvWarn
    lvInfo
    lvTrace
End Enum

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mLogOutputLevel As LogLevelEnum
Private mIwConsoleOutEnabled As Boolean
Private mLogDirPath As String
Private mLogFileName As String
Private mLogFilePath As String
Private mStatusBarOutEnabled As Boolean
Private mUserName As String

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* LogOutputLevel プロパティ
'* ログ出力レベル
'*-----------------------------------------------------------------------------
Public Property Get LogOutputLevel() As LogLevelEnum
    LogOutputLevel = mLogOutputLevel
End Property
Public Property Let LogOutputLevel(arg As LogLevelEnum)
    mLogOutputLevel = arg
End Property

'*-----------------------------------------------------------------------------
'* IwConsoleOutEnabled プロパティ
'* イミディエイトウィンドウ出力可否
'*-----------------------------------------------------------------------------
Public Property Get IwConsoleOutEnabled() As Boolean
    IwConsoleOutEnabled = mIwConsoleOutEnabled
End Property
Public Property Let IwConsoleOutEnabled(arg As Boolean)
    mIwConsoleOutEnabled = arg
End Property

'*-----------------------------------------------------------------------------
'* LogDirPath プロパティ
'* ログ出力先ディレクトリ
'*-----------------------------------------------------------------------------
Public Property Get LogDirPath() As String
    LogDirPath = mLogDirPath
End Property
Public Property Let LogDirPath(arg As String)
    mLogDirPath = FileUtils.AddPathSeparator(arg)
    mLogFilePath = mLogDirPath & mLogFileName
End Property

'*-----------------------------------------------------------------------------
'* LogFileName プロパティ
'* ログファイル名
'*-----------------------------------------------------------------------------
Public Property Get LogFileName() As String
    LogFileName = mLogFileName
End Property
Public Property Let LogFileName(arg As String)
    mLogFileName = arg
    mLogFilePath = mLogDirPath & mLogFileName
End Property

'*-----------------------------------------------------------------------------
'* StatusBarOutEnabled プロパティ
'* ステータスバー表示可否
'*-----------------------------------------------------------------------------
Public Property Get StatusBarOutEnabled() As Boolean
    StatusBarOutEnabled = mStatusBarOutEnabled
End Property
Public Property Let StatusBarOutEnabled(arg As Boolean)
    mStatusBarOutEnabled = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    mLogOutputLevel = LogLevelEnum.lvNone
    mIwConsoleOutEnabled = False
    mLogDirPath = ""
    mLogFileName = ""
    mStatusBarOutEnabled = False
    mUserName = Base.GetWshNetwork.UserName
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Init メソッド
'* [詳  細] 初期化用メソッド｡
'*
'* @param myLogLevel     ログ出力レベル
'* @param iwFlg          イミディエイトウィンドウ出力可否
'* @param strLogDirPath  ログ出力先ディレクトリ
'* @param strLogFileName ログファイル名
'* @param rotateFileSize ログローテートファイルサイズ（デフォルト10MB）
'* @return Loggerオブジェクト
'*
'******************************************************************************
Public Function Init(ByVal myLogLevel As LogLevelEnum, _
                     ByVal iwFlg As Boolean, _
                     ByVal strLogDirPath As String, _
                     ByVal strLogFileName As String, _
                     Optional ByVal rotateFileSize As Long = 10485760) As Logger
    mLogOutputLevel = myLogLevel
    mIwConsoleOutEnabled = iwFlg
    LogDirPath = strLogDirPath
    LogFileName = strLogFileName
    
    ' ログローテーション
    If Dir(mLogFilePath) <> "" Then
        If FileLen(mLogFilePath) > rotateFileSize Then
            Name mLogFilePath As mLogFilePath & "." & Format(Now, "yyyymmdd-hhmm")
        End If
    End If
    
    Set Init = Me
End Function

'******************************************************************************
'* [概  要] Trace
'* [詳  細] Traceログを出力する。
'*
'* @param strLogText ログに出力するテキスト
'******************************************************************************
Public Sub Trace(strLogText As String)
    Call Log(strLogText, LogLevelEnum.lvTrace)
End Sub

'******************************************************************************
'* [概  要] Info
'* [詳  細] Infoログを出力する。
'*
'* @param strLogText ログに出力するテキスト
'******************************************************************************
Public Sub Info(strLogText As String)
    Call Log(strLogText, LogLevelEnum.lvInfo)
End Sub

'******************************************************************************
'* [概  要] Warn
'* [詳  細] Warnログを出力する。
'*
'* @param strLogText ログに出力するテキスト
'******************************************************************************
Public Sub Warn(strLogText As String)
    Call Log(strLogText, LogLevelEnum.lvWarn)
End Sub

'******************************************************************************
'* [概  要] Fatal
'* [詳  細] Fatalログを出力する。
'*
'* @param strLogText ログに出力するテキスト
'******************************************************************************
Public Sub Fatal(strLogText As String)
    Call Log(strLogText, LogLevelEnum.lvFatal)
End Sub

'******************************************************************************
'* [概  要] Log
'* [詳  細] ログを出力する。
'*
'* @param strLogText ログに出力するテキスト
'* @param myLogLevel ログレベル
'******************************************************************************
Public Sub Log(strLogText As String, myLogLevel As LogLevelEnum)
    On Error GoTo ErrorHandler

    If mLogOutputLevel = LogLevelEnum.lvNone And mIwConsoleOutEnabled = False Then
        Exit Sub
    End If

    Dim strLogLevel As String
    Select Case myLogLevel
        Case LogLevelEnum.lvFatal
            strLogLevel = "FATAL"
        Case LogLevelEnum.lvWarn
            strLogLevel = "WARN"
        Case LogLevelEnum.lvInfo
            strLogLevel = "INFO"
        Case Else
            strLogLevel = "TRACE"
    End Select
    
    Dim strOutMsg As String
    strOutMsg = GetTimestampString() & " [" & strLogLevel & "] [" & mUserName & "] " & strLogText
    If myLogLevel <= mLogOutputLevel Then
        Dim lngFileNum As Long: lngFileNum = FreeFile()
        Open mLogFilePath For Append As #lngFileNum
        Print #lngFileNum, strOutMsg
        Close #lngFileNum
    End If
    If mIwConsoleOutEnabled Then
        Debug.Print strOutMsg
    End If
    Exit Sub
ErrorHandler:
    Debug.Print GetTimestampString() & " [FATAL] エラーが発生：Number=" & Err.Number & " Description=" & Err.Description
    Debug.Print GetTimestampString() & " [CAUTION] ログ出力不可。イミディエイトウィンドウ出力に切替。"
    mIwConsoleOutEnabled = True
    mLogOutputLevel = LogLevelEnum.lvNone
    Debug.Print strOutMsg
End Sub

'******************************************************************************
'* [概  要] OutStatusBar
'* [詳  細] ステータスバーに指定したテキストを表示する。
'*
'* @param textStr ステータスバーに表示するテキスト
'******************************************************************************
Public Sub OutStatusBar(textStr As String)
    If mStatusBarOutEnabled Then
        Application.StatusBar = textStr
    End If
End Sub

'******************************************************************************
'* [概  要] GetTimestampString
'* [詳  細] 現在時刻の年月日時分秒ミリ秒を「YYYY/MM/DD HH:NN:SS.000」形式の
'*          文字列で取得する。
'*
'* @return String 現在時刻文字列
'******************************************************************************
Private Function GetTimestampString() As String
    Dim dblTimer As Double
    Dim hour As Integer
    Dim minute As Integer
    Dim second As Integer
    Dim mSec As Double
    
    dblTimer = CDbl(Timer)
    hour = dblTimer \ 3600
    minute = (dblTimer Mod 3600) \ 60
    second = dblTimer Mod 60
    mSec = Fix((dblTimer - Fix(dblTimer)) * 1000)

    GetTimestampString = Format(Now, "yyyy/mm/dd") & " " & Format(hour, "00") & ":" & _
                         Format(minute, "00") & ":" & Format(second, "00") & "." & Format(mSec, "000")
End Function
