VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Commander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CUIコマンド実行クラス
'* [詳  細] DOSコマンド・Powershell等のコマンド実行に関する処理を提供する。
'*
'* [参  考]　https://thom.hateblo.jp/entry/2017/01/31/012913
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mBaseCommand As String

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* BaseCommand プロパティ
'*
'* 実行コマンド文字列のフォーマットを格納する。
'* ex) DOSの場合、"cmd.exe /c {0} > {1}"
'*     （{0}=実行するコマンド,{1}=一時ファイル）
'*-----------------------------------------------------------------------------
Public Property Get BaseCommand() As String
    BaseCommand = mBaseCommand
End Property
Public Property Let BaseCommand(arg As String)
    mBaseCommand = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] GetCommandResultAsTextStream
'* [詳  細] コマンドを実行し、実行した結果を一時ファイルに保存する。
'*          一時ファイルから実行結果をTextStreamで取得する。
'*
'* @param commandString コマンド文字列
'* @param myTempPath 一時ファイルパス（標準のパス以外にしたい場合に指定）
'* @return Object(Of Scripting.TextStream)
'******************************************************************************
Public Function GetCommandResultAsTextStream(commandString, Optional myTempPath) As Object

    ' 一時ファイルの取得
    Dim tempPath As String
    If IsMissing(myTempPath) Then
        tempPath = FileUtils.GetTempFilePath
    Else
        tempPath = myTempPath
    End If
    
    ' 実行コマンド構築
    Dim execComandString As New StringEx
    Call execComandString.Append(mBaseCommand).PlaceHolder(commandString, tempPath)
    
    ' コマンド実行
    Call Core.Wsh.Run(execComandString.ToString, WshHide, True)
    Set GetCommandResultAsTextStream = IO.Fso.OpenTextFile(tempPath, Base.ForReading)
End Function

'******************************************************************************
'* [概  要] GetCommandResult
'* [詳  細] コマンドを実行し、実行した結果文字列を取得する。
'*
'* @param commandString コマンド文字列
'* @return String コマンド実行結果
'******************************************************************************
Public Function GetCommandResult(commandString) As String
    Dim ret As String

    ' Scripting.TextStream
    Dim ts As Object

    Dim tempPath As String: tempPath = FileUtils.GetTempFilePath
    Set ts = GetCommandResultAsTextStream(commandString, tempPath)
    If ts.AtEndOfStream Then
        ret = ""
    Else
        ret = ts.ReadAll
    End If
    ts.Close
    Call IO.Fso.DeleteFile(tempPath, True)
    GetCommandResult = ret
End Function

'******************************************************************************
'* [概  要] GetCommandResultAsArray
'* [詳  細] コマンドを実行し、実行した結果を行毎に分割した文字配列で取得する。
'*
'* @param commandString コマンド文字列
'* @return String() コマンド実行結果
'******************************************************************************
Public Function GetCommandResultAsArray(commandString) As String()
    Dim ret() As String
    ret = Split(GetCommandResult(commandString), vbNewLine)
    GetCommandResultAsArray = ret
End Function

'******************************************************************************
'* [概  要] Exec
'* [詳  細] コマンドを実行し、結果をCommandResultオブジェクトに格納して返却する。
'*
'* @param commandString コマンド文字列
'* @return CommandResult情報
'*
'******************************************************************************
Public Function Exec(commandString As String) As CommandResult
    ' コマンド実行
    Dim oExec: Set oExec = Core.Wsh.Exec(commandString)
    
    ' コマンド実行中待機
    Do While (oExec.Status = 0)
        Application.Wait [Now()] + 100 / 86400000
        DoEvents
    Loop
    
    ' 実行結果設定
    Dim res As CommandResult: Set res = New CommandResult
    res.ExitCd = oExec.ExitCode
    If oExec.ExitCode <> 0 Then
        ' コマンド失敗時処理
        Dim errTxt As String: errTxt = oExec.StdErr.ReadAll
        res.StdErrTxt = errTxt
        Set Exec = res
        Exit Function
    End If

    ' 正常時処理
    Dim varr As New ArrayEx
    While Not oExec.StdOut.AtEndOfStream
        varr.Add oExec.StdOut.ReadLine
    Wend
    res.StdOutList = varr.ToArray
    res.RowMax = varr.Length

    Set Exec = res

End Function


'******************************************************************************
'* [概  要] RunExplorer
'* [詳  細] エクスプローラを起動する。
'*
'******************************************************************************
Public Sub RunExplorer()
    Core.Wsh.Run "explorer.exe"
End Sub

'******************************************************************************
'* [概  要] RunCalc
'* [詳  細] 電卓を起動する。
'*
'******************************************************************************
Public Sub RunCalc()
    Core.Wsh.Run "calc.exe"
    Core.Wsh.AppActivate "電卓"
End Sub

'******************************************************************************
'* [概  要] RunTaskManager
'* [詳  細] タスク マネージャーを起動する。
'*
'******************************************************************************
Public Sub RunTaskManager()
    Core.Wsh.Run "taskmgr.exe /7"
    Core.Wsh.AppActivate "タスク マネージャー"
End Sub

'******************************************************************************
'* [概  要] RunSystemProperties
'* [詳  細] システムのプロパティを起動する。
'*
'******************************************************************************
Public Sub RunSystemProperties()
    Core.Wsh.Run "sysdm.cpl"
    Core.Wsh.AppActivate "システムのプロパティ"
End Sub

'******************************************************************************
'* [概  要] RunEventViewer
'* [詳  細] イベント ビューアーを起動する。
'*
'******************************************************************************
Public Sub RunEventViewer()
    Core.Wsh.Run "eventvwr.msc /s"
    Core.Wsh.AppActivate "イベント ビューアー"
End Sub

'******************************************************************************
'* [概  要] RunComputerManagement
'* [詳  細] コンピューターの管理を起動する。
'*
'******************************************************************************
Public Sub RunComputerManagement()
    Core.Wsh.Run "compmgmt.mSc"
    Core.Wsh.AppActivate "コンピューターの管理"
End Sub

'******************************************************************************
'* [概  要] RunMsInfo
'* [詳  細] システム情報を起動する。
'*
'******************************************************************************
Public Sub RunMsInfo()
    Core.Wsh.Run "msinfo32.exe"
    Core.Wsh.AppActivate "システム情報"
End Sub

'******************************************************************************
'* [概  要] RunInternetOption
'* [詳  細] インターネット オプションを起動する。
'*
'******************************************************************************
Public Sub RunInternetOption()
    Core.Wsh.Run "inetcpl.cpl"
    Core.Wsh.AppActivate "インターネットのプロパティ"
End Sub

'******************************************************************************
'* [概  要] RunPerformanceMonitor
'* [詳  細] パフォーマンス モニターを起動する。
'*
'******************************************************************************
Public Sub RunPerformanceMonitor()
    Core.Wsh.Run "perfmon.exe"
    Core.Wsh.AppActivate "パフォーマンス モニター"
End Sub

'******************************************************************************
'* [概  要] RunResourceMonitor
'* [詳  細] リソース モニターを起動する。
'*
'******************************************************************************
Public Sub RunResourceMonitor()
    Core.Wsh.Run "resmon.exe"
    Core.Wsh.AppActivate "リソース モニター"
End Sub

'******************************************************************************
'* [概  要] RunRegistoryEditor
'* [詳  細] レジストリ エディターを起動する。
'*
'******************************************************************************
Public Sub RunRegistoryEditor()
    Core.Wsh.Run "regedt32.exe"
    Core.Wsh.AppActivate "レジストリ エディター"
End Sub

'******************************************************************************
'* [概  要] RunTaskScheduler
'* [詳  細] タスク スケジューラを起動する。
'*
'******************************************************************************
Public Sub RunTaskScheduler()
    Core.Wsh.Run "taskschd.msc /s"
    Core.Wsh.AppActivate "タスク スケジューラ"
End Sub

'******************************************************************************
'* [概  要] RunServices
'* [詳  細] サービスを起動する。
'*
'******************************************************************************
Public Sub RunServices()
    Core.Wsh.Run "services.msc"
    Core.Wsh.AppActivate "サービス"
End Sub

'******************************************************************************
'* [概  要] RunCharmap
'* [詳  細] 文字コード表を起動する。
'*
'******************************************************************************
Public Sub RunCharmap()
    Core.Wsh.Run "charmap.exe"
    Core.Wsh.AppActivate "文字コード表"
End Sub

'******************************************************************************
'* [概  要] RunMsPaint
'* [詳  細] ペイントを起動する。
'*
'******************************************************************************
Public Sub RunMsPaint()
    Core.Wsh.Run "mspaint.exe"
    Core.Wsh.AppActivate "ペイント"
End Sub

'******************************************************************************
'* [概  要] RunNotepad
'* [詳  細] メモ帳を起動する。
'*
'******************************************************************************
Public Sub RunNotepad()
    Core.Wsh.Run "notepad.exe"
    Core.Wsh.AppActivate "メモ帳"
End Sub

'******************************************************************************
'* [概  要] RunCmd
'* [詳  細] コマンドプロンプトを起動する。
'*
'******************************************************************************
Public Sub RunCmd()
    Core.Wsh.Run "cmd.exe"
    Core.Wsh.AppActivate "コマンドプロンプト"
End Sub

'******************************************************************************
'* [概  要] RunPowerShell
'* [詳  細] PowerShellを起動する。
'*
'******************************************************************************
Public Sub RunPowerShell()
    Core.Wsh.Run "powershell.exe"
    Core.Wsh.AppActivate "Windows PowerShell"
End Sub

'******************************************************************************
'* [概  要] RunVSCode
'* [詳  細] Visual Studio Codeを起動する。
'*
'******************************************************************************
Public Sub RunVSCode()
    Core.Wsh.Run "code"
End Sub

