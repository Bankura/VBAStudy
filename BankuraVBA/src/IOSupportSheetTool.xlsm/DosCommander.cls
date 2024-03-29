VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DosCommander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] DOSコマンド実行クラス
'* [詳  細] DOSコマンドのコマンド実行に関する処理を提供する。
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
Private mCommander As Commander

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mCommander = New Commander
    mCommander.BaseCommand = "cmd.exe /c {0} > {1}"
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
    Set GetCommandResultAsTextStream = mCommander.GetCommandResultAsTextStream(commandString, myTempPath)
End Function

'******************************************************************************
'* [概  要] GetCommandResult
'* [詳  細] コマンドを実行し、実行した結果文字列を取得する。
'*
'* @param commandString コマンド文字列
'* @return String コマンド実行結果
'******************************************************************************
Public Function GetCommandResult(commandString) As String
    GetCommandResult = mCommander.GetCommandResult(commandString)
End Function

'******************************************************************************
'* [概  要] GetCommandResultAsArray
'* [詳  細]コマンドを実行し、実行した結果を行毎に分割した文字配列で取得する。
'*
'* @param commandString コマンド文字列
'* @return String() コマンド実行結果
'******************************************************************************
Public Function GetCommandResultAsArray(commandString) As String()
    GetCommandResultAsArray = mCommander.GetCommandResultAsArray(commandString)
End Function

'******************************************************************************
'* [概  要] WriteBatchAndRun
'* [詳  細] パッチファイルを作成し、実行した結果を取得する。
'*
'* @param batchString パッチファイル文字列、または1次元配列
'* @param recieveResultString 結果を文字列で受け取るか
'*                            （True:String, False:配列）
'* @return Variant コマンド実行結果
'******************************************************************************
Public Function WriteBatchAndRun(batchString, Optional recieveResultString As Boolean = True) As Variant
    Dim tempPath As String: tempPath = FileUtils.GetTempFilePath(, ".bat")
    Call FileUtils.WriteTextFile(tempPath, batchString, "Shift_JIS")
    
    WriteBatchAndRun = Exec(tempPath, recieveResultString)

    Call IO.fso.DeleteFile(tempPath, True)
End Function

'******************************************************************************
'* [概  要] Exec
'* [詳  細] 指定したコマンドまたはスクリプトを実行した結果を取得する。
'*
'* @param execString 実行文字列（コマンドまたはスクリプト）
'* @param recieveResultString 結果を文字列で受け取るか
'*                            （True:String, False:配列）
'* @return Variant スクリプト実行結果
'******************************************************************************
Public Function Exec(execString As String, Optional recieveResultString As Boolean = True) As Variant
    If recieveResultString Then
        Exec = mCommander.GetCommandResult(execString)
    Else
        Exec = mCommander.GetCommandResultAsArray(execString)
    End If
End Function

Public Function CmdDir(ByVal Path As String, ParamArray options() As Variant) As String()
    CmdDir = GetCommandResultAsArray("DIR " & Path & " " & Join(options, " "))
End Function

Public Function GetFilePathsRecursive(ByVal Path As String) As String()
    GetFilePathsRecursive = CmdDir(Path, "/b", "/s", "/a:-d")
End Function
