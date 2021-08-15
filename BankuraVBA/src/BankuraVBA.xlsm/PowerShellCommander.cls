VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerShellCommander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] PowerShellコマンド実行クラス
'* [詳  細] PowerShellコマンド（スクリプトレット）の実行に関する処理を提供する。
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
    mCommander.BaseCommand = "powershell -ExecutionPolicy RemoteSigned -Command $cmd='{0} | Out-File -filePath {1} -encoding Default';Invoke-Expression $cmd"
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] WriteScriptAndRun
'* [詳  細] スクリプトを作成し、実行した結果を取得する。
'*
'* @param scriptString Script文字列、または1次元配列
'* @param recieveResultString 結果を文字列で受け取るか
'*                            （True:String, False:配列）
'* @return Variant コマンド実行結果
'******************************************************************************
Public Function WriteScriptAndRun(scriptString, Optional recieveResultString As Boolean = True) As Variant
    Dim ret As String

    Dim tempScrPath As String: tempScrPath = FileUtils.GetTempFilePath(, ".ps1")
    Call FileUtils.WriteTextFile(tempScrPath, scriptString, "Shift_JIS")
    
    WriteScriptAndRun = Exec(tempScrPath, recieveResultString)

    Call IO.fso.DeleteFile(tempScrPath, True)
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
        Exec = mCommander.GetCommandResult(StringUtils.EscapeSingleQuote(execString))
    Else
        Exec = mCommander.GetCommandResultAsArray(StringUtils.EscapeSingleQuote(execString))
    End If
End Function

'******************************************************************************
'* [概  要] PSVersionTable
'* [詳  細] PSVersionTableを取得する。
'*          PSVersionTableは PowerShell のバージョン情報を含むシェル変数。
'*
'* @return DictionaryEx PSVersionTable
'******************************************************************************
Public Function PSVersionTable() As DictionaryEx
    Dim dic As DictionaryEx: Set dic = New DictionaryEx
    Set PSVersionTable = dic.Init(ParseJson(Exec("ConvertTo-Json $PSVersionTable")))
End Function

'******************************************************************************
'* [概  要] GetExecutionPolicy
'* [詳  細] 実行ポリシーを確認する。
'*
'* @return String 実行ポリシー
'******************************************************************************
Public Function GetExecutionPolicy() As String
    GetExecutionPolicy = Exec("Get-ExecutionPolicy")
End Function

