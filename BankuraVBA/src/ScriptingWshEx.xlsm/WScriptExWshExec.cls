VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WScriptExWshExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WScript.WshExecクラスのラップ・拡張クラス
'* [詳  細] WScript.WshExecクラスのWrapperとしての機能を提供する他、
'*          WScript.WshExecクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WScript.WshExecクラスについて。
' Exec を使って実行したスクリプトのステータス情報を提供します。
' また、StdIn、StdOut、および StdErr の各ストリームへのアクセスも提供します。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364375(v=msdn.10)>。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As IWshRuntimeLibrary.wshExec
Private mOrigin As Object
    
'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* ExitCode プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ExitCode。
'*-----------------------------------------------------------------------------
Property Get ExitCode() As Long
    ExitCode = mOrigin.ExitCode
End Property

'*-----------------------------------------------------------------------------
'* ProcessID プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ProcessID。
'*-----------------------------------------------------------------------------
Property Get ProcessID() As Long
    ProcessID = mOrigin.ProcessID
End Property

'*-----------------------------------------------------------------------------
'* Status プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Execメソッドを使って実行したスクリプトのステータス情報を返します。
'*-----------------------------------------------------------------------------
Property Get Status() As WshExecStatus
    Status = mOrigin.Status
End Property

'*-----------------------------------------------------------------------------
'* StdErr プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Exec オブジェクトの stderr 出力ストリームへのアクセスを提供します。
'*-----------------------------------------------------------------------------
Property Get StdErr() As ScriptingExTextStream
    Dim wrap As ScriptingExTextStream: Set wrap = New ScriptingExTextStream
    Set wrap.Origin = mOrigin.StdErr
    Set StdErr = wrap
End Property

'*-----------------------------------------------------------------------------
'* StdIn プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Exec オブジェクトの stdin 入力ストリームを提供します。
'* ・Exec で起動されたプロセスにデータを引き渡す場合に StdIn プロパティを使用します。
'*-----------------------------------------------------------------------------
Property Get StdIn() As ScriptingExTextStream
    Dim wrap As ScriptingExTextStream: Set wrap = New ScriptingExTextStream
    Set wrap.Origin = mOrigin.StdIn
    Set StdIn = wrap
End Property

'*-----------------------------------------------------------------------------
'* StdOut プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Exec オブジェクトの書き込み専用の stdout 出力ストリームを提供します。
'* ・スクリプトから標準出力に送信された全情報のコピー (読み取り専用) が格納されます。
'*-----------------------------------------------------------------------------
Property Get StdOut() As ScriptingExTextStream
    Dim wrap As ScriptingExTextStream: Set wrap = New ScriptingExTextStream
    Set wrap.Origin = mOrigin.StdOut
    Set StdOut = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WScript.WshExec")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Terminateメソッド
'* [詳  細] Terminateのラッパーメソッド。
'*          Exec メソッドによって開始されたプロセスを終了するようにスクリプト
'*          エンジンに指示します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364387%28v%3dmsdn.10%29>
'*
'******************************************************************************
Sub Terminate()
    Call mOrigin.Terminate
End Sub


