VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingExEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Scripting.Encoderクラスのラップ・拡張クラス
'* [詳  細] Scripting.EncoderクラスのWrapperとしての機能を提供する他、
'*          Scripting.Encoderクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Scripting.Encoderクラスについて。
' VBScriptを暗号化する方法を提供するオブジェクト。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc392170(v%3dmsdn.10)>
' <https://gallery.technet.microsoft.com/scriptcenter/16439c02-3296-4ec8-9134-6eb6fb599880>
' <http://sammaya.jugem.jp/?eid=13>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As Scripting.Encoder
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

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("Scripting.Encoder")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] EncodeScriptFile メソッド(既定メンバー)
'* [詳  細] EncodeScriptFile のラッパーメソッド。
'*          szExt によって決定されたエンコーダを呼び出し、bstrStreamIn
'*          およびオプション引数を渡す。
'*
'* @param szExt エンコード対象拡張子（e.g. ".vbs", ".js", ".htm", ".html"）
'* @param bstrStreamIn ソースScriptテキスト
'* @param cFlags フラグ（0）
'* @param bstrDefaultLang デフォルトのLanguage
'*                        （e.g. "VBScript", "JScript". Can be a null string, ""）
'* @return EncodeされたScript文字列
'******************************************************************************
Function EncodeScriptFile(szExt As String, bstrStreamIn As String, cFlags As Long, bstrDefaultLang As String) As String
Attribute EncodeScriptFile.VB_UserMemId = 0
    EncodeScriptFile = mOrigin.EncodeScriptFile(szExt, bstrStreamIn, cFlags, bstrDefaultLang)
End Function
