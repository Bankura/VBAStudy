VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WScriptExWshURLShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WScript.WshURLShortcutクラスのラップ・拡張クラス
'* [詳  細] WScript.WshURLShortcutクラスのWrapperとしての機能を提供する他、
'*          WScript.WshURLShortcutクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WScript.WshShortcutクラスについて。
' URL ショートカットへのオブジェクト参照を作成します。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364464(v=msdn.10)>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As IWshRuntimeLibrary.WshURLShortcut
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
'* FullName プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ショートカット オブジェクトのリンク先への絶対パスを返します。
'*-----------------------------------------------------------------------------
Property Get FullName() As String
    FullName = mOrigin.FullName
End Property

'*-----------------------------------------------------------------------------
'* TargetPath プロパティ
'*
'* [補  足]
'* ・ショートカットの実行可能ファイルへのパスです。
'*-----------------------------------------------------------------------------
Property Get TargetPath() As String
    TargetPath = mOrigin.TargetPath
End Property
Property Let TargetPath(arg As String)
    mOrigin.TargetPath = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WScript.WshURLShortcut")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Save メソッド
'* [詳  細] Save のラッパーメソッド。
'*          ショートカット オブジェクトを保存します。
'*
'******************************************************************************
Sub Save()
    Call mOrigin.Save
End Sub

'******************************************************************************
'* [概  要] Load メソッド
'* [詳  細] Load のラッパーメソッド。
'*          隠しメソッド。
'*
'* @param PathLink
'******************************************************************************
Sub Load(PathLink As String)
    Call mOrigin.Load(PathLink)
End Sub
