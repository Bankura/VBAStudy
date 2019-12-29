VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExpExMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] VBScript.Matchクラスのラップ・拡張クラス
'* [詳  細] VBScript.MatchクラスのWrapperとしての機能を提供する他、
'*          VBScript.Matchクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' VBScript.Matchクラスについて。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc392427(v=msdn.10)>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As VBScript_RegExp_55.Match
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
'* FirstIndex プロパティ（読み取り専用）
'*
'* [補  足]
'* ・。
'*-----------------------------------------------------------------------------
Property Get FirstIndex() As Long
    FirstIndex = mOrigin.FirstIndex
End Property

'*-----------------------------------------------------------------------------
'* Length プロパティ（読み取り専用）
'*
'* [補  足]
'* ・。
'*-----------------------------------------------------------------------------
Property Get Length() As Long
    Length = mOrigin.Length
End Property

'*-----------------------------------------------------------------------------
'* SubMatches プロパティ（読み取り専用）
'*
'* [補  足]
'* ・。
'*-----------------------------------------------------------------------------
Property Get SubMatches() As RegExpExSubMatches
    Dim wrap As RegExpExSubMatches: Set wrap = New RegExpExSubMatches
    Set wrap.Origin = mOrigin.SubMatches()
    Set SubMatches = wrap
End Property

'*-----------------------------------------------------------------------------
'* Value プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・。
'*-----------------------------------------------------------------------------
Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = mOrigin.Value
End Property


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("VBScript.Match")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************
'なし

