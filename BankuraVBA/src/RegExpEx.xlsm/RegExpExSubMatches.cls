VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExpExSubMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] VBScript.SubMatchesクラスのラップ・拡張クラス
'* [詳  細] VBScript.SubMatchesクラスのWrapperとしての機能を提供する他、
'*          VBScript.SubMatchesクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' VBScript.SubMatchesクラスについて。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc392216(v=msdn.10)>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As VBScript_RegExp_55.SubMatches
Private mOrigin As Object
Private mKeys As Collection

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
    Set mKeys = New Collection
    Dim myObj As Object
    For Each myObj In mOrigin
        Dim wrap As RegExpExMatch: Set wrap = New RegExpExMatch
        Set wrap.Origin = myObj
        mKeys.Add wrap
    Next myObj
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・｡
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'*-----------------------------------------------------------------------------
'* Item プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・。
'*-----------------------------------------------------------------------------
Property Get Item(index As Long)
Attribute Item.VB_UserMemId = 0
    If IsObject(mOrigin.Item(index)) Then
        Set Item = mOrigin.Item(index)
    Else
        Item = mOrigin.Item(index)
    End If
End Property

   
'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("VBScript.SubMatches")
    'Set mKeys = New Collection
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] NewEnum メソッド
'* [詳  細] For Each 構文でイテレーターを返すためのメソッド。
'*
'* @return イテレーター
'******************************************************************************
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Set NewEnum = mOrigin.[_NewEnum]
    Set NewEnum = mKeys.[_NewEnum]
End Function
