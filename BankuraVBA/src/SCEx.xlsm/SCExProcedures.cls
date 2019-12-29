VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SCExProcedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] MSScriptControl.Procedures クラスのラップ・拡張クラス
'* [詳  細] MSScriptControl.Procedures クラスのWrapperとしての機能を提供する他、
'*          MSScriptControl.Procedures クラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' MSScriptControl.Procedures クラスについて。
' Collection of procedures.
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As MSScriptControl.Procedures
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
    Dim myItem As Object
    For Each myItem In mOrigin
        Dim wrap As SCExProcedure: Set wrap = New SCExProcedure
        Set wrap.Origin = myItem
        mKeys.Add wrap, myItem.Name
    Next myItem
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Number of procedures.
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property
    
'*-----------------------------------------------------------------------------
'* Item プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get a procedure indexed by position or procedure name.
'*-----------------------------------------------------------------------------
Property Get Item(Index) As SCExProcedure
Attribute Item.VB_UserMemId = 0
    Dim wrap As SCExProcedure: Set wrap = New SCExProcedure
    Set wrap.Origin = mOrigin.Item(Index)
    Set Item = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("MSScriptControl.Module")
    'Set mKeys = New Collection
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------
'******************************************************************************
'* [概  要] NewEnum メソッド
'* [詳  細] For Each 構文でイテレーターを返すためのメソッド。
'*
'* @return イテレーター
'******************************************************************************
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mKeys.[_NewEnum]
End Function
