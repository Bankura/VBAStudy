VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SCExProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] MSScriptControl.Procedure クラスのラップ・拡張クラス
'* [詳  細] MSScriptControl.Procedure クラスのWrapperとしての機能を提供する他、
'*          MSScriptControl.Procedure クラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' MSScriptControl.Procedure クラスについて。
' Describes a procedure.
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As MSScriptControl.Procedure
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
'* HasReturnValue プロパティ（読み取り専用）
'*
'* [補  足]
'* ・True if procedure returns a value.
'*-----------------------------------------------------------------------------
Property Get HasReturnValue() As Boolean
    HasReturnValue = mOrigin.HasReturnValue
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・Name of the procedure.
'*-----------------------------------------------------------------------------
Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = mOrigin.Name
End Property
    
'*-----------------------------------------------------------------------------
'* NumArgs プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Number of arguments that are expected.
'*-----------------------------------------------------------------------------
Property Get NumArgs() As Long
    NumArgs = mOrigin.NumArgs
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("MSScriptControl.Procedure")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************
'なし
