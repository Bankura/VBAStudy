VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExDFConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Shell32.DFConstraintクラスのラップ・拡張クラス
'* [詳  細] Shell32.DFConstraintクラスのWrapperとしての機能を提供する他、
'*          Shell32.DFConstraintクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Shell32.DFConstraintクラスについて。
' Constraint used in search command.
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As shell32.DFConstraint
'Private mOrigin As Object

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
'* Name プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get the constraint name
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property

'*-----------------------------------------------------------------------------
'* Value プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get the constraint Value
'*-----------------------------------------------------------------------------
Property Get Value() As Variant
    Value = mOrigin.Value
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.DFConstraint")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
