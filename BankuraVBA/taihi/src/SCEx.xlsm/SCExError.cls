VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SCExError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] MSScriptControl.Error クラスのラップ・拡張クラス
'* [詳  細] MSScriptControl.Error クラスのWrapperとしての機能を提供する他、
'*          MSScriptControl.Error クラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' MSScriptControl.Error クラスについて。
' Provides access to scripting error information
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As MSScriptControl.Error
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
'* Column プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Source code column position where the error occurred.
'*-----------------------------------------------------------------------------
Property Get Column() As Long
    Column = mOrigin.Column
End Property

'*-----------------------------------------------------------------------------
'* Description プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Friendly description of error.
'*-----------------------------------------------------------------------------
Property Get Description() As String
    Description = mOrigin.Description
End Property

'*-----------------------------------------------------------------------------
'* HelpContext プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Context ID for the topic with information on the error.
'*-----------------------------------------------------------------------------
Property Get HelpContext() As Long
    HelpContext = mOrigin.HelpContext
End Property

'*-----------------------------------------------------------------------------
'* HelpFile プロパティ（読み取り専用）
'*
'* [補  足]
'* ・File in which help for the error can be found.
'*-----------------------------------------------------------------------------
Property Get HelpFile() As String
    HelpFile = mOrigin.HelpFile
End Property

'*-----------------------------------------------------------------------------
'* Line プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Source code line number where the error occurred.
'*-----------------------------------------------------------------------------
Property Get Line() As Long
    Line = mOrigin.Line
End Property

'*-----------------------------------------------------------------------------
'* Number プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Error Number.
'*-----------------------------------------------------------------------------
Property Get Number() As Long
    Number = mOrigin.Number
End Property

'*-----------------------------------------------------------------------------
'* Source プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Source of the error.
'*-----------------------------------------------------------------------------
Property Get Source() As String
    Source = mOrigin.Source
End Property

'*-----------------------------------------------------------------------------
'* Text プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ Line of source code on which the error occurred.
'*-----------------------------------------------------------------------------
Property Get Text() As String
    Text = mOrigin.Text
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("MSScriptControl.Error")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Clear メソッド
'* [詳  細] Clear のラッパーメソッド。
'*          Clear the script error
'*
'******************************************************************************
Public Sub Clear()
    Call mOrigin.Clear
End Sub

