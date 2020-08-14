VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExINNTPOnPostEarly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CDO.INNTPOnPostEarlyクラスのラップ・拡張クラス
'* [詳  細] CDO.INNTPOnPostEarlyクラスのWrapperとしての機能を提供する他、
'*          CDO.INNTPOnPostEarlyクラスを使用したユーティリティを提供する。
'*
'* [参  考] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' CDO.INNTPOnPostEarlyクラスについて。
' <xx>
' Implement when creating NNTP OnPostEarly event sinks.
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As CDO.INNTPOnPostEarly
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
    'Set mOrigin = CreateObject("CDO.INNTPOnPostEarly")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] OnPostEarly メソッド
'* [詳  細] OnPostEarly のラッパーメソッド。
'*          Called by the NNTP event dispatcher when message headers arrive.
'*
'* @param Msg
'* @param EventStatus
'*
'******************************************************************************
Sub OnPostEarly(Msg As CDOExMessage, EventStatus As CdoEventStatus)
    Call mOrigin.OnPostEarly(Msg.Origin, EventStatus)
End Sub
