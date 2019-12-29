VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExISMTPOnArrival"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CDO.ISMTPOnArrivalクラスのラップ・拡張クラス
'* [詳  細] CDO.ISMTPOnArrivalクラスのWrapperとしての機能を提供する他、
'*          CDO.ISMTPOnArrivalクラスを使用したユーティリティを提供する。
'*
'* [参  考] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' CDO.ISMTPOnArrivalクラスについて。
' <xx>
' Implement when creating SMTP OnArrival event sinks.
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As CDO.ISMTPOnArrival
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
    'Set mOrigin = CreateObject("CDO.ISMTPOnArrival")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] OnArrival メソッド
'* [詳  細] OnArrival のラッパーメソッド。
'*          Called by the SMTP event dispatcher when a message arrives.
'*
'* @param Msg
'* @param EventStatus
'*
'******************************************************************************
Sub OnArrival(Msg As CDOExMessage, EventStatus As CdoEventStatus)
    Call mOrigin.OnArrival(Msg.Origin, EventStatus)
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

