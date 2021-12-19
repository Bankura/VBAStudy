VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] IEnumeratorインタフェースクラス
'* [詳  細] 繰り返しデータの処理を行うクラスのインタフェース。
'*
'* [参  考] https://qiita.com/kazu3516/items/1223ce0838dd4a8d2629
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


Public Property Get eventObject() As EnumeratorEventObject
End Property

Public Property Get CurrentReportCount() As Long
End Property
Public Property Let CurrentReportCount(arg As Long)
End Property

Public Property Get MaxReportCount() As Long
End Property
Public Property Let MaxReportCount(arg As Long)
End Property

Public Property Get IsCancel() As Boolean
End Property
Public Property Let IsCancel(arg As Boolean)
End Property

Public Property Get IsError() As Boolean
End Property
Public Property Let IsError(arg As Boolean)
End Property

Public Sub Start()
End Sub

