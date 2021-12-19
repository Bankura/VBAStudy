VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRowEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] IRowEditorインタフェースクラス
'* [詳  細] 繰り返しデータに対して行編集を行うクラスのインタフェース。
'*
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2020-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


Public Sub StartProcess(enumerator As Object)
End Sub

Public Sub Edit(enumerator As Object)
End Sub

Public Sub EndProcess(enumerator As Object)
End Sub

Public Sub CancelProcess(enumerator As Object)
End Sub
