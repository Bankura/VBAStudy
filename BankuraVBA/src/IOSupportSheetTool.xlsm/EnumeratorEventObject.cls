VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumeratorEventObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �C�x���g�����p�N���X
'* [��  ��] �C�x���g�𔭐������邽�߂̃N���X�B
'*
'* [�Q  �l] https://qiita.com/kazu3516/items/1223ce0838dd4a8d2629
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


Public Event OnAction(sender As IEnumerator)

Public Sub Raise(sender As IEnumerator)
    RaiseEvent OnAction(sender)
End Sub