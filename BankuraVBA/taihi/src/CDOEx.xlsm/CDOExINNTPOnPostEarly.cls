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
'* [�@�\��] CDO.INNTPOnPostEarly�N���X�̃��b�v�E�g���N���X
'* [��  ��] CDO.INNTPOnPostEarly�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          CDO.INNTPOnPostEarly�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* [�Q  �l] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' CDO.INNTPOnPostEarly�N���X�ɂ��āB
' <xx>
' Implement when creating NNTP OnPostEarly event sinks.
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As CDO.INNTPOnPostEarly
Private mOrigin As Object

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* �I���W�i��Object �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("CDO.INNTPOnPostEarly")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] OnPostEarly ���\�b�h
'* [��  ��] OnPostEarly �̃��b�p�[���\�b�h�B
'*          Called by the NNTP event dispatcher when message headers arrive.
'*
'* @param Msg
'* @param EventStatus
'*
'******************************************************************************
Sub OnPostEarly(Msg As CDOExMessage, EventStatus As CdoEventStatus)
    Call mOrigin.OnPostEarly(Msg.Origin, EventStatus)
End Sub