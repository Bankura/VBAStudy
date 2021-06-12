VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SBProgressReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IProgressReporter

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �X�e�[�^�X�o�[�i���\���N���X
'* [��  ��] �X�e�[�^�X�o�[�ɐi���󋵂�\������B
'*
'* [�Q  �l]
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mScaleChar As String
Private mBaseMessage As String
Private mProgressString As String
Private mBaseVisible As Boolean
Private mCountVisible As Boolean

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* ScaleChar �v���p�e�B
'*
'* �ڐ���\���̕����i�u���v���j
'*-----------------------------------------------------------------------------
Public Property Let ScaleChar(arg As String)
    mScaleChar = arg
End Property

'*-----------------------------------------------------------------------------
'* BaseMessage �v���p�e�B
'*
'* �i���\�����b�Z�[�W�̌Œ蕔���i�uCSV�Ǎ��������F�v���j
'*-----------------------------------------------------------------------------
Public Property Let BaseMessage(arg As String)
    mBaseMessage = arg
End Property

'*-----------------------------------------------------------------------------
'* BaseVisible �v���p�e�B
'*
'* �Œ胁�b�Z�[�W��\�����邩
'*-----------------------------------------------------------------------------
Public Property Let BaseVisible(arg As Boolean)
    mBaseVisible = arg
End Property

'*-----------------------------------------------------------------------------
'* CountVisible �v���p�e�B
'*
'* �����\�������邩
'*-----------------------------------------------------------------------------
Public Property Let CountVisible(arg As Boolean)
    mCountVisible = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    mScaleChar = "��"
    mBaseMessage = "������"
    mBaseVisible = True
    mCountVisible = True
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
'******************************************************************************
'* [�T  �v] Start
'* [��  ��] �J�n���̕\���������s���B
'*
'* @param enumerator IEnumerator�I�u�W�F�N�g
'******************************************************************************
Public Sub Start(enumerator As IEnumerator)
    Application.StatusBar = GetProgressMessage(enumerator)
End Sub

'******************************************************************************
'* [�T  �v] Finish
'* [��  ��] �I�����̕\���������s���B
'*
'* @param enumerator IEnumerator�I�u�W�F�N�g
'******************************************************************************
Public Sub Finish(enumerator As IEnumerator)
    Application.StatusBar = False
End Sub

'******************************************************************************
'* [�T  �v] Proceed
'* [��  ��] �i�����i�񂾍ۂ̕\���������s���B
'*
'* @param enumerator IEnumerator�I�u�W�F�N�g
'******************************************************************************
Public Sub Proceed(enumerator As IEnumerator)
    mProgressString = mProgressString + mScaleChar
    Application.StatusBar = GetProgressMessage(enumerator)
End Sub

'******************************************************************************
'* [�T  �v] GetProgressMessage
'* [��  ��] �i�����b�Z�[�W���擾����B
'*
'* @param enumerator IEnumerator�I�u�W�F�N�g
'* @return String �i�����b�Z�[�W
'******************************************************************************
Private Function GetProgressMessage(enumerator As IEnumerator) As String
    Dim msg As New StringEx
    If mBaseVisible Then
        msg.Append mBaseMessage
        If mCountVisible Then
            msg.Append "( "
            msg.Append enumerator.CurrentReportCount
            msg.Append " / "
            msg.Append enumerator.MaxReportCount
            msg.Append " ��)"
        End If
        msg.Append "�F "
    End If
    msg.Append mProgressString
            
    GetProgressMessage = msg.ToString
End Function

'******************************************************************************
'* �C���^�t�F�[�X����
'******************************************************************************
'*-----------------------------------------------------------------------------
'* IProgressReporter#Start
'*
'*-----------------------------------------------------------------------------
Private Sub IProgressReporter_Start(enumerator As IEnumerator)
    Me.Start enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IProgressReporter#Finish
'*
'*-----------------------------------------------------------------------------
Private Sub IProgressReporter_Finish(enumerator As IEnumerator)
    Me.Finish enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IProgressReporter#Proceed
'*
'*-----------------------------------------------------------------------------
Private Sub IProgressReporter_Proceed(enumerator As IEnumerator)
    Me.Proceed enumerator
End Sub