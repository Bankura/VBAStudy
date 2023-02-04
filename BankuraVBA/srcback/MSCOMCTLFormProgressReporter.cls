VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSCOMCTLFormProgressReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IProgressReporter

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �i���o�[�t�H�[���i���\���N���X
'* [��  ��] �i���o�[�t�H�[���ɐi���󋵂�\������B
'*
'* [�Q  �l]
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mProgressBarForm As MSCOMCTLProgressBarForm
Private mBaseMessage As String
Private mProgressString As String
Private mCountVisible As Boolean

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* BaseMessage �v���p�e�B
'*
'* �i���\�����b�Z�[�W�̌Œ蕔���i�uCSV�Ǎ��������v���j
'*-----------------------------------------------------------------------------
Public Property Let BaseMessage(arg As String)
    mBaseMessage = arg
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
    Set mProgressBarForm = New MSCOMCTLProgressBarForm
    mBaseMessage = "������"
    mCountVisible = True
End Sub
Private Sub Class_Terminate()
    Set mProgressBarForm = Nothing
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
    With mProgressBarForm
        .ProcessTitle = mBaseMessage
        With .ProgressBar
            .Min = 0
            .Max = enumerator.MaxReportCount
            .Value = 0
        End With
        .ProgressText.Caption = GetProgressMessage(enumerator)
        .Show vbModeless
    End With
End Sub

'******************************************************************************
'* [�T  �v] Finish
'* [��  ��] �I�����̕\���������s���B
'*
'* @param enumerator IEnumerator�I�u�W�F�N�g
'******************************************************************************
Public Sub Finish(enumerator As IEnumerator)
    mProgressBarForm.Hide
End Sub

'******************************************************************************
'* [�T  �v] Proceed
'* [��  ��] �i�����i�񂾍ۂ̕\���������s���B
'*
'* @param enumerator IEnumerator�I�u�W�F�N�g
'******************************************************************************
Public Sub Proceed(enumerator As IEnumerator)
    With mProgressBarForm
         .ProgressBar.Value = enumerator.CurrentReportCount
         .ProgressText.Caption = GetProgressMessage(enumerator)
         .Repaint
    End With
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
    If mCountVisible Then
        msg.Append enumerator.CurrentReportCount
        msg.Append " / "
        msg.Append enumerator.MaxReportCount
        msg.Append " ��"
    End If
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
