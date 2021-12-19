VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBarForm 
   Caption         =   "���΂炭���҂����������c"
   ClientHeight    =   1456
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4298
   OleObjectBlob   =   "ProgressBarForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ProgressBarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �i���o�[�t�H�[��
'* [��  ��] �i���󋵂�\������i���o�[�t�H�[���B
'*          FormProgressReporter�N���X����g�p����z��B
'* [�Q  �l]
'*          https://excel-ubara.com/excelvba3/EXCELFORM026.html
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Public IsCancel As Boolean
Private mProgressBarLabel As MSForms.Label
Private mMaxValue As Long
Private mBarColor As Long
Private mCurValue As Double  ' �v���O���X�o�[���ݒl
Private mInteractive As Long
Private mSelfDoEvents As Boolean

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* MaxValue �v���p�e�B
'*
'* �ő�l�v���p�e�B
'*-----------------------------------------------------------------------------
Public Property Let MaxValue(arg As Long)
    If arg = 0 Then Exit Property
    mMaxValue = arg
End Property
Public Property Get MaxValue() As Long
    MaxValue = mMaxValue
End Property

'*-----------------------------------------------------------------------------
'* BarColor �v���p�e�B
'*
'* �v���O���X�o�[�̐F�w��
'*-----------------------------------------------------------------------------
Public Property Let BarColor(arg As Long)
    mBarColor = arg
    mProgressBarLabel.BackColor = mBarColor
End Property

'*-----------------------------------------------------------------------------
'* Interactive �v���p�e�B
'*
'* �����݋��ێw��iFalse: ���ہj
'*-----------------------------------------------------------------------------
Public Property Let Interactive(arg As Boolean)
    mInteractive = arg
End Property

'*-----------------------------------------------------------------------------
'* SelfDoEvents �v���p�e�B
'*
'* Form���g�̏�����DoEvents���Ăяo����
'*-----------------------------------------------------------------------------
Public Property Let SelfDoEvents(arg As Boolean)
    mSelfDoEvents = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub UserForm_Initialize()
    mBarColor = rgb(0, 0, 128)
    mInteractive = True
    IsCancel = False
    mSelfDoEvents = True
    
    mMaxValue = 100
    mCurValue = 0

    ' ���x���R���g���[���ǉ�
    Set mProgressBarLabel = Me.ProgressBar.Controls.Add("Forms.Label.1", "lblProgress")
    mProgressBarLabel.Width = 0
    mProgressBarLabel.Height = Me.ProgressBar.Height
    mProgressBarLabel.BackColor = mBarColor

    ' �v���O���X�o�[�̔w�i���ւ��܂���
    Me.ProgressBar.SpecialEffect = fmSpecialEffectSunken
    
    Me.Caption = "���΂炭���҂����������c"
    Me.ProgressText.Caption = ""
    
End Sub
Private Sub UserForm_Terminate()
    If mInteractive = False Then
        Application.Interactive = True
        Application.EnableCancelKey = xlInterrupt
    End If
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
'******************************************************************************
'* [�T  �v] ShowModeless
'* [��  ��] �t�H�[�������[�h���X�ŕ\������B
'*
'* @param formCaptionTxt �t�H�[���̃^�C�g���e�L�X�g
'******************************************************************************
Public Sub ShowModeless(Optional ByVal formCaptionTxt As String)
    ' �����݋��ۂ̐ݒ�
    If mInteractive = False Then
        Me.Enabled = False
        Application.Interactive = False
        Application.EnableCancelKey = xlDisabled
    End If
  
    ' �t�H�[�������[�h���X�ŕ\��
    If formCaptionTxt <> "" Then
        Me.Caption = formCaptionTxt
    End If
    Me.Show vbModeless
End Sub

'******************************************************************************
'* [�T  �v] SetProgressValue
'* [��  ��] �v���O���X�o�[�̐i�����w�肵���l�ōX�V����B
'*
'* @param aValue �i���l�i�w��l�j
'* @param progressTxt �i���\���̃e�L�X�g
'* @param formCaptionTxt �t�H�[���̃^�C�g���e�L�X�g
'******************************************************************************
Public Sub SetProgressValue(ByVal aValue As Double, Optional ByVal progressTxt As String, Optional ByVal formCaptionTxt As String)
    mCurValue = aValue
  
    ' �ő�l�𒴂��Ȃ��悤�ɒ���
    If mCurValue > mMaxValue Then
        mCurValue = mMaxValue
    End If
  
    ' �v���O���X�o�[�̕`��
    mProgressBarLabel.BackColor = mBarColor
    mProgressBarLabel.Width = Me.ProgressBar.Width * (mCurValue / mMaxValue)
    If formCaptionTxt <> "" Then
        Me.Caption = formCaptionTxt
    End If
    If progressTxt <> "" Then
        Me.ProgressText.Caption = progressTxt
    End If
    
    ' �ĕ`��
    Me.Repaint
    If mSelfDoEvents Then
        UXUtils.CheckEvents
    End If
End Sub

'******************************************************************************
'* [�T  �v] AddProgressValue
'* [��  ��] �v���O���X�o�[�̐i�����w�肵���l�ŉ��Z���A�X�V����B
'*
'* @param aValue �i���l�i���Z�l�j
'* @param progressTxt �i���\���̃e�L�X�g
'* @param formCaptionTxt �t�H�[���̃^�C�g���e�L�X�g
'******************************************************************************
Public Sub AddProgressValue(ByVal aValue As Double, Optional ByVal progressTxt As String, Optional ByVal formCaptionTxt As String)
    mCurValue = mCurValue + aValue
    Call SetProgressValue(mCurValue, formCaptionTxt)
End Sub

'******************************************************************************
'* [�T  �v] Unload
'* [��  ��] ���g�̃t�H�[����Unload����B
'*
'******************************************************************************
Public Sub Unload()
    VBA.Unload Me
End Sub

'******************************************************************************
'* �C�x���g����
'******************************************************************************

'******************************************************************************
'* [�T  �v] UserForm�FQueryClose �C�x���g����
'* [��  ��] �t�H�[��������O�ɔ�������QueryClose �C�x���g�̏����B
'*          �u�~�v�{�^�����̃��[�U����ɂ��t�H�[����������ۂɁA
'*          �����݋��ێw��v���p�e�B���uTrue�v�i���荞�݂����e�j�ł���ꍇ�́A
'*          �����̒��f���m�F����_�C�A���O��\������B
'*          �_�C�A���O�ɂāu�͂��v�I�����́AIsCancel�v���p�e�B���uTrue�v�i���f�j
'*          �ɐݒ肵�A�t�H�[�������B
'*          �����݋��ێw��v���p�e�B���uFalse�v�i���荞�݂����ہj�ł���ꍇ�A
'*          �_�C�A���O�ɂāu�������v�I�����́A�t�H�[������鏈�����L�����Z��
'*          ����B
'* [�Q  �l] https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/queryclose-event
'*
'* @param Cancel �C�x���g���L�����Z�����邩�B���� ������ 0 �ȊO�̒l�ɐݒ肷���
'*               �Ǎ��ς̂��ׂẴ��[�U�[�t�H�[���� QueryClose �C�x���g����~��
'*               ��AUserForm �ƃA�v���P�[�V��������鏈�����L�����Z�������B
'* @param CloseMode QueryClose �C�x���g�̌����������l�܂��� �萔
'******************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If mInteractive Then
            If MsgBox("�����𒆒f���܂���?", vbYesNo, "���f�m�F") = vbYes Then
                IsCancel = True
                Exit Sub
            End If
        End If
        Cancel = True
    End If
End Sub
