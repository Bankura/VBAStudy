VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RowEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IEnumerator

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �s�J��Ԃ������N���X
'* [��  ��] �s���ɌJ��Ԃ��������s���B
'*
'* [�Q  �l] https://qiita.com/kazu3516/items/1223ce0838dd4a8d2629
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mSourceData As Variant
Private mDestTarget As Variant
Private mCurrentRow As Long
Private mEventObject As EnumeratorEventObject
Private mStartRow As Long
Private mEndRow As Long
Private mStartCol As Long
Private mEndCol As Long
Private mReportUnit As Long
Private mCurrentReportCount As Long
Private mMaxReportCount As Long
Private mIsActionable As Long
Private mRowProcessFunction As Func
Private mStartProcessFunction As Func
Private mEndProcessFunction As Func
Private mRowEditor As IRowEditor


'******************************************************************************
'* �C�x���g��`
'******************************************************************************
Public Event StartAction(sender As RowEnumerator)
Public Event RowProcessAction(sender As RowEnumerator)
Public Event EndAction(sender As RowEnumerator)

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* SourceData �v���p�e�B
'*
'* ���ƂȂ�f�[�^
'*-----------------------------------------------------------------------------
Public Property Get SourceData() As Variant
    If IsObject(mSourceData) Then
        Set SourceData = mSourceData
    Else
        SourceData = mSourceData
    End If
End Property
Public Property Let SourceData(arg As Variant)
    If IsObject(arg) Then
        Set mSourceData = arg
    Else
        mSourceData = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* DestTarget �v���p�e�B
'*
'* �o�͐�
'*-----------------------------------------------------------------------------
Public Property Get DestTarget() As Variant
    If IsObject(mDestTarget) Then
        Set DestTarget = mDestTarget
    Else
        DestTarget = mDestTarget
    End If
End Property
Public Property Let DestTarget(arg As Variant)
    If IsObject(arg) Then
        Set mDestTarget = arg
    Else
        mDestTarget = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* CurrentRow �v���p�e�B
'*
'* ���ݍs�ԍ�
'*-----------------------------------------------------------------------------
Public Property Get CurrentRow() As Long
    CurrentRow = mCurrentRow
End Property

'*-----------------------------------------------------------------------------
'* StartRow �v���p�e�B
'*
'* �J�n�s�ԍ�
'*-----------------------------------------------------------------------------
Public Property Get StartRow() As Long
    StartRow = mStartRow
End Property
Public Property Let StartRow(arg As Long)
    mStartRow = arg
End Property

'*-----------------------------------------------------------------------------
'* EndRow �v���p�e�B
'*
'* �I���s�ԍ�
'*-----------------------------------------------------------------------------
Public Property Get EndRow() As Long
    EndRow = mEndRow
End Property
Public Property Let EndRow(arg As Long)
    mEndRow = arg
End Property

'*-----------------------------------------------------------------------------
'* ReportUnit �v���p�e�B
'*
'* �i���P�ʁB�i���̃��|�[�g���s���i�C�x���g�𔭐�������j�P�ʂ̏������B
'*-----------------------------------------------------------------------------
Public Property Get ReportUnit() As Long
    ReportUnit = mReportUnit
End Property
Public Property Let ReportUnit(arg As Long)
    mReportUnit = arg
End Property

'*-----------------------------------------------------------------------------
'* CurrentReportCount �v���p�e�B
'*
'* ���݂̐i��������
'*-----------------------------------------------------------------------------
Public Property Get CurrentReportCount() As Long
    CurrentReportCount = mCurrentReportCount
End Property
Public Property Let CurrentReportCount(arg As Long)
    mCurrentReportCount = arg
End Property

'*-----------------------------------------------------------------------------
'* MaxReportCount �v���p�e�B
'*
'* �ő又���Ώی����i�i���\���p�j
'*-----------------------------------------------------------------------------
Public Property Get MaxReportCount() As Long
    MaxReportCount = mMaxReportCount
End Property
Public Property Let MaxReportCount(arg As Long)
    mMaxReportCount = arg
End Property

'*-----------------------------------------------------------------------------
'* IsActionable �v���p�e�B
'*
'* �s����Action�C�x���g�𔭐������邩�ǂ����B
'*-----------------------------------------------------------------------------
Public Property Get IsActionable() As Boolean
    IsActionable = mIsActionable
End Property
Public Property Let IsActionable(arg As Boolean)
    mIsActionable = arg
End Property

'*-----------------------------------------------------------------------------
'* RowProcessFunction �v���p�e�B
'*
'* �s���̏������s��Function�B�ݒ肷��ƍs���ɌĂяo�����s����B
'*-----------------------------------------------------------------------------
Public Property Get RowProcessFunction() As Func
    RowProcessFunction = mRowProcessFunction
End Property
Public Property Set RowProcessFunction(arg As Func)
    Set mRowProcessFunction = arg
End Property

'*-----------------------------------------------------------------------------
'* StartProcessFunction �v���p�e�B
'*
'* �J�n���̏������s��Function�B�ݒ肷��ƊJ�n���ɌĂяo�����s����B
'*-----------------------------------------------------------------------------
Public Property Get StartProcessFunction() As Func
    StartProcessFunction = mStartProcessFunction
End Property
Public Property Set StartProcessFunction(arg As Func)
    Set mStartProcessFunction = arg
End Property

'*-----------------------------------------------------------------------------
'* EndProcessFunction �v���p�e�B
'*
'* �I�����̏������s��Function�B�ݒ肷��ƏI�����ɌĂяo�����s����B
'*-----------------------------------------------------------------------------
Public Property Get EndProcessFunction() As Func
    EndProcessFunction = mEndProcessFunction
End Property
Public Property Set EndProcessFunction(arg As Func)
    Set mEndProcessFunction = arg
End Property

'*-----------------------------------------------------------------------------
'* RowEditor �v���p�e�B
'*
'* �s���̏������s���I�u�W�F�N�g�B�ݒ肷��ƍs���ɌĂяo�����s����B
'* �J�n�E�I�����̏����ł��Ή����郁�\�b�h�̌Ăяo�����s����B
'*-----------------------------------------------------------------------------
Public Property Get RowEditor() As IRowEditor
    RowEditor = mRowEditor
End Property
Public Property Set RowEditor(arg As IRowEditor)
    Set mRowEditor = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mEventObject = New EnumeratorEventObject
    mIsActionable = True
    mReportUnit = 1
    mCurrentReportCount = 0
    mMaxReportCount = 0
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] Start
'* [��  ��] �s���ɌJ��Ԃ��������s���B
'*
'******************************************************************************
Public Sub Start()
    
    ' �J�n���̏���
    Call StartProcess
        
    ' �J��Ԃ�����
    Dim i As Long
    For i = mStartRow To mEndRow
        mCurrentReportCount = mCurrentReportCount + 1
        mCurrentRow = i
        
        ' �s����
        Call RowProcess
        
        ' �i�����|�[�g
        Call ReportProgress(i)
    Next
    
    ' �I�����̏���
    Call EndProcess
End Sub

'******************************************************************************
'* [�T  �v] StartProcess
'* [��  ��] �J�n���̏����������s���i�ݒ�E�Ǎ��������j�B
'*
'******************************************************************************
Private Sub StartProcess()
    If Not mStartProcessFunction Is Nothing Then
        Call mStartProcessFunction.Apply(Me)
    End If
    If Not mRowEditor Is Nothing Then
        Call mRowEditor.StartProcess(Me)
    End If
    If IsActionable Then
        RaiseEvent StartAction(Me)
    End If
End Sub

'******************************************************************************
'* [�T  �v] RowProcess
'* [��  ��] �s�ɑ΂��鏈�����s���i�ҏW�������j�B
'*
'******************************************************************************
Private Sub RowProcess()
    If Not mRowProcessFunction Is Nothing Then
        Call mRowProcessFunction.Apply(Me)
    End If
    If Not mRowEditor Is Nothing Then
        Call mRowEditor.Edit(Me)
    End If
    If IsActionable Then
        RaiseEvent RowProcessAction(Me)
    End If
End Sub

'******************************************************************************
'* [�T  �v] ReportProgress
'* [��  ��] �i���̃��|�[�g�����i�C�x���g�ʒm�j���s���B
'*
'******************************************************************************
Private Sub ReportProgress(i As Long)
    If i Mod mReportUnit = 0 Then
        mEventObject.Raise Me
        UXUtils.CheckEvents
    End If
End Sub

'******************************************************************************
'* [�T  �v] EndProcess
'* [��  ��] �I�����̏������s���i�o�͏������j�B
'*
'******************************************************************************
Private Sub EndProcess()
    If IsActionable Then
        RaiseEvent EndAction(Me)
    End If
    If Not mEndProcessFunction Is Nothing Then
        Call mEndProcessFunction.Apply(Me)
    End If
    If Not mRowEditor Is Nothing Then
        Call mRowEditor.EndProcess(Me)
    End If
End Sub

'******************************************************************************
'* �C���^�t�F�[�X����
'******************************************************************************
'*-----------------------------------------------------------------------------
'* IEnumerator#EventObject
'*
'*-----------------------------------------------------------------------------
Private Property Get IEnumerator_EventObject() As EnumeratorEventObject
    Set IEnumerator_EventObject = mEventObject
End Property

'*-----------------------------------------------------------------------------
'* IEnumerator#CurrentReportCount
'*
'*-----------------------------------------------------------------------------
Private Property Get IEnumerator_CurrentReportCount() As Long
    IEnumerator_CurrentReportCount = mCurrentReportCount
End Property
Private Property Let IEnumerator_CurrentReportCount(arg As Long)
    mCurrentReportCount = arg
End Property

'*-----------------------------------------------------------------------------
'* IEnumerator#MaxReportCount
'*
'*-----------------------------------------------------------------------------
Private Property Get IEnumerator_MaxReportCount() As Long
    IEnumerator_MaxReportCount = mMaxReportCount
End Property
Private Property Let IEnumerator_MaxReportCount(arg As Long)
    mMaxReportCount = arg
End Property

'*-----------------------------------------------------------------------------
'* IEnumerator#Start
'*
'*-----------------------------------------------------------------------------
Private Sub IEnumerator_Start()
    Me.Start
End Sub
