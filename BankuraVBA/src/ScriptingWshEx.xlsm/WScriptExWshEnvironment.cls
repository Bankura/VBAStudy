VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WScriptExWshEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WScript.WshEnvironment�N���X�̃��b�v�E�g���N���X
'* [��  ��] WScript.WshEnvironment�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          WScript.WshEnvironment�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' WScript.WshEnvironment�N���X�ɂ��āB
' Windows ���ϐ��̃R���N�V�����ւ̃A�N�Z�X��񋟂��܂��B
' WshEnvironment �I�u�W�F�N�g�͊��ϐ��̃R���N�V�����ł���AWshShell �I�u�W�F
' �N�g�� Environment �v���p�e�B�ɂ���ĕԂ���܂��B
' ���̃R���N�V�����ɂ́A���ϐ��Z�b�g�S�� (���O�t���ϐ�����і��O�Ȃ��ϐ�)
' ���܂܂�܂��B
' �X�̊��ϐ� (����т��̒l) �����̃R���N�V��������擾����ɂ́A���ϐ�
' �����C���f�b�N�X�Ƃ��Ďg�p���܂��B
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364435(v=msdn.10)>�B
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As IWshRuntimeLibrary.WshEnvironment
Private mOrigin As Object
Private mKeys As Collection

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
    Set mKeys = New Collection
    Dim v
    For Each v In mOrigin
        'mKeys.Add Split(v, "=")(1), Split(v, "=")(0)
        mKeys.Add v
    Next
End Property

'*-----------------------------------------------------------------------------
'* Item �v���p�e�B�i���胁���o�[�j
'*
'* [��  ��]
'* �E�w�肵�����̂ɑ΂��鍀�ڂ�ݒ�B
'*-----------------------------------------------------------------------------
Property Get Item(Name As String) As String
Attribute Item.VB_UserMemId = 0
    Item = mOrigin.Item(Name)
End Property
Property Let Item(Name As String, arg As String)
     mOrigin.Item(Name) = arg
     'mKeys.Add arg, Name
     mKeys.Add Name & "=" & arg
End Property

'*-----------------------------------------------------------------------------
'* length �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E�R���N�V�������̃I�u�W�F�N�g���������B
'*-----------------------------------------------------------------------------
Property Get length() As Long
    length = mOrigin.length
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WScript.Environment")
    'Set mKeys = New Collection
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] Count���\�b�h
'* [��  ��] Count�̃��b�p�[���\�b�h�B
'*          �R���N�V�����̗v�f����ԋp�B
'*
'* @return �R���N�V�����̗v�f��
'*
'******************************************************************************
Function Count() As Long
    Count = mOrigin.Count
End Function

'******************************************************************************
'* [�T  �v] Remove���\�b�h
'* [��  ��] Remove�̃��b�p�[���\�b�h�B
'*          �R���N�V�����̗v�f�����O�B
'*
'* @param Name �R���N�V���������珜�O����v�f�̖���
'*
'******************************************************************************
Sub Remove(Name As String)
    Call mOrigin.Remove(Name)
    Dim i As Long
    For i = 1 To mKeys.Count
        'If mKeys(i) = Name Then
        If Split(mKeys(i), "=")(0) = Name Then
            mKeys.Remove i
            Exit Sub
        End If
    Next
End Sub

'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [�T  �v] NewEnum ���\�b�h
'* [��  ��] For Each �\���ŃC�e���[�^�[��Ԃ����߂̃��\�b�h�B
'*
'* @return �C�e���[�^�[
'******************************************************************************
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Set NewEnum = mOrigin.[_NewEnum]
    Set NewEnum = mKeys.[_NewEnum]
End Function