VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WScriptExWshURLShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WScript.WshURLShortcut�N���X�̃��b�v�E�g���N���X
'* [��  ��] WScript.WshURLShortcut�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          WScript.WshURLShortcut�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' WScript.WshShortcut�N���X�ɂ��āB
' URL �V���[�g�J�b�g�ւ̃I�u�W�F�N�g�Q�Ƃ��쐬���܂��B
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364464(v=msdn.10)>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As IWshRuntimeLibrary.WshURLShortcut
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

'*-----------------------------------------------------------------------------
'* FullName �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E�V���[�g�J�b�g �I�u�W�F�N�g�̃����N��ւ̐�΃p�X��Ԃ��܂��B
'*-----------------------------------------------------------------------------
Property Get FullName() As String
    FullName = mOrigin.FullName
End Property

'*-----------------------------------------------------------------------------
'* TargetPath �v���p�e�B
'*
'* [��  ��]
'* �E�V���[�g�J�b�g�̎��s�\�t�@�C���ւ̃p�X�ł��B
'*-----------------------------------------------------------------------------
Property Get TargetPath() As String
    TargetPath = mOrigin.TargetPath
End Property
Property Let TargetPath(arg As String)
    mOrigin.TargetPath = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WScript.WshURLShortcut")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] Save ���\�b�h
'* [��  ��] Save �̃��b�p�[���\�b�h�B
'*          �V���[�g�J�b�g �I�u�W�F�N�g��ۑ����܂��B
'*
'******************************************************************************
Sub Save()
    Call mOrigin.Save
End Sub

'******************************************************************************
'* [�T  �v] Load ���\�b�h
'* [��  ��] Load �̃��b�p�[���\�b�h�B
'*          �B�����\�b�h�B
'*
'* @param PathLink
'******************************************************************************
Sub Load(PathLink As String)
    Call mOrigin.Load(PathLink)
End Sub