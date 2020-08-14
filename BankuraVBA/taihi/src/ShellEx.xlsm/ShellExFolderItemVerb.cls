VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExFolderItemVerb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Shell32.FolderItemVerb�N���X�̃��b�v�E�g���N���X
'* [��  ��] Shell32.FolderItemVerb�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          Shell32.FolderItemVerb�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' Shell32.FolderItemVerb�N���X�ɂ��āB
' <xxxxxxxxxxxxxx>
'
' Definition of interface FolderItemVerb
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mOrigin As shell32.FolderItemVerb
'Private mOrigin As Object

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
'* Application �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EGet Application object
'*-----------------------------------------------------------------------------
Property Get Application() As Object
    Set Application = mOrigin.Application
End Property

'*-----------------------------------------------------------------------------
'* Name �v���p�e�B�i�ǂݎ���p�F���胁���o�[�j
'*
'* [��  ��]
'* �EGet display name for item
'*-----------------------------------------------------------------------------
Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = mOrigin.Name
End Property

'*-----------------------------------------------------------------------------
'* Parent �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EGet Parent object
'*-----------------------------------------------------------------------------
Property Get Parent() As Object
    Set Parent = mOrigin.Parent
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.FolderItemVerb")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
    
'******************************************************************************
'* [�T  �v] DoIt ���\�b�h
'* [��  ��] DoIt �̃��b�p�[���\�b�h�B
'*          Execute the verb.
'*
'******************************************************************************
Sub DoIt()
    Call mOrigin.DoIt
End Sub