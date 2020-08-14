VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemPrivilege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WbemScripting.SWbemPrivilege�N���X�̃��b�v�E�g���N���X
'* [��  ��] WbemScripting.SWbemPrivilege�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          WbemScripting.SWbemPrivilege�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' WbemScripting.SWbemPrivilege�N���X�ɂ��āB
' <https://msdn.microsoft.com/ja-jp/windows/aa393792(v=vs.80)>
' A Privilege Override
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemPrivilege
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
'* DisplayName �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe display name of the Privilege.
'*-----------------------------------------------------------------------------
Property Get DisplayName() As String
    DisplayName = mOrigin.DisplayName
End Property

'*-----------------------------------------------------------------------------
'* Identifier �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe Privilege identifier.
'*-----------------------------------------------------------------------------
Property Get Identifier() As WbemPrivilegeEnum
    Identifier = mOrigin.Identifier
End Property
    
'*-----------------------------------------------------------------------------
'* IsEnabled �v���p�e�B�i���胁���o�[�j
'*
'* [��  ��]
'* �EWhether the Privilege is to be enabled or disabled.
'*-----------------------------------------------------------------------------
Property Get IsEnabled() As Boolean
Attribute IsEnabled.VB_UserMemId = 0
    IsEnabled = mOrigin.IsEnabled
End Property
Property Let IsEnabled(arg As Boolean)
    mOrigin.IsEnabled = arg
End Property

'*-----------------------------------------------------------------------------
'* Name �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe name of the Privilege.
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemPrivilege")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
'�Ȃ�

'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------
