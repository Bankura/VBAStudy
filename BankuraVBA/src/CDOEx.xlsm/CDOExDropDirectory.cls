VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExDropDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] CDO.DropDirectory�N���X�̃��b�v�E�g���N���X
'* [��  ��] CDO.DropDirectory�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          CDO.DropDirectory�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* [�Q  �l] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' CDO.DropDirectory�N���X�ɂ��āB
' <xx>
' Defines an object used to access messages located on the file system.
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As CDO.DropDirectory
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
    Set mOrigin = CreateObject("CDO.DropDirectory")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] GetMessages ���\�b�h
'* [��  ��] GetMessages �̃��b�p�[���\�b�h�B
'*          Returns a collection of messages contained in the specified
'*          directory on the file system.
'*          The default location is the SMTP drop directory.
'*
'* @param DirName  xx
'* @return IMessages
'******************************************************************************
Function GetMessages(Optional DirName As String) As CDOExIMessages
    Dim wrap As CDOExIMessages: Set wrap = New CDOExIMessages
    Set wrap.Origin = mOrigin.GetMessages(DirName)
    Set GetMessages = wrap
End Function