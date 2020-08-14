VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SCExError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] MSScriptControl.Error �N���X�̃��b�v�E�g���N���X
'* [��  ��] MSScriptControl.Error �N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          MSScriptControl.Error �N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' MSScriptControl.Error �N���X�ɂ��āB
' Provides access to scripting error information
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As MSScriptControl.Error
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
'* Column �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �ESource code column position where the error occurred.
'*-----------------------------------------------------------------------------
Property Get Column() As Long
    Column = mOrigin.Column
End Property

'*-----------------------------------------------------------------------------
'* Description �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EFriendly description of error.
'*-----------------------------------------------------------------------------
Property Get Description() As String
    Description = mOrigin.Description
End Property

'*-----------------------------------------------------------------------------
'* HelpContext �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EContext ID for the topic with information on the error.
'*-----------------------------------------------------------------------------
Property Get HelpContext() As Long
    HelpContext = mOrigin.HelpContext
End Property

'*-----------------------------------------------------------------------------
'* HelpFile �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EFile in which help for the error can be found.
'*-----------------------------------------------------------------------------
Property Get HelpFile() As String
    HelpFile = mOrigin.HelpFile
End Property

'*-----------------------------------------------------------------------------
'* Line �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �ESource code line number where the error occurred.
'*-----------------------------------------------------------------------------
Property Get Line() As Long
    Line = mOrigin.Line
End Property

'*-----------------------------------------------------------------------------
'* Number �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EError Number.
'*-----------------------------------------------------------------------------
Property Get Number() As Long
    Number = mOrigin.Number
End Property

'*-----------------------------------------------------------------------------
'* Source �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �ESource of the error.
'*-----------------------------------------------------------------------------
Property Get Source() As String
    Source = mOrigin.Source
End Property

'*-----------------------------------------------------------------------------
'* Text �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E Line of source code on which the error occurred.
'*-----------------------------------------------------------------------------
Property Get Text() As String
    Text = mOrigin.Text
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("MSScriptControl.Error")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] Clear ���\�b�h
'* [��  ��] Clear �̃��b�p�[���\�b�h�B
'*          Clear the script error
'*
'******************************************************************************
Public Sub Clear()
    Call mOrigin.Clear
End Sub
