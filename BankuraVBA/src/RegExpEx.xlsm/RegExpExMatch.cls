VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExpExMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] VBScript.Match�N���X�̃��b�v�E�g���N���X
'* [��  ��] VBScript.Match�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          VBScript.Match�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' VBScript.Match�N���X�ɂ��āB
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc392427(v=msdn.10)>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As VBScript_RegExp_55.Match
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
'* FirstIndex �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E�B
'*-----------------------------------------------------------------------------
Property Get FirstIndex() As Long
    FirstIndex = mOrigin.FirstIndex
End Property

'*-----------------------------------------------------------------------------
'* Length �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E�B
'*-----------------------------------------------------------------------------
Property Get Length() As Long
    Length = mOrigin.Length
End Property

'*-----------------------------------------------------------------------------
'* SubMatches �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E�B
'*-----------------------------------------------------------------------------
Property Get SubMatches() As RegExpExSubMatches
    Dim wrap As RegExpExSubMatches: Set wrap = New RegExpExSubMatches
    Set wrap.Origin = mOrigin.SubMatches()
    Set SubMatches = wrap
End Property

'*-----------------------------------------------------------------------------
'* Value �v���p�e�B�i�ǂݎ���p:���胁���o�[�j
'*
'* [��  ��]
'* �E�B
'*-----------------------------------------------------------------------------
Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = mOrigin.Value
End Property


'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("VBScript.Match")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
'�Ȃ�
