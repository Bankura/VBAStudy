VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WbemScripting.SWbemProperty�N���X�̃��b�v�E�g���N���X
'* [��  ��] WbemScripting.SWbemProperty�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          WbemScripting.SWbemProperty�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' WbemScripting.SWbemProperty�N���X�ɂ��āB
' <https://msdn.microsoft.com/ja-jp/windows/aa393804(v=vs.80)>
' A Property
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemProperty
Private mOrigin As Object
    
'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* �I���W�i��Object �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Original() As Object
    Set Original = mOrigin
End Property
Property Set Original(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* CIMType �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe CIM Type of this Property.
'*-----------------------------------------------------------------------------
Property Get CIMType() As WbemCimtypeEnum
    CIMType = mOrigin.CIMType
End Property

'*-----------------------------------------------------------------------------
'* IsArray �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EIndicates whether this Property is an array type.
'*-----------------------------------------------------------------------------
Property Get IsArray() As Boolean
    IsArray = mOrigin.IsArray
End Property

'*-----------------------------------------------------------------------------
'* IsLocal �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EIndicates whether this Property is local or propagated.
'*-----------------------------------------------------------------------------
Property Get IsLocal() As Boolean
    IsLocal = mOrigin.IsLocal
End Property
    
'*-----------------------------------------------------------------------------
'* Name �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe name of this Property.
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property
    
'*-----------------------------------------------------------------------------
'* Origin �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe originating class of this Property.
'*-----------------------------------------------------------------------------
Property Get Origin() As String
    Origin = mOrigin.Origin
End Property

'*-----------------------------------------------------------------------------
'* Qualifiers_ �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe collection of Qualifiers of this Method.
'*-----------------------------------------------------------------------------
Property Get Qualifiers_() As WMIExSWbemQualifierSet
    Dim wrap As WMIExSWbemQualifierSet: Set wrap = New WMIExSWbemQualifierSet
    Set wrap.Origin = mOrigin.Qualifiers_
    Set Qualifiers_ = wrap
End Property
    
'*-----------------------------------------------------------------------------
'* Value �v���p�e�B�i���胁���o�[�j
'*
'* [��  ��]
'* �EThe value of this Property.
'*-----------------------------------------------------------------------------
Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    If IsObject(mOrigin.Value) Then
        Set Value = mOrigin.Value
    Else
        Value = mOrigin.Value
    End If
End Property
Property Set Value(arg As Variant)
    If IsObject(arg) Then
        Set mOrigin.Value = arg
    Else
        mOrigin.Value = arg
    End If
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemProperty")
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

