VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WbemScripting.SWbemMethod�N���X�̃��b�v�E�g���N���X
'* [��  ��] WbemScripting.SWbemMethod�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          WbemScripting.SWbemMethod�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' WbemScripting.SWbemMethod�N���X�ɂ��āB
' <https://msdn.microsoft.com/ja-jp/windows/aa393722(v=vs.80)>
' A Method
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemMethod
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
'* InParameters �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe in parameters for this Method.
'*-----------------------------------------------------------------------------
Property Get InParameters() As WMIExSWbemObject
    Dim wrap As WMIExSWbemObject: Set wrap = New WMIExSWbemObject
    Set wrap.Origin = mOrigin.InParameters
    Set InParameters = wrap
End Property

'*-----------------------------------------------------------------------------
'* Name �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe name of this Method.
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property

'*-----------------------------------------------------------------------------
'* Origin �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe originating class of this Method.
'*-----------------------------------------------------------------------------
Property Get Origin() As String
    Origin = mOrigin.Origin
End Property

'*-----------------------------------------------------------------------------
'* OutParameters �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe out parameters for this Method.
'*-----------------------------------------------------------------------------
Property Get OutParameters() As WMIExSWbemObject
    Dim wrap As WMIExSWbemObject: Set wrap = New WMIExSWbemObject
    Set wrap.Origin = mOrigin.OutParameters
    Set OutParameters = wrap
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
    
'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemMethod")
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