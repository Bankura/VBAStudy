VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] CDO.Configuration�N���X�̃��b�v�E�g���N���X
'* [��  ��] CDO.Configuration�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          CDO.Configuration�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* [�Q  �l] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' CDO.Configuration�N���X�ɂ��āB
' <xx>
' Defines an object used to store configuration information for CDO objects.
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As CDO.Configuration
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
'* Fields �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �Ethe Object 's Fields collection.
'*-----------------------------------------------------------------------------
Property Get Fields() As ADODBExFields
Attribute Fields.VB_UserMemId = 0
    Dim wrap As ADODBExFields: Set wrap = New ADODBExFields
    Set wrap.Origin = mOrigin.Fields()
    Set Fields = wrap
End Property
    
'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("CDO.Configuration")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] GetInterface ���\�b�h
'* [��  ��] GetInterface �̃��b�p�[���\�b�h�B
'*          Returns a specified interface on this object; provided for script languages.
'*
'* @param Interface  xx
'* @return Object
'******************************************************************************
Function GetInterface(Interface As String) As Object
    Set GetInterface = mOrigin.GetInterface(Interface)
End Function

'******************************************************************************
'* [�T  �v] Load ���\�b�h
'* [��  ��] Load �̃��b�p�[���\�b�h�B
'*          Loads the specified configuration.
'*
'* @param LoadFrom xx
'* @param URL xx
'******************************************************************************
Sub Load(LoadFrom As CdoConfigSource, Optional URL As String)
    Call mOrigin.Load(LoadFrom, URL)
End Sub

'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------