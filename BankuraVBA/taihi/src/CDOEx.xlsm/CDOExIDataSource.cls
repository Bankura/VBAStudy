VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExIDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] CDO.IDataSource�N���X�̃��b�v�E�g���N���X
'* [��  ��] CDO.IDataSource�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          CDO.IDataSource�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* [�Q  �l] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' CDO.IDataSource�N���X�ɂ��āB
' <xx>
' Defines methods, properties, and collections used to extract messages from or embed messages into other CDO message body parts.
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As CDO.IDataSource
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
End Property

'*-----------------------------------------------------------------------------
'* ActiveConnection �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �ENot implemented. Reserved for future use.
'*-----------------------------------------------------------------------------
'Property Get ActiveConnection() As Connection
'    Set ActiveConnection = mOrigin.ActiveConnection
'End Property

'*-----------------------------------------------------------------------------
'* IsDirty �v���p�e�B
'*
'* [��  ��]
'* �EIndicates whether the local data has been changed since the last save or bind operation.
'*-----------------------------------------------------------------------------
Property Get IsDirty() As Boolean
    IsDirty = mOrigin.IsDirty
End Property
Property Let IsDirty(arg As Boolean)
    mOrigin.IsDirty = arg
End Property
   
'*-----------------------------------------------------------------------------
'* Source �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EReturns the currently bound object. When you bind resources by URL, an ADO _Record interface is returned on an open Record object.
'*-----------------------------------------------------------------------------
Property Get Source() As IUnknown
    Set Source = mOrigin.Source
End Property
    
'*-----------------------------------------------------------------------------
'* SourceClass �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe interface name (type) of the currently bound object. When you bind resources by URL, the value _Record is returned.
'*-----------------------------------------------------------------------------
Property Get SourceClass() As String
    SourceClass = mOrigin.SourceClass
End Property

'*-----------------------------------------------------------------------------
'* SourceURL �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �ENot Implemented. Reserved for future use.
'*-----------------------------------------------------------------------------
'Property Get SourceURL() As String
'    SourceURL = mOrigin.SourceURL
'End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("CDO.IDataSource")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] OpenDs ���\�b�h
'* [��  ��] Open �̃��b�p�[���\�b�h�B
'*          Not implemented. Reserved for future use.
'*
'* @param SourceURL
'* @param ActiveConnection  �ȗ��\�B
'* @param Mode  �ȗ��\�B
'* @param CreateOptions  �ȗ��\�B
'* @param Options  �ȗ��\�B
'* @param UserName  �ȗ��\�B
'* @param Password  �ȗ��\�B
'*
'******************************************************************************
Sub OpenDs(SourceURL As String, Optional ActiveConnection As Object, _
                                Optional Mode As ConnectModeEnum, _
                                Optional CreateOptions As RecordCreateOptionsEnum = adFailIfNotExists, _
                                Optional Options As RecordOpenOptionsEnum, _
                                Optional UserName As String, _
                                Optional Password As String)
    Call mOrigin.Open(SourceURL, ActiveConnection, Mode, CreateOptions, Options, UserName, Password)
End Sub

'******************************************************************************
'* [�T  �v] OpenObject ���\�b�h
'* [��  ��] OpenObject �̃��b�p�[���\�b�h�B
'*          Binds to and opens data from the specified object.
'*
'* @param Source
'* @param InterfaceName
'*
'******************************************************************************
Sub OpenObject(Source As IUnknown, InterfaceName As String)
    Call mOrigin.OpenObject(Source, InterfaceName)
End Sub

'******************************************************************************
'* [�T  �v] Save ���\�b�h
'* [��  ��] Save �̃��b�p�[���\�b�h�B
'*          Saves data into the currently bound object.
'*
'******************************************************************************
Sub Save()
    Call mOrigin.Save
End Sub

'******************************************************************************
'* [�T  �v] SaveTo ���\�b�h
'* [��  ��] SaveTo �̃��b�p�[���\�b�h�B
'*          Not implemented. Reserved for future use.
'*
'* @param SourceURL
'* @param ActiveConnection  �ȗ��\�B
'* @param Mode  �ȗ��\�B
'* @param CreateOptions  �ȗ��\�B
'* @param Options  �ȗ��\�B
'* @param UserName  �ȗ��\�B
'* @param Password  �ȗ��\�B
'*
'******************************************************************************
Sub SaveTo(SourceURL As String, Optional ActiveConnection As Object, _
                                Optional Mode As ConnectModeEnum, _
                                Optional CreateOptions As RecordCreateOptionsEnum = adFailIfNotExists, _
                                Optional Options As RecordOpenOptionsEnum, _
                                Optional UserName As String, _
                                Optional Password As String)
    Call mOrigin.SaveTo(SourceURL, ActiveConnection, Mode, CreateOptions, Options, UserName, Password)
End Sub

'******************************************************************************
'* [�T  �v] SaveToContainer ���\�b�h
'* [��  ��] SaveToContainer �̃��b�p�[���\�b�h�B
'*          Not implemented. Reserved for future use.
'*
'* @param ContainerURL
'* @param ActiveConnection  �ȗ��\�B
'* @param Mode  �ȗ��\�B
'* @param CreateOptions  �ȗ��\�B
'* @param Options  �ȗ��\�B
'* @param UserName  �ȗ��\�B
'* @param Password  �ȗ��\�B
'*
'******************************************************************************
Sub SaveToContainer(ContainerURL As String, Optional ActiveConnection As Object, _
                                Optional Mode As ConnectModeEnum, _
                                Optional CreateOptions As RecordCreateOptionsEnum = adFailIfNotExists, _
                                Optional Options As RecordOpenOptionsEnum, _
                                Optional UserName As String, _
                                Optional Password As String)
    Call mOrigin.SaveToContainer(ContainerURL, ActiveConnection, Mode, CreateOptions, Options, UserName, Password)
End Sub

'******************************************************************************
'* [�T  �v] SaveToObject ���\�b�h
'* [��  ��] SaveToObject �̃��b�p�[���\�b�h�B
'*           Binds to and saves data into the specified object.
'*
'* @param Source
'* @param InterfaceName
'*
'******************************************************************************
Sub SaveToObject(Source As IUnknown, InterfaceName As String)
    Call mOrigin.SaveToObject(Source, InterfaceName)
End Sub