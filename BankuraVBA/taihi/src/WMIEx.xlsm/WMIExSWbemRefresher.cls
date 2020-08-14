VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemRefresher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WbemScripting.SWbemRefresher�N���X�̃��b�v�E�g���N���X
'* [��  ��] WbemScripting.SWbemRefresher�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          WbemScripting.SWbemRefresher�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' WbemScripting.SWbemRefresher�N���X�ɂ��āB
' <https://msdn.microsoft.com/ja-jp/windows/aa393838(v=vs.80)>
' Refresher
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemRefresher
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
'* AutoReconnect �v���p�e�B
'*
'* [��  ��]
'* �EWhether to attempt auto-reconnection to a remote provider.
'*-----------------------------------------------------------------------------
Property Get AutoReconnect() As Boolean
    AutoReconnect = mOrigin.AutoReconnect
End Property
Property Let AutoReconnect(arg As Boolean)
    mOrigin.AutoReconnect = arg
End Property

'*-----------------------------------------------------------------------------
'* Count �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe number of items in this refresher.
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("WbemScripting.SWbemRefresher")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] Add ���\�b�h
'* [��  ��] Add �̃��b�p�[���\�b�h�B
'*         Add a refreshable instance to this refresher.
'*
'* @param objWbemServices xx
'* @param bsInstancePath xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemRefreshableItem
'*
'******************************************************************************
Function Add(objWbemServices As WMIExSWbemServicesEx, bsInstancePath As String, Optional iFlags As Long, Optional objWbemNamedValueSet As Object) As WMIExSWbemRefreshableItem
    Dim wrap As WMIExSWbemRefreshableItem: Set wrap = New WMIExSWbemRefreshableItem
    Set wrap.Origin = mOrigin.Add(objWbemServices.Origin, bsInstancePath, iFlags, objWbemNamedValueSet)
    Set Add = wrap
End Function

'******************************************************************************
'* [�T  �v] AddEnum ���\�b�h
'* [��  ��] AddEnum �̃��b�p�[���\�b�h�B
'*          Add a refreshable enumerator to this refresher.
'*
'* @param objWbemServices xx
'* @param bsInstancePath xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemRefreshableItem
'*
'******************************************************************************
Function AddEnum(objWbemServices As WMIExSWbemServicesEx, bsClassName As String, Optional iFlags As Long, Optional objWbemNamedValueSet As Object) As WMIExSWbemRefreshableItem
    Dim wrap As WMIExSWbemRefreshableItem: Set wrap = New WMIExSWbemRefreshableItem
    Set wrap.Origin = mOrigin.AddEnum(objWbemServices.Origin, bsClassName, iFlags, objWbemNamedValueSet)
    Set AddEnum = wrap
End Function

'******************************************************************************
'* [�T  �v] DeleteAll ���\�b�h
'* [��  ��] DeleteAll �̃��b�p�[���\�b�h�B
'*          Delete all items in this collection.
'*
'******************************************************************************
Sub DeleteAll()
    Call mOrigin.DeleteAll
End Sub
    
'******************************************************************************
'* [�T  �v] Item ���\�b�h�i���胁���o�[�j
'* [��  ��] Item �̃��b�p�[���\�b�h�B
'*          Get an item from this refresher.
'*
'* @param Name xx
'* @param iFlags xx
'* @return WMIExSWbemProperty
'*
'******************************************************************************
Function Item(iIndex As Long) As WMIExSWbemRefreshableItem
Attribute Item.VB_UserMemId = 0
    Dim wrap As WMIExSWbemRefreshableItem: Set wrap = New WMIExSWbemRefreshableItem
    Set wrap.Origin = mOrigin.Item(iIndex)
    Set Item = wrap
End Function

'******************************************************************************
'* [�T  �v] Remove ���\�b�h
'* [��  ��] Remove �̃��b�p�[���\�b�h�B
'*          Refresh all items in this collection.
'*
'* @param strName xx
'* @param iFlags xx
'*
'******************************************************************************
Sub Refresh(Optional iFlags As Long)
    Call mOrigin.Refresh(iFlags)
End Sub
    
'******************************************************************************
'* [�T  �v] Remove ���\�b�h
'* [��  ��] Remove �̃��b�p�[���\�b�h�B
'*          Remove an item from this refresher.
'*
'* @param iIndex xx
'* @param iFlags xx
'*
'******************************************************************************
Sub Remove(iIndex As Long, Optional iFlags As Long)
    Call mOrigin.Remove(iIndex, iFlags)
End Sub

'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------