VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemRefreshableItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WbemScripting.SWbemRefreshableItem�N���X�̃��b�v�E�g���N���X
'* [��  ��] WbemScripting.SWbemRefreshableItem�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          WbemScripting.SWbemRefreshableItem�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' WbemScripting.SWbemRefreshableItem�N���X�ɂ��āB
' <https://msdn.microsoft.com/ja-jp/windows/aa393831(v=vs.80)>
' A single item from a Refresher
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemRefreshableItem
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
'* Index �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe index of this item in the parent refresher.
'*-----------------------------------------------------------------------------
Property Get Index() As Long
    Index = mOrigin.Index
End Property

'*-----------------------------------------------------------------------------
'* IsSet �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EWhether this item represents a single object or an object set.
'*-----------------------------------------------------------------------------
Property Get IsSet() As Boolean
    IsSet = mOrigin.IsSet
End Property
    
'*-----------------------------------------------------------------------------
'* Object �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe Object.
'*-----------------------------------------------------------------------------
Property Get Object() As WMIExSWbemObjectEx
    Dim wrap As WMIExSWbemObjectEx: Set wrap = New WMIExSWbemObjectEx
    Set wrap.Origin = mOrigin.Object
    Set Object = wrap
End Property
    
'*-----------------------------------------------------------------------------
'* ObjectSet �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe object set.
'*-----------------------------------------------------------------------------
Property Get ObjectSet() As WMIExSWbemObjectSet
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.ObjectSet
    Set ObjectSet = wrap
End Property
    
'*-----------------------------------------------------------------------------
'* Refresher �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe parent refresher.
'*-----------------------------------------------------------------------------
Property Get Refresher() As WMIExSWbemRefresher
    Dim wrap As WMIExSWbemRefresher: Set wrap = New WMIExSWbemRefresher
    Set wrap.Origin = mOrigin.Refresher
    Set Refresher = wrap
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemRefreshableItem")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] Remove ���\�b�h
'* [��  ��] Remove �̃��b�p�[���\�b�h�B
'*          Remove this item from the parent refresher.
'*
'* @param iFlags xx
'*
'******************************************************************************
Sub Remove(Optional iFlags As Long)
    Call mOrigin.Remove(iFlags)
End Sub

'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------


