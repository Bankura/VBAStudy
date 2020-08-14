VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ADODB.Properties�N���X�̃��b�v�E�g���N���X
'* [��  ��] ADODB.Properties�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          ADODB.Properties�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' ADODB.Properties�N���X�ɂ��āB
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Contains all the Property objects for a specific instance of an object.
'
' �I�u�W�F�N�g�̓���̃C���X�^���X�̂��ׂĂ�Property�I�u�W�F�N�g���܂܂��B
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As ADODB.Properties
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
    Set mKeys = New Collection
    Dim myProp As Object
    For Each myProp In mOrigin
        Dim wrap As ADODBExProperty: Set wrap = New ADODBExProperty
        Set wrap.Origin = myProp
        mKeys.Add wrap, myProp.Name
    Next myProp
End Property

'*-----------------------------------------------------------------------------
'* Count �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E�R���N�V�������̃I�u�W�F�N�g���������B
'* �ECount �v���p�e�B�ͤ����̃R���N�V�������̃I�u�W�F�N�g�̐��𒲂ׂ邽�߂�
'*   �g���
'* �E�R���N�V�����̃����o�[�� 0 ���珇�ɔԍ������蓖�Ă��邽�߁A���[�v���g��
'*   �ꍇ�͏�� 0 ����n�߂āA Count �v���p�e�B��� 1 �������l�ŏI��点��K�v
'*   ������B
'*   Microsoft Visual Basic �� Count �v���p�e�B���`�F�b�N�����ɃR���N�V������
'*   �����o�[�����[�v��������ɂ́A For Each...Next �R�}���h���g���B
'* �ECount �� 0 �̏ꍇ�A�R���N�V�����ɂ̓I�u�W�F�N�g���܂܂�Ă��Ȃ����Ƃ��Ӗ�
'*   ����B
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'*-----------------------------------------------------------------------------
'* Item �v���p�e�B�i�ǂݎ���p:���胁���o�[�j
'*
'* [��  ��]
'* �E�R���N�V�����̓���̃����o�[�����̖��O�܂��͏����Ŏ����B
'* �EIndex�ɂ́A�R���N�V�������̃I�u�W�F�N�g�̖��O�܂��͏����ɕ]�������
'*   �o���A���g�^ ( Variant ) �̎����w��B
'* �EItem �v���p�e�B�́A�R���N�V�������̓���̃I�u�W�F�N�g��Ԃ����߂Ɏg���B
'*   �R���N�V�������� Item �� Index �����ɑΉ�����I�u�W�F�N�g���������Ȃ�
'*   �ꍇ�́A�G���[����������B
'*   �܂��A�R���N�V�����̒��ɂ͖��O�t���I�u�W�F�N�g���T�|�[�g���Ă��Ȃ����̂�
'*   ����B���̂悤�ȃR���N�V�����ł́A�����Q�Ƃ��g���K�v������B
'* �EItem �v���p�e�B�͂��ׂẴR���N�V�����̊���v���p�e�B�Ȃ̂Ť���̂������
'*   �\���`���ł��������ʂ�������
'*     collection.Item (Index)
'*     collection (Index)
'*-----------------------------------------------------------------------------
Property Get Item(Index) As ADODBExProperty
Attribute Item.VB_UserMemId = 0
    Dim wrap As ADODBExProperty: Set wrap = New ADODBExProperty
    Set wrap.Origin = mOrigin.Item(Index)
    Set Item = wrap
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("ADODB.Properties")
    'Set mKeys = New Collection
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
'�Ȃ�

'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [�T  �v] NewEnum ���\�b�h
'* [��  ��] For Each �\���ŃC�e���[�^�[��Ԃ����߂̃��\�b�h�B
'*
'* @return �C�e���[�^�[
'******************************************************************************
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mKeys.[_NewEnum]
End Function