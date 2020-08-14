VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ADODB.Fields�N���X�̃��b�v�E�g���N���X
'* [��  ��] ADODB.Fields�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          ADODB.FielFieldsds�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' ADODB.Fields�N���X�ɂ��āB
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Contains all the Field objects of a Recordset object.
'
' Recordset�I�u�W�F�N�g�̂��ׂĂ�Field�I�u�W�F�N�g���܂܂��B
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As ADODB.Fields
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
    Dim myField As Object
    For Each myField In mOrigin
        Dim wrap As ADODBExField: Set wrap = New ADODBExField
        Set wrap.Origin = myField
        mKeys.Add wrap, myField.Name
    Next myField
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
Property Get Item(Index) As ADODBExField
Attribute Item.VB_UserMemId = 0
    Dim wrap As ADODBExField: Set wrap = New ADODBExField
    Set wrap.Origin = mOrigin.Item(Index)
    Set Item = wrap
End Property


'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Public Sub Class_Initialize()
    'Set mOrigin = CreateObject("ADODB.Fields")
    'Set mKeys = New Collection
End Sub
Public Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] Append���\�b�h
'* [��  ��] Append�̃��b�p�[���\�b�h�B
'*          �R���N�V�����ɃI�u�W�F�N�g��ǉ�����B
'*          �R���N�V�����ɒǉ�����O�ɁA�V���� Field �I�u�W�F�N�g���쐬�\�B
'*
'* @param Name        �V����Field�I�u�W�F�N�g�̖��O�BFields�Ɋ܂܂��
'*                    ���̃I�u�W�F�N�g�Ƃ͈قȂ閼�O�ɂ���K�v������B
'* @param Type_       �V�����t�B�[���h�̃f�[�^�^�B����l�� adEmpty �B
'*                    adIDispatch �A adIUnknown �A adVariant �̊e�f�[�^�^��
'*                    ADO �ł̓T�|�[�g����Ă��Ȃ��̂ŁA Recordset �ɐV����
'*                    �t�B�[���h��ǉ�����Ƃ��ɁA�����̃f�[�^�^���g�p����
'*                    ���Ƃ͂ł��Ȃ��B
'* @param DefinedSize �ȗ��\�B�V�����t�B�[���h�̒�`���ꂽ�T�C�Y�B
'*                    ���̃p�����[�^�[�̊���l�́AType �ɂ���Č��܂�B
'*                    DefinedSize �� 255 �o�C�g���傫���t�B�[���h�́A
'*                    �ϒ���Ƃ��Ĉ����� (�����DefinedSize�͎w�肳��Ȃ�)�B
'* @param Attrib      �ȗ��\�B�V�����t�B�[���h�̑������w��B
'*                    ����l�� adFldDefault �B�l���w�肵�Ȃ��ƁAType �Ɋ�Â�
'*                    �������ݒ肳���
'* @param FieldValue  �ȗ��\�B�V�����t�B�[���h�̒l�B
'*                    �l���w�肵�Ȃ��ƁA�t�B�[���h�� Null�l�Œǉ������B
'*
'******************************************************************************
Sub Append(Name As String, _
           Type_ As DataTypeEnum, _
           Optional DefinedSize As Long, _
           Optional Attrib As FieldAttributeEnum = adFldUnspecified, _
           Optional FieldValue)
           
    If IsMissing(FieldValue) Then
        Call mOrigin.Append(Name, Type_, DefinedSize, Attrib)
    Else
        Call mOrigin.Append(Name, Type_, DefinedSize, Attrib, FieldValue)
    End If
    
    mKeys.Add mOrigin.Item(Name), Name
    
End Sub

'******************************************************************************
'* [�T  �v] CancelUpdate���\�b�h
'* [��  ��] CancelUpdate�̃��b�p�[���\�b�h�B
'*          Update ���\�b�h���Ăяo���O�ɍs�����ARecordset �I�u�W�F�N�g��
'*          �J�����g�s��V�K�s�ɑ΂���ύX�A�܂��� Record �I�u�W�F�N�g��
'*          Fields �R���N�V�����ɑ΂���ύX���A���ׂăL�����Z������B
'*
'******************************************************************************
Sub CancelUpdate()
    Call mOrigin.CancelUpdate
End Sub

'******************************************************************************
'* [�T  �v] Delete���\�b�h
'* [��  ��] Delete�̃��b�p�[���\�b�h�B
'*          Fields �R���N�V��������I�u�W�F�N�g���폜����B
'*
'* @param Index �폜����Field�I�u�W�F�N�g���w�肷��l�B
'*              ���̃p�����[�^�[�ɂ́A Field �I�u�W�F�N�g�̖��O�܂��� Field
'*              �I�u�W�F�N�g���̂̃C���f�b�N�X���g�p�ł���B
'*
'******************************************************************************
Sub Delete(Index)
    Dim keyName: keyName = mOrigin.Item(Index).Name
    Call mOrigin.Delete(Index)

    Dim i As Long
    For i = 1 To mKeys.Count
        If mKeys(i) = keyName Then
            mKeys.Remove i
            Exit Sub
        End If
    Next
End Sub

'******************************************************************************
'* [�T  �v] Refresh���\�b�h
'* [��  ��] Refresh�̃��b�p�[���\�b�h�B
'*          �R���N�V�����̃I�u�W�F�N�g���X�V����B
'*          Fields �R���N�V�����ɑ΂��� Refresh ���\�b�h���g�p���Ă��A
'*          �ڂɌ�������ʂ͂Ȃ��B��ɂȂ��Ă���f�[�^�x�[�X�\������
'*          �ύX���擾����ɂ́ARequery ���\�b�h���g�p���邩�A�܂��� Recordset
'*          �I�u�W�F�N�g���u�b�N�}�[�N���T�|�[�g���Ȃ��ꍇ�� MoveFirst ���\�b�h
'*          ���g�p����K�v������B
'*
'******************************************************************************
Sub Refresh()
    Call mOrigin.Refresh
End Sub

'******************************************************************************
'* [�T  �v] Resync���\�b�h
'* [��  ��] Resync�̃��b�p�[���\�b�h�B
'*          ���݂� Recordset �I�u�W�F�N�g�A�܂��� Record �I�u�W�F�N�g�� Fields
'*          �R���N�V�����̃f�[�^���ɂȂ�f�[�^�x�[�X�̃f�[�^�ōX�V����B
'*          Resync ���\�b�h�́ARecord �I�u�W�F�N�g�� Fields �R���N�V�����̒l���A
'*          ��ɂȂ�f�[�^�\�[�X�ƍē���������ꍇ�Ɏg�p����B
'*          Count �v���p�e�B�́A���̃��\�b�h�ɂ��e�����󂯂Ȃ��B
'*          ResyncValues �� adResyncAllValues (����l) �ɐݒ肷��ƁA�R���N�V��
'*          ���Ɋ܂܂�� Field �I�u�W�F�N�g�̃v���p�e�B UnderlyingValue�AValue�A
'*          ����� OriginalValue ������������B
'*          ResyncValues �� adResyncUnderlyingValues �ɐݒ肷��ƁA
'*          UnderlyingValue �v���p�e�B�����������������B
'*          �Ăяo�����̊e Field �I�u�W�F�N�g�� Status �v���p�e�B�̒l���A
'*          Resync �̓���ɉe����^����B
'*          Status �̒l�� adFieldPendingUnknown �܂��� adFieldPendingInsert
'*          �ł��� Field �I�u�W�F�N�g�ɑ΂��ẮA Resync �͉����s��Ȃ��B
'*          Status �̒l�� adFieldPendingChange �܂��� adFieldPendingDelete
'*          �ł���ꍇ�́A Resync �̓f�[�^�\�[�X�ɂ܂����݂��Ă���t�B�[���h��
'*          �f�[�^�l�𓯊�������B
'*          Resync���Ăяo���ꂽ�Ƃ��ɃG���[���������Ȃ�����AResync��Field�I�u
'*          �W�F�N�g��Status�l��ύX���Ȃ��B
'*          ���Ƃ��΁A�t�B�[���h�����݂��Ȃ��Ȃ����ꍇ�A�v���o�C�_�[��
'*          adFieldDoesNotExist�Ȃǂ�Field�I�u�W�F�N�g�̓K�؂�Status�l��Ԃ��B
'*          �Ԃ����Status�l�́AStatus�v���p�e�B�̒l���Ř_���I�ɑg�ݍ��킹��
'*          ���Ƃ��ł���B
'*
'* @param ResyncValues �ȗ��\�B��ɂȂ�l���㏑�����邩�ǂ������w��B
'*
'******************************************************************************
Sub Resync(Optional ResyncValues As ResyncEnum = adResyncAllValues)
    Call mOrigin.Resync(ResyncValues)
End Sub

'******************************************************************************
'* [�T  �v] Update���\�b�h
'* [��  ��] Update�̃��b�p�[���\�b�h�B
'*          Record �I�u�W�F�N�g�� Fields �R���N�V�����ɉ������ύX��ۑ�����B
'*
'******************************************************************************
Sub Update()
    Call mOrigin.Update
End Sub

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