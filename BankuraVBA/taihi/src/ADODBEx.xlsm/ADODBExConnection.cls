VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ADODB.Connection�N���X�̃��b�v�E�g���N���X
'* [��  ��] ADODB.Connection�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          ADODB.Connection�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' ADODB.Connection�N���X�ɂ��āB
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Represents a unique session with a data source.
' In the case of a client/server database system, it may be equivalent to
' an actual network connection to the server.
' Depending on the functionality supported by the provider, some collections,
' methods, or properties of a Connection object may not be available.
'
' Connection�I�u�W�F�N�g�́A�f�[�^�\�[�X�Ƃ̈�ӂ̃Z�b�V������\���B
' �N���C�A���g/�T�[�o�[�f�[�^�x�[�X�V�X�e���ł́A�T�[�o�[�ւ̎��ۂ�
' �l�b�g���[�N�ڑ��Ɠ����ڑ���\���ꍇ������
' �v���o�C�_�[�̃T�|�[�g�@�\�ɂ���ẮA�R���N�V�����I�u�W�F�N�g�̂�������
' �R���N�V�����A���\�b�h�A�v���p�e�B���g�p�ł��Ȃ��\��������
'
' <�C�x���g���>
' Event BeginTransComplete(TransactionLevel As Long, pError As Error, adStatus As EventStatusEnum, pConnection As Connection)
' Event CommitTransComplete(pError As Error, adStatus As EventStatusEnum, pConnection As Connection)
' Event ConnectComplete(pError As Error, adStatus As EventStatusEnum, pConnection As Connection)
' Event Disconnect(adStatus As EventStatusEnum, pConnection As Connection)
' Event ExecuteComplete(RecordsAffected As Long, pError As Error, adStatus As EventStatusEnum, pCommand As Command, pRecordset As Recordset, pConnection As Connection)
' Event InfoMessage(pError As Error, adStatus As EventStatusEnum, pConnection As Connection)
' Event RollbackTransComplete(pError As Error, adStatus As EventStatusEnum, pConnection As Connection)
' Event WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As EventStatusEnum, pConnection As Connection)
' Event WillExecute(Source As String, CursorType As CursorTypeEnum, LockType As LockTypeEnum, Options As Long, adStatus As EventStatusEnum, pCommand As Command, pRecordset As Recordset, pConnection As Connection)
'
'------------------------------------------------------------------------------


'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As ADODB.Connection
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
    If Not mOrigin Is Nothing Then
        If mOrigin.State <> adStateClosed Then
            Call mOrigin.Close
        End If
    End If
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* Attributes �v���p�e�B
'*
'* [��  ��]
'* �E�I�u�W�F�N�g�� 1 �܂��͕����̑����������B
'* �EAttributes �v���p�e�B�͒l�̎擾����ѐݒ肪�\�ŁA���̒l��1�܂��͕�����
'*   XactAttributeEnum �l�̍��v�ɂȂ�B����l�� 0 �B
'* �E�����̑�����ݒ肷��ꍇ�́A�Y������萔�̍��v���g�p�ł���B
'*   �v���p�e�B�̒l��݊����̂Ȃ��萔���܂ލ��v�ɐݒ肷��Ƥ�G���[����������
'* �E���̃v���p�e�B�ͤ�N���C�A���g����Connection�I�u�W�F�N�g�ł͎g�p�ł��Ȃ��
'*-----------------------------------------------------------------------------
Property Get Attributes() As Long
    Attributes = mOrigin.Attributes
End Property
Property Let Attributes(arg As Long)
    mOrigin.Attributes = arg
End Property

'*-----------------------------------------------------------------------------
'* CommandTimeout �v���p�e�B
'*
'* [��  ��]
'* �E�R�}���h�����s�����܂ł̑ҋ@���Ԃ�b�P�ʂŎ����l��ݒ�܂��͎擾����B
'*   ����l�� 30 �B
'* �E�l�b�g���[�N �g���t�B�b�N��T�[�o�[�̉ߕ��ׂɂ����s���x��Ă��� Execute
'*   ���\�b�h�̌Ăяo�������������Ƃ��ł���悤�ɂ���ɂ́A Connection �I�u�W
'*   �F�N�g�܂��� Command �I�u�W�F�N�g�� CommandTimeout �v���p�e�B���g�p����B
'*   �R�}���h�̎��s����������O�� CommandTimeout �v���p�e�B�Őݒ肳�ꂽ���Ԃ�
'*   �o�߂���ƁA�G���[���������ăR�}���h�����������B
'*   �v���p�e�B�� 0 �ɐݒ肷��ƁA�R�}���h���s����������܂Ŗ����ɑҋ@����B
'*   �R�[�h���������ރv���o�C�_�[�ƃf�[�^ �\�[�X�� CommandTimeout �@�\���T�|�[
'*   �g���Ă��邱�Ƃ��m�F���邱�ƁB
'* �EConnection �I�u�W�F�N�g�� CommandTimeout �ݒ�́A���� Connection ���
'*   Command �I�u�W�F�N�g�� CommandTimeout �ݒ�ɉe�����Ȃ��B
'*   �܂�A Command �I�u�W�F�N�g�� CommandTimeout �v���p�e�B�́A Connection
'*   �I�u�W�F�N�g�� CommandTimeout �̒l���p�����Ȃ��B
'* �EConnection �I�u�W�F�N�g�ł́A CommandTimeout �v���p�e�B�� Connection ���J
'*   ���ꂽ����A�ǂݎ��/�������݂��\�B
'*-----------------------------------------------------------------------------
Property Get CommandTimeout() As Long
    CommandTimeout = mOrigin.CommandTimeout
End Property
Property Let CommandTimeout(arg As Long)
    mOrigin.CommandTimeout = arg
End Property

'*-----------------------------------------------------------------------------
'* ConnectionString �v���p�e�B�i���胁���o�[�j
'*
'* [��  ��]
'* �E�f�[�^�\�[�X�Ƃ̐ڑ����m�����邽�߂Ɏg�p�������������B
'* �E��A�̈��� = value�X�e�[�g�����g���܂ޏڍׂȐڑ���������Z�~�R�����ŋ�؂�
'*   �ēn���āA�f�[�^�\�[�X���w�肷��B
'* �EADO �ł́A ConnectionString �v���p�e�B�ɑ΂��� 5 ��ނ̈������T�|�[�g����
'*   ��B���̑��̈����� ADO �ŏ������ꂸ�ɒ��ڃv���o�C�_�[�ɓn�����B
'*   ADO �ŃT�|�[�g�������������Ɏ����B
'*
'*   Provider =       �ڑ��Ɏg�p����v���o�C�_�[�̖��O���w�肷��B
'*   File Name=       �ݒ�ς݂̐ڑ������i�[�����v���o�C�_�[�ŗL�̃t�@�C��
'*                    (���Ƃ��΁A�����I�ȃf�[�^ �\�[�X �I�u�W�F�N�g) �̖��O��
'*                    �w�肷��B
'*   Remote Provider= �N���C�A���g���̐ڑ����J���Ƃ��Ɏg�p����v���o�C�_�[��
'*                    ���O���w�肷��B (�����[�g�f�[�^�T�[�r�X�̂�)
'*   Remote Server=   �N���C�A���g���̐ڑ����J���Ƃ��Ɏg�p����T�[�o�[�̃p�X��
'*                    ���w�肷�� (�����[�g �f�[�^ �T�[�r�X�̂�)�B
'*   URL =            �ڑ���������A�t�@�C����f�B���N�g���Ȃǂ̃��\�[�X������
'*                    ������ URL �Ƃ��Ďw�肷��B
'*
'* �EConnectionString �v���p�e�B��ݒ肵�� Connection �I�u�W�F�N�g���J������A
'*   ADO �ɂ���Ē�`���ꂽ���������v���o�C�_�[�̑Ή�����������Ƀ}�b�v�����
'*   �ȂǁA�v���o�C�_�[�ɂ���ăv���p�e�B�̓��e���ύX�����ꍇ������B
'* �EConnectionString �v���p�e�B�́AOpen ���\�b�h�� ConnectionString ������
'*   �g�p���ꂽ�l�������I�Ɍp������B
'*   ���������āAOpen ���\�b�h�̌Ăяo�����Ɍ��݂� ConnectionString �v���p�e�B
'*   ���㏑���ł���B
'* �EFile Name �����ɂ��֘A�̂���v���o�C�_�[���Ăяo����邽�߁AProvider
'*   ������ File Name �����̗�����n�����Ƃ͂ł��Ȃ��B
'* �EConnectionString �v���p�e�B�́A�ڑ������Ă���Ƃ��͓ǂݎ��/��������
'*   �\�ŁA�J���Ă���Ƃ��͓ǂݎ���p�ɂȂ�B
'* �EConnectionString �v���p�e�B�ɂ����ďd�����Ă�������͖��������B
'*   �����̍Ō�̃C���X�^���X���g�p�����B
'* �E�����[�g�f�[�^�T�[�r�X�̎g�p�󋵃N���C�A���g����Connection�I�u�W�F�N�g��
'*   �g�p����ꍇ�A ConnectionString�v���p�e�B�Ɋ܂߂邱�Ƃ��ł���̂́A�����[
'*   �g�v���o�C�_�[����у����[�g�T�[�o�[�̃p�����[�^�[�����ł���B
'*
'*-----------------------------------------------------------------------------
Property Get ConnectionString() As String
Attribute ConnectionString.VB_UserMemId = 0
    ConnectionString = mOrigin.ConnectionString
End Property
Property Let ConnectionString(arg As String)
    mOrigin.ConnectionString = arg
End Property

'*-----------------------------------------------------------------------------
'* ConnectionTimeout �v���p�e�B
'*
'* [��  ��]
'* �E�ڑ����J���܂ł̑ҋ@���Ԃ�b�P�ʂŎ����B����l�� 15 �B
'* �E�l�b�g���[�N �g���t�B�b�N��T�[�o�[�̉ߓx�̎g�p�������Őڑ����J�����݂�
'*   ���~����K�v������ꍇ�́AConnection �I�u�W�F�N�g�� ConnectionTimeout
'*   �v���p�e�B���g�p����B
'*   �ڑ����J�����O�� ConnectionTimeout �v���p�e�B�Őݒ肵�����Ԃ��o�߂���
'*   �ꍇ�́A�G���[���������Đڑ��̎��݂����������B
'*   ���̃v���p�e�B�� 0 �ɐݒ肵���ꍇ�́AADO �͐ڑ����J�����܂Ŗ����ɑҋ@
'*   ����B�R�[�h���������ރv���o�C�_�[���A ConnectionTimeout �@�\���T�|�[�g
'*   ���Ă��邱�Ƃ��m�F���邱�ƁB
'* �EConnectionTimeout �v���p�e�B�ͤ�ڑ������Ă���Ƃ��͓ǂݎ�� / ��������
'*   �\�Ť�J���Ă���Ƃ��͓ǂݎ���p�ɂȂ�
'*-----------------------------------------------------------------------------
Property Get ConnectionTimeout() As Long
    ConnectionTimeout = mOrigin.ConnectionTimeout
End Property
Property Let ConnectionTimeout(arg As Long)
    mOrigin.ConnectionTimeout = arg
End Property

'*-----------------------------------------------------------------------------
'* CursorLocation �v���p�e�B
'*
'* [��  ��]
'* �E�J�[�\���T�[�r�X�̏ꏊ�������B
'* �E���̃v���p�e�B�ŁA�v���o�C�_�[�ɃA�N�Z�X�\�Ȃ��܂��܂ȃJ�[�\�� ���C�u��
'*   ���̒�����J�[�\�� �T�[�r�X��I������B
'*   �ʏ�́A�N���C�A���g���J�[�\�� ���C�u�����A�܂��̓T�[�o�[���J�[�\�� ���C�u
'*   ��������I������B
'* �E���̃v���p�e�B�ݒ�́A�v���p�e�B���ݒ肳�ꂽ��Ɋm�����ꂽ�ڑ��݂̂ɍ�p
'*   ����BCursorLocation �v���p�e�B��ύX���Ă������̐ڑ��ɂ͉e�����Ȃ��B
'* �EExecute ���\�b�h���Ԃ��J�[�\���́A���̐ݒ���p������B
'*   Recordset �I�u�W�F�N�g�́A�֘A�t����ꂽ�ڑ����炱�̐ݒ�������I�Ɍp������B
'* �E���̃v���p�e�B�́AConnection �܂��͕��Ă��� Recordset ��ł͓ǂݎ��/����
'*   ���݉\�����A�J���Ă��� Recordset ��ł͓ǂݎ���p�ɂȂ�B
'* �E�����[�g�f�[�^�T�[�r�X�̎g�p�󋵃N���C�A���g����Recordset�I�u�W�F�N�g�܂�
'*   ��Connection�I�u�W�F�N�g�Ŏg�p����ꍇ��J�[�\���ʒu�v���p�e�B��adUseClient
'*   �ɂ̂ݐݒ�ł���
'*-----------------------------------------------------------------------------
Property Get CursorLocation() As CursorLocationEnum
    CursorLocation = mOrigin.CursorLocation
End Property
Property Let CursorLocation(arg As CursorLocationEnum)
    mOrigin.CursorLocation = arg
End Property

'*-----------------------------------------------------------------------------
'* DefaultDatabase �v���p�e�B
'*
'* [��  ��]
'* �EConnection �I�u�W�F�N�g�̊���̃f�[�^�x�[�X�������B
'* �E����̃f�[�^�x�[�X������ꍇ�A���̃f�[�^�x�[�X�̃I�u�W�F�N�g�ɃA�N�Z�X����
'*   SQL ���̍\�����s�K�؂ȏꍇ������B
'*   DefaultDatabase �v���p�e�B�Ŏw�肳�ꂽ�f�[�^�x�[�X�ȊO�̃f�[�^�x�[�X�̃I�u
'*   �W�F�N�g�ɃA�N�Z�X����ɂ́A�I�u�W�F�N�g����ړI�̃f�[�^�x�[�X���ŏC������
'*   �K�v������B
'*   �ڑ����ɁA�v���o�C�_�[�� DefaultDatabase �v���p�e�B�Ɋ���̃f�[�^�x�[�X��
'*   ����������ށB
'*   �v���o�C�_�[�̒��ɂ� 1 �̐ڑ��� 1 �̃f�[�^�x�[�X���������Ȃ����̂���
'*   ��A���̏ꍇ�� DefaultDatabase �v���p�e�B��ύX�ł��Ȃ��B
'* �E�f�[�^ �\�[�X�ƃv���o�C�_�[�ɂ���Ăͤ���̋@�\���T�|�[�g������G���[�܂���
'*   ��̕������Ԃ��ꍇ������
'* �E���̃v���p�e�B�ͤ�N���C�A���g����Connection�I�u�W�F�N�g�ł͎g�p�ł��Ȃ��
'*-----------------------------------------------------------------------------
Property Get DefaultDatabase() As String
    DefaultDatabase = mOrigin.DefaultDatabase
End Property
Property Let DefaultDatabase(arg As String)
    mOrigin.DefaultDatabase = arg
End Property

'*-----------------------------------------------------------------------------
'* Errors �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EConnection�ɐݒ肳��Ă���Errors�I�u�W�F�N�g���擾����B
'*-----------------------------------------------------------------------------
Property Get Errors() As ADODBExErrors
    Dim wrap As ADODBExErrors: Set wrap = New ADODBExErrors
    Set wrap.Origin = mOrigin.Errors
    Set Errors = wrap
End Property

'*-----------------------------------------------------------------------------
'* IsolationLevel �v���p�e�B
'*
'* [��  ��]
'* �EConnection �I�u�W�F�N�g�̕������x���������B
'*   ����l�� adXactChaos �B
'* �EIsolationLevel �v���p�e�B�ł́A Connection �I�u�W�F�N�g�̕������x����ݒ�
'*   ����B
'*   �ݒ�l�́A���� BeginTrans ���\�b�h���Ăяo���܂ŗL���ɂȂ�Ȃ��B
'*   �v�������������x�����g�p�ł��Ȃ��ꍇ�A�v���o�C�_�[�͂��̎��ɍ����������x��
'*   ��Ԃ��B
'* �EIsolationLevel �v���p�e�B�͓ǂݎ�� / �������݉\�
'* �E�����[�g�f�[�^�T�[�r�X�̎g�p�󋵃N���C�A���g���� Connection �I�u�W�F�N�g��
'*   �g�p����ꍇ�A IsolationLevel�v���p�e�B��adXactUnspecified�ɂ̂ݐݒ�ł���B
'* �E���[�U�[�́A�N���C�A���g���̃L���b�V����̐ڑ�����Ă��Ȃ� Recordset �I�u
'*   �W�F�N�g�ō�Ƃ��邽�߁A�}���`���[�U�[�̏ꍇ�͖��ɂȂ邱�Ƃ�����B
'*   ���Ƃ��΁A2 �l�̃��[�U�[���������R�[�h���X�V���悤�Ƃ����ہA�����[�g �f�[�^
'*   �T�[�r�X�͒P���ɐ�ɑ�����s�������[�U�[�̍X�V���󂯕t����B
'*   2 �Ԗڂ̃��[�U�[�̍X�V�v���̓G���[�ɂȂ��Ď��s����B
'*-----------------------------------------------------------------------------
Property Get IsolationLevel() As IsolationLevelEnum
    IsolationLevel = mOrigin.IsolationLevel
End Property
Property Let IsolationLevel(arg As IsolationLevelEnum)
    mOrigin.IsolationLevel = arg
End Property

'*-----------------------------------------------------------------------------
'* Mode �v���p�e�B
'*
'* [��  ��]
'* �E�g�p�\�ȃf�[�^�ύX�����������B
'*   Connection �̊���l�� adModeUnknown �B
'* �EMode �v���p�e�B�́A�v���o�C�_�[�����݂̐ڑ��Ŏg�p���Ă���A�N�Z�X����ݒ�
'*   �܂��͎擾���邽�߂Ɏg�p����B
'*   Mode �v���p�e�B�́A Connection �I�u�W�F�N�g�����Ă���Ƃ��ɂ̂ݐݒ�ł���B
'* �E���̃v���p�e�B�́A�I�u�W�F�N�g�����Ă���Ƃ��͓ǂݎ��/�������݉\�ŁA
'*   �I�u�W�F�N�g���J���Ă���Ƃ��͓ǂݎ���p�ɂȂ�B
'* �E�����[�g�f�[�^�T�[�r�X�̎g�p�󋵃N���C�A���g���� Connection �I�u�W�F�N�g��
'*   �g�p����ꍇ�A Mode�v���p�e�B��admodeunknown�ɂ̂ݐݒ�ł���B
'*-----------------------------------------------------------------------------
Property Get Mode() As ConnectModeEnum
    Mode = mOrigin.Mode
End Property
Property Let Mode(arg As ConnectModeEnum)
    mOrigin.Mode = arg
End Property

'*-----------------------------------------------------------------------------
'* Properties �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EConnection�ɐݒ肳��Ă���Properties�I�u�W�F�N�g���擾����B
'*-----------------------------------------------------------------------------
Property Get Properties() As ADODBExProperties
    Dim wrap As ADODBExProperties: Set wrap = New ADODBExProperties
    Set wrap.Origin = mOrigin.Properties
    Set Properties = wrap
End Property

'*-----------------------------------------------------------------------------
'* Provider �v���p�e�B
'*
'* [��  ��]
'* �EConnection �I�u�W�F�N�g�̃v���o�C�_�[���������B
'* �E�ڑ��p�̃v���o�C�_�[�̖��O��ݒ�܂��͎擾����ɂ́AProvider �v���p�e�B��
'*   �g�p����B
'*   ���̃v���p�e�B�́AConnectionString �v���p�e�B�܂��� Open ���\�b�h��
'*   ConnectionString �����̓��e�ɂ���Đݒ肷�邱�Ƃ��ł���B
'*   �������AOpen ���\�b�h���Ăяo���Ƃ��ɕ����̉ӏ��Ńv���o�C�_�[���w�肷���
'*   �\�����Ȃ����ʂ�������\��������B
'*   �v���o�C�_�[���w�肵�Ȃ��ꍇ�A���̃v���p�e�B�͊���l�� MSDASQL (Microsoft
'*   OLE DB Provider for ODBC) �ɂȂ�B
'* �EProvider �v���p�e�B�́A�ڑ������Ă���Ƃ��͒l�̐ݒ肨��ю擾���\�ŁA
'*   �ڑ����J���Ă���Ƃ��͒l�̎擾�݂̂��\�B
'*   �ݒ�l�� Connection �I�u�W�F�N�g���J�����AConnection �I�u�W�F�N�g��
'*   Properties �R���N�V�����ɃA�N�Z�X����܂ŗL���ɂȂ�Ȃ��B
'*   �ݒ肪�����ł���ꍇ�́A�G���[����������B
'*-----------------------------------------------------------------------------
Property Get Provider() As String
    Provider = mOrigin.Provider
End Property
Property Let Provider(arg As String)
    mOrigin.Provider = arg
End Property

'*-----------------------------------------------------------------------------
'* State �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E�I�u�W�F�N�g�̏�Ԃ��J���Ă��邩�A���Ă��邩�������B
'* �E�񓯊����\�b�h�����s����ΏۂɂȂ邷�ׂẴI�u�W�F�N�g�ɂ��Ĥ�I�u�W�F�N
'*   �g�̏�Ԃ��ڑ��A���s�A�擾�̂�����ł��邩�������
'* �EObjectStateEnum �̒l�ɂȂ钷�����^ ( Long ) �̒l��Ԃ��B
'*   ����l�� adStateClosed �B
'* �EState �v���p�e�B���g�p���Ĥ����̃I�u�W�F�N�g�̌��݂̏�Ԃ����ł����ׂ�
'*   ���Ƃ��ł���
'* �E�I�u�W�F�N�g�� State �v���p�e�B�́A�l�̑g�ݍ��킹�ɂȂ�ꍇ������B
'*   ���Ƃ��΁A�X�e�[�g�����g�����s���ł���ꍇ�A���̃v���p�e�B�̒l��
'*   adStateOpen �� adStateExecuting �̑g�ݍ��킹�ɂȂ�B
'* �EState �v���p�e�B�͒l�̎擾�̂݉\�
'*-----------------------------------------------------------------------------
Property Get State() As Long
    State = mOrigin.State
End Property

'*-----------------------------------------------------------------------------
'* Version �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EADO �̃o�[�W�����ԍ��������B
'* �EADO �����̃o�[�W�����ԍ���Ԃ��ɂ́AVersion �v���p�e�B���g�p����B
'* �E�v���o�C�_�[�̃o�[�W�����ͤProperties �R���N�V�����̓��I�v���p�e�B�Ƃ���
'*   �g�p�\�
'*-----------------------------------------------------------------------------
Property Get Version() As String
    Version = mOrigin.Version
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("ADODB.Connection")
End Sub
Private Sub Class_Terminate()
    If Not mOrigin Is Nothing Then
        If mOrigin.State <> adStateClosed Then
            Call mOrigin.Close
        End If
    End If
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] BeginTrans���\�b�h
'* [��  ��] BeginTrans�̃��b�p�[���\�b�h�B
'*          �V�����g�����U�N�V�������J�n����B
'*          �������� BeginTrans ���\�b�h���Ăяo���ƁA CommitTrans ���\�b�h
'*          �܂���RollbackTrans ���\�b�h���Ăяo���ăg�����U�N�V�������I������
'*          �܂ŁA�ύX���R�~�b�g����邱�Ƃ͂Ȃ��B
'*          ����q�ɂȂ����g�����U�N�V�������T�|�[�g����v���o�C�_�[�̏ꍇ�A
'*          �J���Ă���g�����U�N�V������ BeginTrans ���\�b�h���Ăяo���ƁA
'*          ����q�ɂȂ����V�����g�����U�N�V�������J�n�����B
'*          �߂�l�́A����q�̃��x���������B
'*          �߂�l "1" �͍ŏ�ʃ��x���̃g�����U�N�V���� (���̃g�����U�N�V����
'*          �̓���q�ɂȂ��Ă��Ȃ��g�����U�N�V����) ���J�������Ƃ������A
'*          "2" �͑� 2 ���x���̃g�����U�N�V���� (�ŏ�ʃ��x���̃g�����U�N�V��
'*          ���̓���q�ɂȂ��Ă���g�����U�N�V����) ���J�������Ƃ����� ("3"
'*          �ȉ������l)�B
'*          CommitTrans ���\�b�h�܂��� RollbackTrans ���\�b�h���Ăяo���ƁA
'*          �Ō�ɊJ�����g�����U�N�V�������������������B
'*          ����ɏ�̃��x���̃g�����U�N�V��������������ɂ́A���݂̃g�����U
'*          �N�V��������邩�A�܂��̓��[���o�b�N����K�v������B
'*
'* @return �g�����U�N�V�����̓���q���x���������l
'******************************************************************************
Function BeginTrans() As Long
    Call mOrigin.BeginTrans
End Function

'******************************************************************************
'* [�T  �v] Cancel���\�b�h
'* [��  ��] Cancel�̃��b�p�[���\�b�h�B
'*          �ۗ����̔񓯊����\�b�h�Ăяo���̎��s���������B
'*          �񓯊����\�b�h�Ăяo�� (�܂�AadAsyncConnect�AadAsyncExecute�A
'*          �܂��� adAsyncFetch �I�v�V�������w�肵�ČĂяo���ꂽ���\�b�h) ��
'*          ���s�𒆎~����B
'*          Connection�̏ꍇ�AExecute�܂���Open���Ō�̔񓯊��Ăяo�������~����
'*          �郁�\�b�h�
'*
'******************************************************************************
Sub Cancel()
    Call mOrigin.Cancel
End Sub

'******************************************************************************
'* [�T  �v] CloseCn���\�b�h
'* [��  ��] Close�̃��b�p�[���\�b�h�B
'*          �J���Ă���I�u�W�F�N�g����т���Ɋ֘A���邷�ׂẴI�u�W�F�N�g��
'*          ����B
'*
'*          �֘A���邷�ׂẴV�X�e�� ���\�[�X���������ꍇ�Ɏg�p����B
'*          �I�u�W�F�N�g����Ă�����������͍폜���ꂸ�A�v���p�e�B�ݒ��ύX
'*          ���Ă�����x�J�����Ƃ��ł���B
'*          �I�u�W�F�N�g�����������犮�S�ɍ폜����ɂ́A�I�u�W�F�N�g�������
'*          �ɁA�I�u�W�F�N�g�ϐ���Nothing�ɐݒ肷��B
'*          Close ���\�b�h���g�p���� Connection �I�u�W�F�N�g�����ƁA����
'*          �ڑ��Ɋ֘A����A�N�e�B�u�� Recordset �I�u�W�F�N�g�����ׂĕ���B
'*          ���� Connection �I�u�W�F�N�g�Ɋ֘A���� Command �I�u�W�F�N�g��
'*          ���̂܂܈ێ�����邪�A Connection �I�u�W�F�N�g�Ƃ̊֘A�͂Ȃ��Ȃ�B
'*          �܂�A ActiveConnection �v���p�e�B�� Nothing �ɐݒ肳���B
'*          �܂��A Command �I�u�W�F�N�g�� Parameters �R���N�V�������̃v���o�C
'*          �_�[��`�̃p�����[�^�[�͂��ׂăN���A�����B
'*          ��� Open ���\�b�h���Ăяo���āA�����f�[�^ �\�[�X�܂��͕ʂ̃f�[�^
'*          �\�[�X�ւ̐ڑ����ēx�m�����邱�Ƃ��ł���B
'*          Connection �I�u�W�F�N�g�����Ă���ԂɁA�f�[�^ �\�[�X�ɑ΂���J��
'*          ���ڑ���K�v�Ƃ��郁�\�b�h���Ăяo���ƁA�G���[����������B
'*          �ڑ���ɊJ���� Recordset �I�u�W�F�N�g�������Ԃ� Connection �I�u
'*          �W�F�N�g�����ƁA���ׂĂ� Recordset �I�u�W�F�N�g�ۗ̕����̕ύX
'*          �����ׂă��[���o�b�N�����B
'*          �g�����U�N�V�����̏������� Close ���\�b�h���Ăяo���Ė����I��
'*          Connection �I�u�W�F�N�g�����ƁA�G���[����������B
'*          �g�����U�N�V�����̏������� Connection �I�u�W�F�N�g���K�p�͈͂��O��
'*          ��ƁA�g�����U�N�V�����͎����I�Ƀ��[���o�b�N�����B
'*
'******************************************************************************
Sub CloseCn()
    Call mOrigin.Close
End Sub

'******************************************************************************
'* [�T  �v] CommitTrans���\�b�h
'* [��  ��] CommitTrans�̃��b�p�[���\�b�h�B
'*          ���ׂĂ̕ύX��ۑ����Č��݂̃g�����U�N�V�������I������
'*          �V�����g�����U�N�V�������J�n����ꍇ������
'*
'*          Connection �I�u�W�F�N�g�� Attributes �v���p�e�B�ɂ���ẮA
'*          CommitTrans ���\�b�h�܂��� RollbackTrans ���\�b�h���Ăяo���ƁA
'*          �V�����g�����U�N�V�����������I�ɊJ�n����ꍇ������B
'*          Attributes �v���p�e�B�� adXactCommitRetaining �ɐݒ肳��Ă���ꍇ
'*          �́A CommitTrans ���\�b�h���Ăяo���ƁA�v���o�C�_�[���V�����g����
'*          �U�N�V�����������I�ɊJ�n����B
'*
'******************************************************************************
Sub CommitTrans()
    Call mOrigin.CommitTrans
End Sub

'******************************************************************************
'* [�T  �v] Execute���\�b�h
'* [��  ��] Execute�̃��b�p�[���\�b�h�B
'*          �w�肳�ꂽ�N�G���ASQL �X�e�[�g�����g�A�X�g�A�h �v���V�[�W���A�܂���
'*          �v���o�C�_�[�ŗL�̃e�L�X�g�����s����B
'*
'*          Connection �I�u�W�F�N�g�� Execute ���\�b�h���g�p����ƁACommandText
'*          �����Ń��\�b�h�ɓn�����N�G�����A���̐ڑ��Ŏ��s�����B
'*          CommandText �����ōs��Ԃ��N�G�����w�肷��ƁA���s���ʂ͂��ׂĐV����
'*          Recordset �I�u�W�F�N�g�Ɋi�[�����B
'*          �R�}���h�����ʂ�Ԃ����Ƃ��Ӑ}���Ă��Ȃ��ꍇ (SQL UPDATE �N�G����
'*          ��)�A�v���o�C�_�[�̓I�v�V����adExecuteNoRecords���w�肳��Ă����
'*          ��A Nothing��Ԃ��B
'*          ����ȊO�̏ꍇ�AExecute �͕���ꂽRecordset��Ԃ��B
'*          �Ԃ���� Recordset �I�u�W�F�N�g�͏�ɁA�ǂݎ���p�őO���X�N���[��
'*          �^�C�v�̃J�[�\���ł���B
'*          ��葽���̋@�\������� Recordset �I�u�W�F�N�g���K�v�ȏꍇ�́A�܂�
'*          �ړI�̃v���p�e�B�ݒ������ Recordset �I�u�W�F�N�g���쐬���A����
'*          Recordset �I�u�W�F�N�g�� Open ���\�b�h���g�p���ăN�G�������s���āA
'*          �ړI�̃J�[�\���̎�ނ�Ԃ��B
'*          CommandText�����̓��e�̓v���o�C�_�[�ɌŗL�̂��̂ł���A�W���� SQL
'*          �\���܂��̓v���o�C�_�[���T�|�[�g������ʂȃR�}���h�`�����g�p�ł���B
'*          ���̑��삪��������ƤExecuteComplete �C�x���g����������
'*
'* @param CommandText     ���s���� SQL �X�e�[�g�����g�A�X�g�A�h �v���V�[�W���A
'*                        URL�A�܂��̓v���o�C�_�[�ŗL�̃e�L�X�g�B
'*                        �e�[�u�������g�p���邱�Ƃ��ł��邪�A�v���o�C�_�[��
'*                        SQL �Ή��ł���ꍇ�Ɍ�����B
'*                        ���Ƃ��΁A"Customers" �Ƃ����e�[�u�������g�p����ƁA
'*                        SQL �̕W���� Select �\���������I�ɑO�ɕt�������
'*                        "SELECT * FROM Customers" �Ƃ����X�e�[�g�����g���쐬
'*                        ����AT-SQL�X�e�[�g�����g�Ƃ��ăv���o�C�_�[�ɓn�����B
'* @param RecordsAffected �ȗ��\�B���̑���̉e�����󂯂����R�[�h���B
'* @param Options         �ȗ��\�BMoveRecordOptionsEnum �l���w��B
'*                        �v���o�C�_�[�� CommandText ������]��������@�������B
'*                        1�܂��͕����� CommandTypeEnum �l�܂���
'*                        ExecuteOptionEnum �l�̃r�b�g�}�X�N���w��ł���B
'*                        [!����] �����������ŏ����ɗ}���ăp�t�H�[�}���X������
'*                        ������ɂ́A ExecuteOptionEnum �̒l adExecuteNoRecords
'*                        ���g�p����B
'*                        Execute �ł́A CommandTypeEnum �̒l adCmdFile �܂���
'*                        adCmdTableDirect ���g�p���Ȃ����ƁB
'*                        �����̒l�́ARecordset �� Open ���\�b�h����� Requery
'*                        ���\�b�h�̃I�v�V�����Ƃ��Ă̂ݎg�p�ł���B
'*
'* @return Recordset �I�u�W�F�N�g�̎Q��
'******************************************************************************
Function Execute(CommandText As String, _
                 Optional RecordsAffected, _
                 Optional Options As Long = -1) As ADODBExRecordset
    Dim wrap As ADODBExRecordset: Set wrap = New ADODBExRecordset
    
    If IsMissing(RecordsAffected) Then
        Set wrap.Origin = mOrigin.Execute(CommandText, , Options)
    Else
        Set wrap.Origin = mOrigin.Execute(CommandText, RecordsAffected, Options)
    End If
    Set Execute = wrap
End Function

'******************************************************************************
'* [�T  �v] OpenCn���\�b�h
'* [��  ��] Open�̃��b�p�[���\�b�h�B
'*          �f�[�^�\�[�X�ւ̐ڑ����J���
'*
'*          Connection �I�u�W�F�N�g�� Open ���\�b�h���g�p����ƁA�f�[�^�\�[�X
'*          �ւ̕����I�Ȑڑ����m�������B
'*          ���̃��\�b�h������ɏI������ƁA�ڑ����L���ɂȂ�A�ڑ��ɑ΂��ăR�}
'*          ���h�𔭍s���āA���̌��ʂ��������邱�Ƃ��ł���B
'*          �ȗ��\��ConnectionString�������g�p���āA�Z�~�R�����ŋ�؂�ꂽ
'*          ��A�̈��� = value�X�e�[�g�����g���܂ސڑ���������w�肷�邩�A
'*          �܂��� URL �Ŏ��ʂ����t�@�C���܂��̓f�B���N�g�����\�[�X���w�肷��B
'*          connectionstring�v���p�e�B�́A connectionstring�����Ɏg�p���ꂽ�l��
'*          �����I�Ɍp������B
'*          ���̂��߁A connection�I�u�W�F�N�g���J���O�ɁA connectionstring
'*          �v���p�e�B��ݒ肷�邩�A�܂���connectionstring�������g�p���āA
'*          Open���\�b�h�̌Ăяo�����Ɍ��݂̐ڑ��p�����[�^�[��ݒ�܂��͏㏑��
'*          ���邱�Ƃ��ł���B
'*          ���[�U�[�ƃp�X���[�h�̏����AConnectionString �����ƁA�ȗ��\��
'*          UserID ��������� Password �����̗����œn���ƁAUserID ������
'*          Password �����̕����AConnectionString �Ŏw�肵���l���D�悳���B
'*          �J���Ă��� Connection �I�u�W�F�N�g�ɑ΂��鑀�삪����������AClose
'*          ���\�b�h���g�p���Ċ֘A���邷�ׂẴV�X�e�� ���\�[�X���������B
'*          �I�u�W�F�N�g����Ă��A�I�u�W�F�N�g�̓���������폜����Ȃ��B
'*          �I�u�W�F�N�g�̃v���p�e�B�ݒ��ύX���AOpen ���\�b�h���g�p���āA
'*          ��ōēx�J�����Ƃ��ł���B
'*          �I�u�W�F�N�g�����������犮�S�ɍ폜����ɂ́A�I�u�W�F�N�g�ϐ���
'*          Nothing ��ݒ肷��B
'*          �����[�g�f�[�^�T�[�r�X�̎g�p�󋵃N���C�A���g����connection�I�u�W�F
'*          �N�g�Ŏg�p����ꍇ�A connection�I�u�W�F�N�g��Recordset���J�����
'*          �܂ŁA Open���\�b�h�͎��ۂɂ̓T�[�o�[�ւ̐ڑ����m�����Ȃ��B
'*
'* @param ConnectionString �ȗ��\�B�ڑ���� �̒l�B
'*                         �L���Ȑݒ�l�̏ڍׂɂ��ẮA ConnectionString
'*                         �v���p�e�B���Q�ƁB
'* @param UserID           �ȗ��\�B�ڑ����m������Ƃ��Ɏg�p���郆�[�U�[���B
'* @param Password         �ȗ��\�B�ڑ����m������Ƃ��Ɏg�p����p�X���[�h�B
'* @param Options          �ȗ��\�B�ڑ����m�����ꂽ�� (����) �܂��͐ڑ���
'*                         �m�������O (�񓯊�) �̂ǂ���ł��̃��\�b�h���Ԃ�
'*                         �����w�肷�� ConnectOptionEnum �l�B
'******************************************************************************
Sub OpenCn(Optional ConnectionString As String, _
           Optional UserID As String, _
           Optional Password As String, _
           Optional Options As Long = -1)
    Call mOrigin.Open(ConnectionString, UserID, Password, Options)
End Sub
         
'******************************************************************************
'* [�T  �v] OpenSchema���\�b�h
'* [��  ��] OpenSchema�̃��b�p�[���\�b�h�B
'*          �v���o�C�_�[����f�[�^�x�[�X�̃X�L�[�}�����擾����
'*
'*          OpenSchema ���\�b�h�́A�f�[�^�\�[�X�Ɋ܂܂��e�[�u���A�e�[�u����
'*          �܂܂���A�T�|�[�g����Ă���f�[�^�^�Ȃǂ̃f�[�^�\�[�X�Ɋւ���
'*          ����Ԃ��B
'*          QueryType �����́A�Ԃ����� (�X�L�[�}) ������ GUID �ł���BOLE DB
'*          �̎d�l�ɂ́A���ׂẴX�L�[�}�̈ꗗ������B
'*          ����Restrictions�́A�X�L�[�}�N�G���̌��ʂ𐧌�����B
'*          Restrictions�ɂ́A���ʂ�Recordset�ŁA�Ή������̗�̃T�u�Z�b�g
'*           (�����)�ɕK�v�Ȓl�̔z����w�肷��B
'*          OLE DB �d�l�ȊO�̔�W���X�L�[�} �N�G�����v���o�C�_�[���Ǝ��ɒ�`����
'*          ����ꍇ�́AQueryType ������ adSchemaProviderSpecific ���g�p����B
'*          ���̒萔���g�p����ꍇ�́ASchemaID �����ɁA���s����X�L�[�} �N�G��
'*          �� GUID ���w�肷��K�v������B
'*          QueryType �� adSchemaProviderSpecific �ɐݒ肳��ASchemaID ���w��
'*          ����Ă��Ȃ��ꍇ�A�G���[���������Ȃ��B
'*          �v���o�C�_�[�́A���ׂĂ� OLE DB �W���X�L�[�} �N�G�����T�|�[�g����
'*          �K�v�͂Ȃ��B
'*          OLE DB �̎d�l�ł́AadSchemaTables�AadSchemaColumns�A�����
'*          adSchemaProviderTypes �݂̂��K�v�Ƃ����B
'*          �������A�����̃X�L�[�} �N�G���ł́A�v���o�C�_�[�� Criteria ��
'*          ������T�|�[�g����K�v�͂Ȃ��B
'*          �����[�g�f�[�^�T�[�r�X�̎g�p�� OpenSchema���\�b�h�́A�N���C�A��
'*          �g����Connection�I�u�W�F�N�g�ł͎g�p�ł��Ȃ��B
'*
'* @param Schema       ���s����X�L�[�}�N�G���̎�ނ�\�� SchemaEnum �l���w��
'* @param Restrictions �ȗ��\�B SchemaEnum �̎w��ɏ]���āA�eQueryType�I�v
'*                     �V�����̃N�G������̔z����w��
'* @param SchemaID     OLE DB �̎d�l�Œ�`����Ă��Ȃ��A�v���o�C�_�[�X�L�[�}
'*                     �̃N�G���� GUID (�O���[�o����ӎ��ʎq) ���w��B
'*                     ���̃p�����[�^�[�́AQueryType �� adSchemaProviderSpecific
'*                     �ɐݒ肳��Ă���ꍇ�͕K�{�B
'*                     ����ȊO�̏ꍇ�A���̃p�����[�^�[�͎g�p���Ȃ��B
'* @return �X�L�[�}�����܂� Recordset �I�u�W�F�N�g�B Recordset �͓ǂݎ��
'*         ��p�̐ÓI�J�[�\���Ƃ��ĊJ�����BQueryType �ɂ��ARecordset ��
'*         �\�������񂪌��肳���B
'******************************************************************************
Function OpenSchema(Schema As SchemaEnum, _
                    Optional Restrictions, _
                    Optional SchemaID) As ADODBExRecordset
    Dim wrap As ADODBExRecordset: Set wrap = New ADODBExRecordset
    
    If IsMissing(Restrictions) Then
        If IsMissing(SchemaID) Then
            Set wrap.Origin = mOrigin.OpenSchema(Schema)
        Else
            Set wrap.Origin = mOrigin.OpenSchema(Schema, , SchemaID)
        End If
    ElseIf IsMissing(SchemaID) Then
        Set wrap.Origin = mOrigin.OpenSchema(Schema, Restrictions)
    Else
        Set wrap.Origin = mOrigin.OpenSchema(Schema, Restrictions, SchemaID)
    End If
    Set OpenSchema = wrap
End Function

'******************************************************************************
'* [�T  �v] RollbackTrans���\�b�h
'* [��  ��] RollbackTrans�̃��b�p�[���\�b�h�B
'*          ���݂̃g�����U�N�V�����̊Ԃɍs��ꂽ���ׂĂ̕ύX���L�����Z�����A
'*          �g�����U�N�V�������I������B
'*          �V�����g�����U�N�V�������J�n����ꍇ������B
'*
'*          Connection �I�u�W�F�N�g�� Attributes �v���p�e�B�ɂ���ẮA
'*          CommitTrans ���\�b�h�܂��� RollbackTrans ���\�b�h���Ăяo���ƁA
'*          �V�����g�����U�N�V�����������I�ɊJ�n����ꍇ������B
'*          Attributes �v���p�e�B�� adXactAbortRetaining �ɐݒ肳��Ă���ꍇ
'*          �́A RollbackTrans ���\�h���Ăяo���ƁA�v���o�C�_�[���V�����g����
'*          �U�N�V�����������I�ɊJ�n����B
'*
'******************************************************************************
Sub RollbackTrans()
    Call mOrigin.RollbackTrans
End Sub


'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------
'******************************************************************************
'* [�T  �v] OpenPostgreSql ���\�b�h
'* [��  ��] PostgreSql�f�[�^�x�[�X�ւ̐ڑ����J���
'*          PostgreSql�ڑ���������쐬���AConnection.Open ���\�b�h�����s����B
'*
'* @param Server   �ڑ���T�[�o��
'* @param Port     �ڑ���|�[�g�ԍ�
'* @param Database �ڑ���f�[�^�x�[�X��
'* @param UserID   �ȗ��\�B�ڑ����m������Ƃ��Ɏg�p���郆�[�U�[���B
'* @param Password �ȗ��\�B�ڑ����m������Ƃ��Ɏg�p����p�X���[�h�B
'* @param Options  �ȗ��\�B�ڑ����m�����ꂽ�� (����) �܂��͐ڑ���
'*                 �m�������O (�񓯊�) �̂ǂ���ł��̃��\�b�h���Ԃ�
'*                 �����w�肷�� ConnectOptionEnum �l�B
'******************************************************************************
Sub OpenPostgreSql(Server As String, _
                   Port As Integer, _
                   Database As String, _
                   Optional UserID As String, _
                   Optional Password As String, _
                   Optional Options As Long = -1)
    Dim cnnStr As String
    cnnStr = "Driver={PostgreSQL};" & _
             "Server=" & Server & ";" & _
             "Port=" & Port & ";" & _
             "Database=" & Database
    Call mOrigin.Open(cnnStr, UserID, Password, Options)
End Sub