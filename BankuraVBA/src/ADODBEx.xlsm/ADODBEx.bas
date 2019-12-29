Attribute VB_Name = "ADODBEx"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ADODB���b�v�E�g�����W���[��
'* [��  ��] ADODB��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�AADODB���g�p����
'*          ���[�e�B���e�B��񋟂���B
'*          ���b�v����ADODB���C�u�����͈ȉ��̂��̂Ƃ���B
'*              [name] Microsoft ActiveX Data Objects 6.1 Library
'*              [dll] C:\Program Files\Common Files\System\ado\msado15.dll
'* [�Q  �l]
'*  <https://docs.microsoft.com/ja-jp/sql/ado/microsoft-activex-data-objects-ado?view=sql-server-2017>
'*  <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-enumerated-constants>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* RDS �� Recordset �I�u�W�F�N�g�ɑ΂��āA�f�[�^���擾����񓯊��X���b�h��
'* ���s�D��x��\���܂��
'*-----------------------------------------------------------------------------
Public Enum ADCPROP_ASYNCTHREADPRIORITY_ENUM
    adPriorityLowest = 1      '�D��x���\�ȍŒ჌�x���ɐݒ肵�܂��B
    adPriorityBelowNormal = 2 '�D��x���Œ�ƕW���̊Ԃɐݒ肵�܂��B
    adPriorityNormal = 3      '�D��x��W���ɐݒ肵�܂��B
    adPriorityAboveNormal = 4 '�D��x��W���ƍō��̊Ԃɐݒ肵�܂��B
    adPriorityHighest = 5     '�D��x���\�ȍō����x���ɐݒ肵�܂��B
End Enum

'*-----------------------------------------------------------------------------
'* �K�w Recordset �̏W�v��ƌv�Z��� MSDataShape �v���o�C�_�[�����Čv�Z���邩
'* ���w�肵�܂��B
'*-----------------------------------------------------------------------------
Public Enum ADCPROP_AUTORECALC_ENUM
    adRecalcAlways = 1  '����l�ł��B�v�Z�񂪈ˑ�����l���ύX���ꂽ�� MSDataShape �v���o�C�_�[�����f�����Ƃ��ɍČv�Z���܂��B
    adRecalcUpFront = 0 '�K�w Recordset �̍ŏ��̍쐬���̂݌v�Z���܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Recordset �I�u�W�F�N�g���g�p���ăf�[�^ �\�[�X�s�̋��L�I�X�V���s���ۂɁA����
'* �̌��o�Ɏg�p����t�B�[���h��\���܂��B
'*-----------------------------------------------------------------------------
Public Enum ADCPROP_UPDATECRITERIA_ENUM
    adCriteriaAllCols = 1   '�f�[�^ �\�[�X�s�̗񂪕ύX���ꂽ�ꍇ�ɋ��������o���܂��B
    adCriteriaKey = 0       '�f�[�^ �\�[�X�s�̃L�[�񂪕ύX���ꂽ�ꍇ�A�܂�s���폜���ꂽ�ꍇ�ɋ��������o���܂��B
    adCriteriaTimeStamp = 3 '�f�[�^ �\�[�X�s�̃^�C���X�^���v���ύX���ꂽ�ꍇ�A�܂� Recordset ���擾������ɍs�ɃA�N�Z�X���������ꍇ�ɋ��������o���܂��B
    adCriteriaUpdCols = 2   'Recordset �̍X�V���ꂽ�t�B�[���h�ɑΉ�����f�[�^ �\�[�X�s�̗񂪕ύX���ꂽ�ꍇ�ɋ��������o���܂��B
End Enum

'*-----------------------------------------------------------------------------
'* UpdateBatch ���\�b�h�ɈÖق� Resync ���\�b�h���삪�������ǂ����������A����
'* �ꍇ�͂��̑���̓K�p�͈͂��w�肵�܂��B
'*-----------------------------------------------------------------------------
Public Enum ADCPROP_UPDATERESYNC_ENUM
    adResyncAll = 15          '���̂��ׂĂ� ADCPROP_UPDATERESYNC_ENUM �����o�[�̌������ꂽ�l���g�p���āAResync ���Ăяo���܂��B
    adResyncAutoIncrement = 1 '����l�ł��BMicrosoft Jet AutoNumber �t�B�[���h�� Microsoft SQL Server �� Identity ��ȂǁA�f�[�^ �\�[�X�ɂ���Ď����I�ɑ����܂��͐���������̐V���� ID �l���擾���܂��B
    adResyncConflicts = 2     '�������s�̋����ɂ��X�V����܂��͍폜���삪���s�������ׂĂ̍s�ɂ��āAResync ���Ăяo���܂��B
    adResyncInserts = 8       '����ɑ}�����ꂽ���ׂĂ̍s�ɂ��āAResync ���Ăяo���܂��B �������AAutoIncrement ��̒l�͍ē�������܂���B ����ɁA�����̎�L�[�̒l�Ɋ�Â��āA�V�����}�����ꂽ�s�̓��e���ē�������܂��B ��L�[�� AutoIncrement �l�̏ꍇ�AResync �ł͑Ώۂ̍s�̓��e���擾���܂���B �I�[�g�C���N�������g�̎�L�[�l�������I�ɃC���N�������g����ɂ́A adResyncAutoIncrement + adResyncInserts��g�ݍ��킹���l��UpdateBatch���Ăяo���܂��B
    adResyncNone = 0          'Resync ���Ăяo���܂���B
    adResyncUpdates = 4       '����ɍX�V���ꂽ���ׂĂ̍s�ɂ��āAResync ���Ăяo���܂��B
End Enum

'*-----------------------------------------------------------------------------
'* ����̑ΏۂƂȂ郌�R�[�h��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum AffectEnum
    adAffectAll = 3         'Recordset �ɓK�p����Ă��� Filter ���Ȃ��ꍇ�A���ׂẴ��R�[�h���Ώۂł��B Filter�v���p�e�B�������񒊏o���� (Author = ' Smith ' ""�Ȃ�) �ɐݒ肳��Ă���ꍇ�A����͌��݂̃`���v�^�[�̕\������Ă��郌�R�[�h�ɉe�����܂��B filter�v���p�e�B��filtergroupenum�̃����o�[�܂��̓u�b�N�}�[�N�̔z��ɐݒ肳��Ă���ꍇ�A���̑����Recordset�̂��ׂĂ̍s�ɉe�����܂��B
    adAffectAllChapters = 4 '���ݓK�p����Ă��� Filter �Ŕ�\���ɂȂ��Ă��郌�R�[�h���܂ށARecordset �̂��ׂĂ̌Z��`���v�^�[�̑S���R�[�h�ɔ��f����܂��B
    adAffectCurrent = 1     '���݂̃��R�[�h�ɂ̂ݔ��f����܂��B
    adAffectGroup = 2       '���݂� Filter �v���p�e�B�̐ݒ�𖞂������R�[�h�ɂ̂ݔ��f����܂��B���̃I�v�V�������g�p����ɂ́AFilter �v���p�e�B�� FilterGroupEnum �l�܂��� Bookmark �̔z��ɐݒ肷��K�v������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* ����̊J�n�ʒu�������u�b�N�}�[�N��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum BookmarkEnum
    adBookmarkCurrent = 0 '���݂̃��R�[�h����J�n���܂��B
    adBookmarkFirst = 1   '�ŏ��̃��R�[�h����J�n���܂��B
    adBookmarkLast = 2    '�Ō�̃��R�[�h����J�n���܂��B
End Enum

'*-----------------------------------------------------------------------------
'* �R�}���h�����̉��ߕ��@��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum CommandTypeEnum
    adCmdUnspecified = -1  'Does not specify the command type argument.
    adCmdText = 1          'CommandText ���A�R�}���h�܂��̓X�g�A�h �v���V�[�W���̃e�L�X�g��`�Ƃ��ĕ]�����܂��B
    adCmdTable = 2         'CommandText ���A�����I�ɐ������ꂽ SQL �N�G������Ԃ��ꂽ��݂̂ō\�������e�[�u�����Ƃ��ĕ]�����܂��B
    adCmdStoredProc = 4    'CommandText ���X�g�A�h �v���V�[�W�����Ƃ��ĕ]�����܂��B
    adCmdUnknown = 8       '����l�BCommandText �v���p�e�B�̃R�}���h�̎�ނ��s���ł��邱�Ƃ������܂��B
    adCmdFile = 256        'CommandText ���A�ۑ����ꂽ Recordset �̃t�@�C�����Ƃ��ĕ]�����܂��BRecordset.Open �܂��� Requery �Ƒg�ݍ��킹�Ă̂ݎg�p�ł��܂��B
    adCmdTableDirect = 512 'CommandText ���A���ׂĂ̗񂪕Ԃ��ꂽ�e�[�u�����Ƃ��ĕ]�����܂��B Recordset.Open �܂��� Requery �Ƒg�ݍ��킹�Ă̂ݎg�p�ł��܂��B Seek ���\�b�h���g�p����ꍇ�ARecordset �� adCmdTableDirect ���w�肵�ĊJ���K�v������܂��B ���̒l�́AExecuteOptionEnum �̒l adAsyncExecute �Ƒg�ݍ��킹�Ďg�p�ł��܂���B
End Enum

'*-----------------------------------------------------------------------------
'* �u�b�N�}�[�N�ŕ\���ꂽ 2 �̃��R�[�h�̑��Έʒu��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum CompareEnum
    adCompareEqual = 1         '�u�b�N�}�[�N�����������Ƃ������܂��B
    adCompareGreaterThan = 2   '�ŏ��̃u�b�N�}�[�N�� 2 �Ԗڂ̃u�b�N�}�[�N�̌�ɂȂ邱�Ƃ������܂��B
    adCompareLessThan = 0      '�ŏ��̃u�b�N�}�[�N�� 2 �Ԗڂ̃u�b�N�}�[�N�̑O�ɂȂ邱�Ƃ������܂��B
    adCompareNotComparable = 4 '�u�b�N�}�[�N���r�ł��Ȃ����Ƃ������܂��B
    adCompareNotEqual = 3      '2 �̃u�b�N�}�[�N�͈قȂ��Ă���A���ʂ��Ȃ����Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Connection ���̃f�[�^�̕ҏW�ARecord �̃I�[�v���A�܂��� Record ����� Stream
'*  �I�u�W�F�N�g�� Mode �v���p�e�B�̒l�̎w��ɑ΂��錠����\���܂��B
'*-----------------------------------------------------------------------------
Public Enum ConnectModeEnum
    adModeRead = 1            '�ǂݎ���p�̌�����\���܂��B
    adModeReadWrite = 3       '�ǂݎ��/�������ݗ����̌�����\���܂��B
    adModeRecursive = 4194304 '���̋��L���ےl (adModeShareDenyNone�A adModeShareDenyWrite�A�܂���adModeShareDenyRead) �Ƌ��Ɏg�p���āA���݂̃��R�[�h�̂��ׂẴT�u���R�[�h�ɋ��L������`�B���܂��B Record �Ɏq���Ȃ��ꍇ�͋@�\���܂���BadModeShareDenyNone �݂̂Ƒg�ݍ��킹�Ďg�p����ƁA���s���G���[���������܂��B �������A���̑��̒l�Ƒg�ݍ��킹���ꍇ�� adModeShareDenyNone �Ƒg�ݍ��킹�Ďg�p�ł��܂��B
    adModeShareDenyNone = 16  '�����̎�ނɊ֌W�Ȃ��A���̃��[�U�[���ڑ����J����悤�ɂ��܂��B���̃��[�U�[�ɑ΂��āA�ǂݎ��Ə������݂̗����̃A�N�Z�X�������܂��B
    adModeShareDenyRead = 4   '���̃��[�U�[���ǂݎ�茠���Őڑ����J���̂��֎~���܂��B
    adModeShareDenyWrite = 8  '���̃��[�U�[���������݌����Őڑ����J���̂��֎~���܂��B
    adModeShareExclusive = 12 '���̃��[�U�[���ڑ����J���̂��֎~���܂��B
    adModeUnknown = 0         '����l�B�������ݒ肳��Ă��Ȃ����A�����𔻒�ł��Ȃ����Ƃ������܂��B
    adModeWrite = 2           '�������ݐ�p�̌����������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Connection �I�u�W�F�N�g�� Open ���\�b�h���琧�䂪�߂�̂��A�ڑ��m���̌�
'* (����) ���O (�񓯊�) ����\���܂��B
'*-----------------------------------------------------------------------------
Public Enum ConnectOptionEnum
    adAsyncConnect = 16       '�ڑ���񓯊��ŊJ���܂��B�ڑ��\�����ǂ����𔻕ʂ��邽�߂ɁAConnectComplete �C�x���g���g�p�����ꍇ������܂��B
    adConnectUnspecified = -1 'Default. Opens the connection synchronously.
End Enum

'*-----------------------------------------------------------------------------
'* �f�[�^ �\�[�X�Ƃ̐ڑ����J���Ƃ��ɁA�s�����Ă���p�����[�^�[��v������_�C�A
'* ���O �{�b�N�X��\�����邩�ǂ�����\���܂��B
'*-----------------------------------------------------------------------------
Public Enum ConnectPromptEnum
    adPromptAlways = 1           '��ɗv�����܂��B
    adPromptComplete = 2         '����ɏ�񂪕K�v�ȏꍇ�ɗv�����܂��B
    adPromptCompleteRequired = 3 '����ɏ�񂪕K�v�����A�C�ӂ̃p�����[�^�[���֎~����Ă���ꍇ�ɗv�����܂��B
    adPromptNever = 4            '�v�����܂���B
End Enum

'*-----------------------------------------------------------------------------
'* CopyRecord ���\�b�h�̓����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum CopyRecordOptionsEnum
    adCopyAllowEmulation = 4 '"�R�s�[��" ���ʂ̃T�[�o�[�ɂ��邩�A"�R�s�[��" �ȊO�̃v���o�C�_�[�̃T�[�r�X���󂯂Ă��邽�߂ɂ��̃��\�b�h�����s�����ꍇ�A"�R�s�[��" �v���o�C�_�[���_�E�����[�h����ƃA�b�v���[�h������s���ăR�s�[���V�~�����[�g���悤�Ƃ��邱�Ƃ������܂��B�v���o�C�_�[�̋@�\���قȂ�ƁA�p�t�H�[�}���X���ቺ������f�[�^�������邱�Ƃ�����܂��B
    adCopyNonRecursive = 2   '�R�s�[��Ɍ��݂̃f�B���N�g�����R�s�[���܂����A�T�u�f�B���N�g���̓R�s�[���܂���B�R�s�[����͍ċA�I�ł͂���܂���B
    adCopyOverWrite = 1      '"�R�s�[��" �������̃t�@�C����f�B���N�g�����w���ꍇ�A���̃t�@�C����f�B���N�g�����㏑�����܂��B
    adCopyUnspecified = -1   '����l�B����̃R�s�[��������s���܂��B�R�s�[����͍ċA�I�ɍs���A�R�s�[��̃t�@�C����f�B���N�g�������ɑ��݂���ꍇ�͑��삪���s���܂��B
End Enum

'*-----------------------------------------------------------------------------
'* �J�[�\�� �T�[�r�X�̏ꏊ��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum CursorLocationEnum
    adUseClient = 3 '���[�J���J�[�\�����C�u�����ɂ���Ē񋟂����N���C�A���g���J�[�\�����g�p���܂��B �����̏ꍇ�A���[�J�� �J�[�\�� �T�[�r�X�ɂ̓h���C�o�[�ɂ���Ē񋟂����J�[�\�����������̃J�[�\���@�\������̂ŁA���̐ݒ�𗘗p����ƁA��荂�x�ȋ@�\��񋟂ł��܂��B �ȑO�̃o�[�W�����Ƃ̌݊�����ۂ��߂ɁA�����Ӗ������� adUseClientBatch ���T�|�[�g���Ă��܂��B
    adUseServer = 2 '����l�B �f�[�^ �v���o�C�_�[ �J�[�\���܂��̓h���C�o�[�ɂ���Ē񋟂����J�[�\�����g�p���܂��B �����̃J�[�\���́A�����̏ꍇ�_��������A���̃��[�U�[���s���f�[�^ �\�[�X�ւ̕ύX�����o�ł��܂��B �������A Microsoft Cursor Service for OLE DB (�֘A�t�����Ă��Ȃ�Recordset�I�u�W�F�N�g�Ȃ�) �̈ꕔ�̋@�\�́A�T�[�o�[���J�[�\�����g�p���ăV�~�����[�g���邱�Ƃ͂ł��܂���B���̐ݒ�ł́A�����̋@�\�͎g�p�ł��܂���B
    adUseNone = 1   'Does not use cursor services. (This constant is obsolete and appears solely for the sake of backward compatibility.)
End Enum

'*-----------------------------------------------------------------------------
'* Supports ���\�b�h���e�X�g����@�\��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum CursorOptionEnum
    adAddNew = 16778240      '�V�������R�[�h��ǉ����� AddNew ���\�b�h���T�|�[�g���܂��B
    adApproxPosition = 16384 'AbsolutePosition �v���p�e�B�� AbsolutePage �v���p�e�B���T�|�[�g���܂��B
    adBookmark = 8192        '����̃��R�[�h�ւ̃A�N�Z�X���m�ۂ��� Bookmark �v���p�e�B���T�|�[�g���܂��B
    adDelete = 16779264      '���R�[�h���폜���� Delete ���\�b�h���T�|�[�g���܂��B
    adFind = 524288          'Recordset ���̍s�̈ʒu���m�F���� Find ���\�b�h���T�|�[�g���܂��B
    adHoldRecords = 256      '�ۗ����̂��ׂĂ̕ύX���R�~�b�g�����ɁA�V���ȃ��R�[�h���i�[���邩�A�܂��͎��̊i�[�ʒu��ύX���܂��B
    adIndex = 8388608        '�C���f�b�N�X�ɖ��O��t���� Index �v���p�e�B���T�|�[�g���܂��B
    adMovePrevious = 512     '�u�b�N�}�[�N���g�p�����Ɍ��݂̃��R�[�h�̈ʒu������Ɉړ����� MoveFirst ���\�b�h�� MovePrevious ���\�b�h�A����� Move ���\�b�h�� GetRows ���\�b�h���T�|�[�g���܂��B
    adNotify = 262144        '��ɂȂ�f�[�^ �v���o�C�_�[���ʒm���T�|�[�g���Ă��邱�Ƃ������܂� (����ɂ�� Recordset �C�x���g�̃T�|�[�g�̗L�������܂�܂�)�B
    adResync = 131072        '��ɂȂ�f�[�^�x�[�X�̃J�[�\���ɂ�����f�[�^���X�V���� Resync ���\�b�h���T�|�[�g���܂��B
    adSeek = 4194304         'Recordset ���̍s���������� Seek ���\�b�h���T�|�[�g���܂��B
    adUpdate = 16809984      '�����̃f�[�^��ύX���� Update ���\�b�h���T�|�[�g���܂��B
    adUpdateBatch = 65536    '�����̕ύX���O���[�v�Ƃ��ăv���o�C�_�[�ɑ��M����o�b�`�X�V (UpdateBatch ���\�b�h�� CancelBatch ���\�b�h) ���T�|�[�g���܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Recordset �I�u�W�F�N�g���g�p����J�[�\���̎�ނ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum CursorTypeEnum
    adOpenDynamic = 2      '���I�J�[�\�����g�p���܂��B���̃��[�U�[�ɂ��ǉ��A�ύX�A����э폜���\������A�v���o�C�_�[���u�b�N�}�[�N���T�|�[�g���Ă��Ȃ��ꍇ�������ARecordset ���ł̂��ׂĂ̎�ނ̈ړ��������܂��B
    adOpenForwardOnly = 0  '����l�B�O���X�N���[�� �J�[�\�����g�p���܂��B���R�[�h�̃X�N���[���������O�����Ɍ��肳��Ă��邱�Ƃ������A�ÓI�J�[�\���Ɠ������������܂��BRecordset �̃X�N���[���� 1 �񂾂��ŏ\���ȏꍇ�́A����ɂ���ăp�t�H�[�}���X������ł��܂��B
    adOpenKeyset = 1       '�L�[�Z�b�g �J�[�\�����g���܂��B������ Recordset ���瑼�̃��[�U�[���폜�������R�[�h�̓A�N�Z�X�ł��܂��񂪁A���̃��[�U�[���ǉ��������R�[�h�͕\���ł��Ȃ��_�������ē��I�J�[�\���Ɠ����ł��B���̃��[�U�[���ύX�����f�[�^�͕\���ł��܂��B
    adOpenStatic = 3       '�ÓI�J�[�\�����g�p���܂��B�f�[�^�̌����⃌�|�[�g�̐����Ɏg�p�ł���̐ÓI�R�s�[�ł��B���̃��[�U�[�ɂ��ǉ��A�ύX�A�܂��͍폜�͕\������܂���B
    adOpenUnspecified = -1 '�J�[�\���̎�ނ��w�肵�܂���B
End Enum

'*-----------------------------------------------------------------------------
'* �t�B�[���h�A�p�����[�^�[�A�܂��̓v���p�e�B�̃f�[�^�^���w�肵�܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum DataTypeEnum
    adArray = 8192          '��ɕʂ̃f�[�^�^�萔�Ƒg�ݍ��킳��A���̃f�[�^�^�̔z��������t���O�l�ł��B
    adBigInt = 20           '8 �o�C�g�̕����t�������������܂� (DBTYPE_I8)�B
    adBinary = 128          '�o�C�i���l�������܂� (DBTYPE_BYTES)�B
    adBoolean = 11          '�u�[���l�������܂� (DBTYPE_BOOL)�B
    adBSTR = 8              'null �ŏI��镶���� (Unicode) �������܂� (DBTYPE_BSTR)�B
    adChapter = 136         '�q�s�Z�b�g�̍s�����ʂ��� 4 �o�C�g �`���v�^�[�l�������܂� (DBTYPE_HCHAPTER)�B
    adChar = 129            '������l�������܂� (DBTYPE_STR)�B
    adCurrency = 6          '�ʉݒl�������܂� (DBTYPE_CY)�B�ʉ݌^�͏����_�ȉ� 4 ���̌Œ菬���_�̐��l�ł��B�X�P�[���� 10,000 �́A8 �o�C�g�̕����t�������Ŋi�[���܂��B
    adDate = 7              '���t�l�������܂� (DBTYPE_DATE)�B���t�͔{���x���������_���^ (Double) �Ŋi�[����A���������� 1899 �N 12 �� 30 ������̓������A���������͎�����\���܂��B
    adDBDate = 133          '���t�l (yyyymmdd) �������܂� (DBTYPE_DBDATE)�B
    adDBTime = 134          '�����l (hhmmss) �������܂� (DBTYPE_DBTIME)�B
    adDBTimeStamp = 135     '���t/�^�C�� �X�^���v (yyyymmddhhmmss ����� 10 ������ 1 ���܂ł̕���) �������܂� (DBTYPE_DBTIMESTAMP)�B
    adDecimal = 14          '�Œ萸�x����уX�P�[���̐��m�Ȑ��l�������܂� (DBTYPE_DECIMAL)�B
    adDouble = 5            '�{���x���������_�l�������܂� (DBTYPE_R8)�B
    adEmpty = 0             '�l���w�肵�܂��� (DBTYPE_EMPTY)�B
    adError = 10            '32 �r�b�g �G���[ �R�[�h�������܂� (DBTYPE_ERROR)�B
    adFileTime = 64         '1601 �N 1 �� 1 ������̎��Ԃ� 100 �i�m�b�P�ʂŎ��� 64 �r�b�g�l�������܂� (DBTYPE_FILETIME)�B
    adGUID = 72             '�O���[�o����ӎ��ʎq (GUID) �������܂� (DBTYPE_GUID)�B
    adIDispatch = 9         'COM �I�u�W�F�N�g�� IDispatch �C���^�[�t�F�C�X�ւ̃|�C���^�[�������܂� (DBTYPE_IDISPATCH)�B
    adInteger = 3           '4 �o�C�g�̕����t�������������܂� (DBTYPE_I4)�B
    adIUnknown = 13         'COM �I�u�W�F�N�g�� IUnknown �C���^�[�t�F�C�X�ւ̃|�C���^�[�������܂� (DBTYPE_IUNKNOWN)�B
    adLongVarBinary = 205   '�����O �o�C�i���l�������܂��B
    adLongVarChar = 201     '����������l�������܂��B
    adLongVarWChar = 203    '�����Anull �ŏI��� Unicode ������l�������܂��B
    adNumeric = 131         '�Œ萸�x����уX�P�[���̐��m�Ȑ��l�������܂� (DBTYPE_NUMERIC)�B
    adPropVariant = 138     '�I�[�g���[�V���� PROPVARIANT �������܂� (DBTYPE_PROP_VARIANT)�B
    adSingle = 4            '�P���x���������_�l�������܂� (DBTYPE_R4)�B
    adSmallInt = 2          '2 �o�C�g�̕����t�������������܂� (DBTYPE_I2)�B
    adTinyInt = 16          '1 �o�C�g�̕����t�������������܂� (DBTYPE_I1)�B
    adUnsignedBigInt = 21   '8 �o�C�g�̕����Ȃ������������܂� (DBTYPE_UI8)�B
    adUnsignedInt = 19      '4 �o�C�g�̕����Ȃ������������܂� (DBTYPE_UI4)�B
    adUnsignedSmallInt = 18 '2 �o�C�g�̕����Ȃ������������܂� (DBTYPE_UI2)�B
    adUnsignedTinyInt = 17  '1 �o�C�g�̕����Ȃ������������܂� (DBTYPE_UI1)�B
    adUserDefined = 132     '���[�U�[��`�̕ϐ��������܂� (DBTYPE_UDT)�B
    adVarBinary = 204       '�o�C�i���l�������܂� (Parameter �I�u�W�F�N�g�̂�)�B
    adVarChar = 200         '������l�������܂��B
    adVariant = 12          '�I�[�g���[�V���� �o���A���g�^ (Variant) �������܂� (DBTYPE_VARIANT)�B
    adVarNumeric = 139      '���l�������܂� (Parameter �I�u�W�F�N�g�̂�)�B
    adVarWChar = 202        'null �ŏI��� Unicode ������������܂��B
    adWChar = 130           'null �ŏI��� Unicode ������������܂� (DBTYPE_WSTR)�B
End Enum

'*-----------------------------------------------------------------------------
'* ���R�[�h�̕ҏW�󋵂������܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum EditModeEnum
    adEditNone = 0       '�i�s���̕ҏW���삪�Ȃ����Ƃ������܂��B
    adEditInProgress = 1 '���݂̃��R�[�h�̃f�[�^���ύX���ꂽ���A�ۑ�����Ă��Ȃ����Ƃ������܂��B
    adEditAdd = 2        'AddNew ���\�b�h���Ăяo����A�R�s�[ �o�b�t�@�[���̌��݂̃��R�[�h���A�f�[�^�x�[�X�ɕۑ�����Ă��Ȃ��V�������R�[�h�ł��邱�Ƃ������܂��B
    adEditDelete = 4     '���݂̃��R�[�h���폜���ꂽ���Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* ADO ���s���G���[�̎�ނ�\���܂��B
'* https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/errorvalueenum
'*-----------------------------------------------------------------------------
Public Enum ErrorValueEnum
    adErrBoundToCommand = 3707           'Command �I�u�W�F�N�g���\�[�X�Ɏ��� Recordset �I�u�W�F�N�g�� ActiveConnection �v���p�e�B��ύX�ł��܂���B
    adErrCannotComplete = 3732           '�T�[�o�[�͑���������ł��܂���B
    adErrCantChangeConnection = 3748     '�ڑ������ۂ���܂����B �v�����ꂽ�V�K�ڑ��̓��������ݎg�p���̓����ƈقȂ�܂��B
    adErrCantChangeProvider = 3220       '�w�肳�ꂽ�v���o�C�_�[�����Ɏg�p����Ă�����̂ƈقȂ�܂��B
    adErrCantConvertvalue = 3724         '�����̕s��v�܂��̓f�[�^ �I�[�o�[�t���[�ȊO�̗��R�ɂ��A�f�[�^�l��ϊ��ł��܂���B ���Ƃ��΁A�ϊ��ɂ��f�[�^�̈ꕔ���؂�̂Ă���ꍇ�Ȃǂł��B
    adErrCantCreate = 3725               '�t�B�[���h�̃f�[�^�^���s���ł��������A�v���o�C�_�[����������s����̂ɏ\���ȃ��\�[�X�������Ă��Ȃ��������߁A�f�[�^�l��ݒ�܂��͎擾�ł��܂���B
    adErrCatalogNotSet = 3747            '����ɂ͗L���� ParentCatalog ���K�v�ł��B
    adErrColumnNotOnThisRow = 3726       '���R�[�h�ɂ��̃t�B�[���h�����݂��܂���B
    adErrConnectionStringTooLong = 3754
    adErrDataConversion = 3421           '���݂̑���ɑ΂��āA�Ԉ�����^�̒l���g�p���Ă��܂��B
    adErrDataOverflow = 3721             '�f�[�^�l���傫�����邽�߂ɁA�t�B�[���h�̃f�[�^�^�ŕ\���ł��܂���B
    adErrDelResOutOfScope = 3738         '�폜�����I�u�W�F�N�g�� URL �͌��݂̃��R�[�h�͈̔͊O�ł��B
    adErrDenyNotSupported = 3750         '�v���o�C�_�[�����L�̐�����T�|�[�g���Ă��܂���B
    adErrDenyTypeNotSupported = 3751     '�v���o�C�_�[���A�v�����ꂽ��ނ̋��L�̐�����T�|�[�g���Ă��܂���B
    adErrFeatureNotAvailable = 3251      '�I�u�W�F�N�g�܂��̓v���o�C�_�[�͗v�����ꂽ��������s�ł��܂���B
    adErrFieldsUpdateFailed = 3749       '�t�B�[���h���X�V�ł��܂���ł����B �ڍׂɂ��ẮA�e Field �I�u�W�F�N�g�� Status �v���p�e�B���Q�Ƃ��Ă��������B
    adErrIllegalOperation = 3219         '���̃R���e�L�X�g�ő���͋�����Ă��܂���B
    adErrIntegrityViolation = 3719       '�f�[�^�̒l���t�B�[���h�̐���������ɔ����Ă��܂��B
    adErrInTransaction = 3246            '�g�����U�N�V�����̎��s���� Connection �I�u�W�F�N�g�𖾎��I�ɕ��邱�Ƃ��ł��܂���B
    adErrInvalidArgument = 3001          '�Ԉ������ނ܂��͋��e�͈͊O�̈������g�p���Ă��邩�A�g�p���Ă���������������Ă��܂��B
    adErrInvalidConnection = 3709        '���̑�������s���邽�߂ɐڑ����g�p�ł��܂���B ���̃R���e�L�X�g�ŕ��Ă��邩���邢�͖����ł��B
    adErrInvalidParamInfo = 3708         'Parameter �I�u�W�F�N�g���K�؂ɒ�`����Ă��܂���B ���������A�܂��͕s���S�ȏ�񂪎w�肳��܂����B
    adErrInvalidTransaction = 3714       '�����g�����U�N�V�����������ł��邩�A�J�n����Ă��܂���B
    adErrInvalidURL = 3729               'URL �ɖ����ȕ������܂܂�Ă��܂��B URL �����������͂���Ă��邩�m�F���Ă��������B
    adErrItemNotFound = 3265             '�v�����ꂽ���O�A�܂��͏����ɑΉ����鍀�ڂ��R���N�V�����Ō�����܂���B
    adErrNoCurrentRecord = 3021          'BOF �܂��� EOF �� True �ł��邩�A���݂̃��R�[�h���폜����Ă��܂��B �v�����ꂽ����ɂ͌��݂̃��R�[�h���K�v�ł��B
    adErrNotReentrant = 3710             '�C�x���g�������ɑ�����s�����Ƃ͂ł��܂���B
    adErrObjectClosed = 3704             '�I�u�W�F�N�g�����Ă���ꍇ�́A����͋�����܂���B
    adErrObjectInCollection = 3367       '�I�u�W�F�N�g�͊��ɃR���N�V�����ɑ��݂��܂��B �ǉ��ł��܂���B
    adErrObjectNotSet = 3420             '�I�u�W�F�N�g�������ɂȂ��Ă��܂��B
    adErrObjectOpen = 3705               '�I�u�W�F�N�g���J���Ă���ꍇ�́A����͋�����܂���B
    adErrOpeningFile = 3002              '�t�@�C�����J�����Ƃ��ł��܂���ł����B
    adErrOperationCancelled = 3712       '���[�U�[�ɂ�葀�삪��������܂����B
    adErrOutOfSpace = 3734               '��������s�ł��܂���B �v���o�C�_�[�ɂ���ď\���ȋL���悪�m�ۂł��܂���B
    adErrPermissionDenied = 3720         '�����s���̂��߃t�B�[���h�̏������݂͂ł��܂���B
    adErrPropConflicting = 3742
    adErrPropInvalidColumn = 3739
    adErrPropInvalidOption = 3740
    adErrPropInvalidValue = 3741
    adErrPropNotAllSettable = 3743
    adErrPropNotSet = 3744
    adErrPropNotSettable = 3745
    adErrPropNotSupported = 3746
    adErrProviderFailed = 3000           '�v���o�C�_�[���v�����ꂽ��������s�ł��܂���ł����B
    adErrProviderNotFound = 3706         '�v���o�C�_�[��������܂���B �������C���X�g�[������Ă��Ȃ��\��������܂��B
    adErrProviderNotSpecified = 3753
    adErrReadFile = 3003                 '�t�@�C����ǂݍ��ނ��Ƃ��ł��܂���ł����B
    adErrResourceExists = 3731           '�R�s�[��������s�ł��܂���B ����� URL �Ŏw�肳�ꂽ�I�u�W�F�N�g�͊��ɑ��݂��܂��B �I�u�W�F�N�g��u�������邽�߂ɂ� adCopyOverwrite ���w�肵�Ă��������B
    adErrResourceLocked = 3730           '�w�肳�ꂽ URL �ɂ���ĕ\���ꂽ�I�u�W�F�N�g�� 1 �ȏ�̑��̃v���Z�X�ɂ���ă��b�N����Ă��܂��B�v���Z�X���I������܂ő҂��āA������ēx���s���Ă��������B
    adErrResourceOutOfScope = 3735       '�\�[�X�܂��͈���� URL ���A���݂̃��R�[�h�͈̔͊O�ł��B
    adErrSchemaViolation = 3722          '�f�[�^�l���t�B�[���h�̃f�[�^�^�ƈ�v���Ă��Ȃ����A�t�B�[���h�̐���ɔ����Ă��܂��B
    adErrSignMismatch = 3723             '�f�[�^�̒l�͕����t���ł����A�v���o�C�_�[�ɂ���Ďg�p�����t�B�[���h �f�[�^�^�͕����Ȃ��̂��߁A�ϊ��Ɏ��s���܂����B
    adErrStillConnecting = 3713          '�񓯊�����ۗ̕����ɁA������s�����Ƃ͂ł��܂���B
    adErrStillExecuting = 3711           '�񓯊����s���ɑ�����s�����Ƃ͂ł��܂���B
    adErrTreePermissionDenied = 3728     '�������s�\���Ȃ��߂ɁA�c���[�܂��̓T�u�c���[�ɃA�N�Z�X�ł��܂���B
    adErrUnavailable = 3736              '����̊����Ɏ��s���A��Ԃ͗��p�ł��܂���B �t�B�[���h�����p�ł��Ȃ����A���삪���s����Ȃ������\��������܂��B
    adErrUnsafeOperation = 3716          '���̃R���s���[�^�[�̈��S���̐ݒ�ɂ��A���̃h���C���̃f�[�^ �\�[�X�ւ̃A�N�Z�X���֎~����Ă��܂��B
    adErrURLDoesNotExist = 3727          '�\�[�X URL �܂��͈���� URL �̐e�����݂��܂���B
    adErrURLNamedRowDoesNotExist = 3737  '���� URL �ɂ���Ė��O��t����ꂽ���R�[�h�����݂��܂���B
    adErrVolumeNotFound = 3733           '�v���o�C�_�[���AURL �Ŏ����ꂽ�L�����u�̏ꏊ�����ł��܂���B URL �����������͂���Ă��邩�m�F���Ă��������B
    adErrWriteFile = 3004                '�t�@�C���ւ̏������݂Ɏ��s���܂����B
    adwrnSecurityDialog = 3717           '�����g�p�̂��߂ɗp�ӂ���Ă��܂��B �g�p���Ȃ��ł��������B
    adwrnSecurityDialogHeader = 3718     '�����g�p�̂��߂ɗp�ӂ���Ă��܂��B �g�p���Ȃ��ł��������B
End Enum

'*-----------------------------------------------------------------------------
'* �C�x���g�������������R��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum EventReasonEnum
    adRsnAddNew = 1        '�V�������R�[�h���ǉ�����܂����B
    adRsnClose = 9         'Recordset �������܂����B
    adRsnDelete = 2        '���R�[�h���폜����܂����B
    adRsnFirstChange = 11  '���R�[�h�ɏ��߂Ă̕ύX���������܂����B
    adRsnMove = 10         'Recordset ���̃��R�[�h �|�C���^�[���ړ����܂����B
    adRsnMoveFirst = 12    '���R�[�h �|�C���^�[�� Recordset �̍ŏ��̃��R�[�h�Ɉړ����܂����B
    adRsnMoveLast = 15     '���R�[�h �|�C���^�[�� Recordset �̍Ō�̃��R�[�h�Ɉړ����܂����B
    adRsnMoveNext = 13     '���R�[�h �|�C���^�[�� Recordset �̎��̃��R�[�h�Ɉړ����܂����B
    adRsnMovePrevious = 14 '���R�[�h �|�C���^�[�� Recordset �̑O�̃��R�[�h�Ɉړ����܂����B
    adRsnRequery = 7       'Recordset ���ăN�G������܂����B
    adRsnResynch = 8       'Recordset ���f�[�^�x�[�X�ƍē������܂����B
    adRsnUndoAddNew = 5    '�V�������R�[�h�̒ǉ�����������܂����B
    adRsnUndoDelete = 6    '���R�[�h�̍폜����������܂����B
    adRsnUndoUpdate = 4    '���R�[�h�̍X�V����������܂����B
    adRsnUpdate = 3        '�����̃��R�[�h���X�V����܂����B
End Enum

'*-----------------------------------------------------------------------------
'* �C�x���g�̎��s�̌��݂̏�Ԃ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum EventStatusEnum
    adStatusCancel = 4         '�C�x���g�𔭐�����������̎�������v�����܂��B
    adStatusCantDeny = 3       '�ۗ����̑���̎�������v���ł��Ȃ����Ƃ������܂��B
    adStatusErrorsOccurred = 2 '�C�x���g�𔭐����������삪�G���[�ɂ���Ď��s�������Ƃ������܂��B
    adStatusOK = 1             '�C�x���g�𔭐����������삪�����������Ƃ������܂��B
    adStatusUnwantedEvent = 5  '�C�x���g ���\�b�h�̎��s���I������܂ŁA�㑱�̒ʒm���s���܂���B
End Enum

'*-----------------------------------------------------------------------------
'* �v���o�C�_�[�ɂ��R�}���h�̎��s���@��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum ExecuteOptionEnum
    adAsyncExecute = 16          '�R�}���h��񓯊��Ɏ��s���邱�Ƃ������܂��B ���̒l�́ACommandTypeEnum �̒l adCmdTableDirect �Ƒg�ݍ��킹�Ďg�p�ł��܂���B
    adAsyncFetch = 32            'CacheSize �v���p�e�B�Ŏw�肵�������ʂ̎c��̍s��񓯊��Ɏ擾���邱�Ƃ������܂��B
    adAsyncFetchNonBlocking = 64 '�擾���Ƀ��C�� �X���b�h���u���b�N���Ȃ����Ƃ������܂��B �v�����ꂽ�s���܂��擾����Ă��Ȃ��ꍇ�A���݂̍s�������I�Ƀt�@�C���̍Ō�Ɉړ����܂��B�i���I�ɕۑ����ꂽ Recordset ������ Stream ���� Recordset ���J�����ꍇ�AadAsyncFetchNonBlocking �͖����ɂȂ�A����͓����Ŏ��s����A�u���b�L���O���������܂��B adCmdTableDirect �I�v�V�������g�p���� Recordset ���J�����ꍇ�AadAsynchFetchNonBlocking �͖����ɂȂ�܂��B
    adExecuteNoRecords = 128     '�R�}���h �e�L�X�g���A�s��Ԃ��Ȃ��R�}���h�܂��̓X�g�A�h �v���V�[�W�� (���Ƃ��΁A�f�[�^�̑}���݂̂��s���R�}���h) �ł��邱�Ƃ������܂��B �擾�����s�������Ă��폜�����̂ŁA�R�}���h����͕Ԃ���܂���B adExecuteNoRecords�́A�R�}���h�܂���Connection��Execute���\�b�h�ɁA�ȗ��\�ȃp�����[�^�[�Ƃ��Ă̂ݓn�����Ƃ��ł��܂��B
    adExecuteStream = 1024       '�R�}���h�̎��s���ʂ��X�g���[���Ƃ��ĕԂ���邱�Ƃ������܂��B adExecuteStream�́A Command Execute���\�b�h�ɃI�v�V�����̃p�����[�^�[�Ƃ��ēn�����Ƃ��ł��܂��B
    adExecuteRecord = 512        'Indicates that the CommandText is a command or stored procedure that returns a single row which should be returned as a Record object.
    adOptionUnspecified = -1     'Indicates that the command is unspecified.
End Enum

'*-----------------------------------------------------------------------------
'* Field �I�u�W�F�N�g�� 1 �ȏ�̑�����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum FieldAttributeEnum
    adFldCacheDeferred = 4096     '�v���o�C�_�[�Ńt�B�[���h�l���L���b�V������A���̌�̓ǂݎ��̓L���b�V������s���邱�Ƃ������܂��B
    adFldFixed = 16               '�t�B�[���h���Œ蒷�f�[�^���܂ނ��Ƃ������܂��B
    adFldIsChapter = 8192         '�t�B�[���h���`���v�^�[�l���܂݁A���̐e�t�B�[���h�Ɋ֘A�t����ꂽ����̎q���R�[�h�Z�b�g���w�肵�Ă��邱�Ƃ������܂��B�ʏ�A�`���v�^�[ �t�B�[���h�̓f�[�^ �V�F�C�v��t�B���^�[�p�Ɏg���܂��B
    adFldIsCollection = 262144    '���R�[�h���������\�[�X���A�e�L�X�g �t�@�C���Ȃǂ̒P���ȃ��\�[�X�ł͂Ȃ��A�t�H���_�[�Ȃǂ̂悤�ɑ��̃��\�[�X�̃R���N�V�����ł��邱�Ƃ��A�t�B�[���h���\���Ă��邱�Ƃ������܂��B
    adFldIsDefaultStream = 131072 '�t�B�[���h���A���R�[�h���������\�[�X�̊���X�g���[�����܂ނ��Ƃ������܂��B ���Ƃ��΁A����̃X�g���[���́Aweb �T�C�g�̃��[�g�t�H���_�[�� HTML �R���e���c�ɂ��邱�Ƃ��ł��܂��B����́A���[�g URL ���w�肳�ꂽ�Ƃ��Ɏ����I�ɒ񋟂���܂��B
    adFldIsNullable = 32          '�t�B�[���h�� null �l���w��ł��邱�Ƃ������܂��B
    adFldIsRowURL = 65536         '�t�B�[���h���A���R�[�h�������f�[�^ �X�g�A�̃��\�[�X���w�肷�� URL ���܂ނ��Ƃ������܂��B
    adFldKeyColumn = 32768
    adFldLong = 128               '�t�B�[���h�������O �o�C�i���^�̃t�B�[���h�ł��邱�Ƃ������܂��B�܂��AAppendChunk ���\�b�h�� GetChunk ���\�b�h���g�p�ł��邱�Ƃ������܂��B
    adFldMayBeNull = 64           '�t�B�[���h����� null �l�̓ǂݎ�肪�\�ł��邱�Ƃ������܂��B
    adFldMayDefer = 2             '�t�B�[���h���x���t�B�[���h�ł��邱�Ƃ������܂��B�t�B�[���h�l�́A���R�[�h�S�̂̃f�[�^ �\�[�X����擾���ꂸ�A�����I�ɃA�N�Z�X�����ꍇ�̂ݎ擾����܂��B
    adFldNegativeScale = 16384    '���̃X�P�[���l���T�|�[�g�����̐��l���A�t�B�[���h���\���Ă��邱�Ƃ������܂��B�X�P�[���́ANumericScale �v���p�e�B�Ŏw�肵�܂��B
    adFldRowID = 256              '�t�B�[���h���������݋֎~�̉i�������ꂽ�s���ʎq���܂݁A�s�����ʂ������ (���R�[�h�ԍ��A��ӎ��ʎq�Ȃ�) �ȊO�ɗL���Ȓl�͎����Ȃ����Ƃ������܂��B
    adFldRowVersion = 512         '�t�B�[���h���X�V���L�^���邽�߂̎����܂��͓��t�X�^���v���܂ނ��Ƃ������܂��B
    adFldUnknownUpdatable = 8     '�t�B�[���h�ւ̏������݂��\���ǂ������v���o�C�_�[���m�F�ł��Ȃ����Ƃ������܂��B
    adFldUnspecified = -1         '�v���o�C�_�[���t�B�[���h�������w�肵�Ȃ����Ƃ������܂��B
    adFldUpdatable = 4            '�t�B�[���h�ւ̏������݂��\�ł��邱�Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Record �I�u�W�F�N�g�� Fields �R���N�V�����ŎQ�Ƃ�������̃t�B�[���h��
'* �\���܂��B
'*-----------------------------------------------------------------------------
Public Enum FieldEnum
    adDefaultStream = -1 'Record �Ɋ֘A�t����ꂽ����� Stream �I�u�W�F�N�g���܂ރt�B�[���h���Q�Ƃ��܂��B
    adRecordURL = -2     '���݂� Record �̐�� URL ��������܂ރt�B�[���h���Q�Ƃ��܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Field �I�u�W�F�N�g�̏�Ԃ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum FieldStatusEnum
    adFieldAlreadyExists = 26             '�w�肵���t�B�[���h�����ɑ��݂��邱�Ƃ������܂��B
    adFieldBadStatus = 12                 'ADO ���� OLE DB �v���o�C�_�[�ɖ����ȏ�Ԓl�����M���ꂽ���Ƃ������܂��B�����Ƃ��ẮAOLE DB 1.0 �v���o�C�_�[�܂��� 1.1 �v���o�C�_�[�A���邢�͕s�K�؂ȑg�ݍ��킹�� Value �� Status ���l�����܂��B
    adFieldCannotComplete = 20            'Source �Ŏw�肳�ꂽ URL �̃T�[�o�[������������ł��Ȃ��������Ƃ������܂��B
    adFieldCannotDeleteSource = 23        '�ړ�����ŁA�c���[�܂��̓T�u�c���[��V�����ʒu�Ɉړ��������\�[�X���폜�ł��Ȃ��������Ƃ������܂��B
    adFieldCantConvertValue = 2           '�t�B�[���h�̎擾�܂��͕ۑ����s���Ƃ��Ƀf�[�^�������Ă��܂����Ƃ������܂��B
    adFieldCantCreate = 7                 '�v���o�C�_�[�̌��x (���e�t�B�[���h���Ȃ�) �𒴂������߂Ƀt�B�[���h��ǉ��ł��Ȃ��������Ƃ������܂��B
    adFieldDataOverflow = 6               '�v���o�C�_�[����Ԃ��ꂽ�f�[�^���t�B�[���h�̃f�[�^�^���I�[�o�[�t���[�������Ƃ������܂��B
    adFieldDefault = 13                   '�f�[�^�̐ݒ莞�Ƀt�B�[���h�̊���l���g��ꂽ���Ƃ������܂��B
    adFieldDoesNotExist = 16              '�w�肵���t�B�[���h�����݂��Ȃ����Ƃ������܂��B
    adFieldIgnore = 15                    '�\�[�X�ł̃f�[�^�l�̐ݒ莞�ɂ��̃t�B�[���h���X�L�b�v���ꂽ���Ƃ������܂��B�v���o�C�_�[�Œl���ݒ肳��܂���ł����B
    adFieldIntegrityViolation = 10        '�v�Z�G���e�B�e�B�܂��͔h���G���e�B�e�B�ł��邽�߁A�t�B�[���h��ҏW�ł��Ȃ����Ƃ������܂��B
    adFieldInvalidURL = 17                '�f�[�^ �\�[�X URL �ɖ����ȕ��������邱�Ƃ������܂��B
    adFieldIsNull = 3                     '�v���o�C�_�[����� VT_NULL �̃o���A���g�^ (VARIANT) �̒l��Ԃ��A�t�B�[���h����łȂ����Ƃ������܂��B
    adFieldOK = 0                         '����l�B�t�B�[���h�̒ǉ��܂��͍폜������ɍs��ꂽ���Ƃ������܂��B
    adFieldOutOfSpace = 22                '�ړ��܂��̓R�s�[��������s���邽�߂ɕK�v�ȋL������v���o�C�_�[���m�ۂł��Ȃ����Ƃ������܂��B
    adFieldPendingChange = 262144         '�t�B�[���h���폜����A�قȂ�f�[�^�^���w�肵�čēx�ǉ����ꂽ���A�ȑO�ɏ�Ԃ� adFieldOK �ł������t�B�[���h�̒l���ύX���ꂽ���Ƃ������܂��BUpdate ���\�b�h�̌Ăяo����Ƀt�B�[���h�̍ŏI�`���ɂ���� Fields �R���N�V�������ύX����܂��B
    adFieldPendingDelete = 131072         'Delete ����ŏ�Ԃ��ݒ肳�ꂽ���Ƃ������܂��B�t�B�[���h�́AUpdate ���\�b�h�̌Ăяo����� Fields �R���N�V��������폜����悤�}�[�N����Ă��܂��B
    adFieldPendingInsert = 65536          'Append ����ŏ�Ԃ��ݒ肳�ꂽ���Ƃ������܂��BField �́AUpdate ���\�b�h�̌Ăяo����� Fields �R���N�V�����ɒǉ�����悤�}�[�N����Ă��܂��B
    adFieldPendingUnknown = 524288        '�t�B�[���h�̏�Ԃ�ݒ肷�錴���ƂȂ���������v���o�C�_�[�����ʂł��Ȃ����Ƃ������܂��B
    adFieldPendingUnknownDelete = 1048576 '�t�B�[���h�̏�Ԃ�ݒ肷�錴���ƂȂ���������v���o�C�_�[�����ʂł����AUpdate ���\�b�h�̌Ăяo����� Fields �R���N�V��������t�B�[���h���폜����邱�Ƃ������܂��B
    adFieldPermissionDenied = 9           '�ǂݎ���p�Ƃ��Ē�`����Ă��邽�߁A�t�B�[���h��ҏW�ł��Ȃ����Ƃ������܂��B
    adFieldReadOnly = 24                  '�f�[�^ �\�[�X���̃t�B�[���h���ǂݎ���p�Ƃ��Ē�`����Ă��邱�Ƃ������܂��B
    adFieldResourceExists = 19            '���� URL �ɃI�u�W�F�N�g�����ɑ��݂��A�㏑���ł��Ȃ����߁A�v���o�C�_�[����������s�ł��Ȃ��������Ƃ������܂��B
    adFieldResourceLocked = 18            '�f�[�^ �\�[�X�� 1 �ȏ�̑��̃A�v���P�[�V�����܂��̓v���Z�X�ɂ���ă��b�N����Ă��邽�߁A�v���o�C�_�[����������s�ł��Ȃ��������Ƃ������܂��B
    adFieldResourceOutOfScope = 25        '�\�[�X�܂��͈���� URL �����݂̃��R�[�h�͈̔͊O�ł��邱�Ƃ������܂��B
    adFieldSchemaViolation = 11           '�l���t�B�[���h�̃f�[�^ �\�[�X �X�L�[�}����Ɉᔽ���邱�Ƃ������܂��B
    adFieldSignMismatch = 5               '�v���o�C�_�[���Ԃ��f�[�^�l�������t���ŁAADO �t�B�[���h�l�̃f�[�^�^�������Ȃ��ł��邱�Ƃ������܂��B
    adFieldTruncated = 4                  '�f�[�^ �\�[�X����̓ǂݎ�莞�ɉϒ��f�[�^���؂�̂Ă�ꂽ���Ƃ������܂��B
    adFieldUnavailable = 8                '�f�[�^ �\�[�X����̓ǂݎ�莞�Ƀv���o�C�_�[���l�𔻕ʂł��Ȃ��������Ƃ������܂��B���Ƃ��΁A�s���쐬���ꂽ����ł��邱�ƁA��̊���l���g�p�s�ł��邱�ƁA�܂��͐V�����l���܂��w�肳��Ă��Ȃ����Ƃ������Ƃ��čl�����܂��B
    adFieldVolumeNotFound = 21            'URL �������L����{�����[�����v���o�C�_�[������ł��Ȃ����Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Recordset �Ńt�B���^�[�̑ΏۂƂȂ郌�R�[�h �O���[�v��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum FilterGroupEnum
    adFilterAffectedRecords = 2    '�Ō�ɍs���� Delete�AResync�AUpdateBatch�A�܂��� CancelBatch �Ăяo���ŉe�����󂯂����R�[�h�݂̂�\������悤�Ƀt�B���^�[�������܂��B
    adFilterConflictingRecords = 5 '�Ō�ɍs�����o�b�`�X�V�����s�������R�[�h��\������悤�Ƀt�B���^�[�������܂��B
    adFilterFetchedRecords = 3     '�f�[�^�x�[�X����Ō�Ɏ擾���ꂽ���R�[�h�ł��錻�݂̃L���b�V�����̃��R�[�h��\������悤�Ƀt�B���^�[�������܂��B
    adFilterNone = 0               '���݂̃t�B���^�[���폜���A���ׂẴ��R�[�h�𕜌����ĕ\�����܂��B
    adFilterPendingRecords = 1     '�ύX���s��ꂽ���A�ύX���e���T�[�o�[�ɂ܂����M����Ă��Ȃ����R�[�h�݂̂�\������悤�Ƀt�B���^�[�������܂��B�o�b�`�X�V���[�h�̏ꍇ�̂ݎg�p�ł��܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Recordset ����擾���郌�R�[�h����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum GetRowsOptionEnum
    adGetRowsRest = -1 '���݂̈ʒu�܂��� GetRows ���\�b�h�� Start �p�����[�^�[�Ŏw�肳�ꂽ�u�b�N�}�[�N����ARecordset ���̎c��̃��R�[�h���擾���܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Connection �I�u�W�F�N�g�̃g�����U�N�V�����������x����\���܂�
'*
'*-----------------------------------------------------------------------------
Public Enum IsolationLevelEnum
    adXactUnspecified = -1       '�v���o�C�_�[���w�肳�ꂽ���̂Ƃ͈قȂ镪�����x�����g�p���Ă��܂����A���x�������ł��Ȃ����Ƃ������܂��B
    adXactChaos = 16             '�����x�̍����g�����U�N�V��������ۗ̕����̕ύX���㏑���ł��Ȃ����Ƃ������܂��B
    adXactBrowse = 256           '1�̃g�����U�N�V��������A���̃g�����U�N�V�����̃R�~�b�g����Ă��Ȃ��ύX��\���ł��邱�Ƃ������܂��B
    adXactReadUncommitted = 256  'adXactBrowse�Ɠ����ł��B
    adXactCursorStability = 4096 '1�̃g�����U�N�V��������A�R�~�b�g���ꂽ��ɂ̂ݑ��̃g�����U�N�V�����̕ύX��\���ł��邱�Ƃ������܂��B
    adXactReadCommitted = 4096   'adXactCursorStability�Ɠ����ł��B
    adXactRepeatableRead = 65536 '1�̃g�����U�N�V��������A���̃g�����U�N�V�����ōs��ꂽ�ύX��\���ł��Ȃ����A�ăN�G���ŐV����Recordset�I�u�W�F�N�g���擾�ł��邱�Ƃ������܂��B
    adXactIsolated = 1048576     '�g�����U�N�V���������̃g�����U�N�V�����Ƃ͕������Ď��s����邱�Ƃ������܂��B
    adXactSerializable = 1048576 'adXactIsolated�Ɠ����ł��B
End Enum

'*-----------------------------------------------------------------------------
'* �e�L�X�g Stream �I�u�W�F�N�g�̍s��؂�L���Ɏg���Ă��镶����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum LineSeparatorEnum
    adCR = 13   '���A�������܂��B
    adCRLF = -1 '����l�B���A���s�������܂��B
    adLF = 10   '���s�������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* �ҏW���Ƀ��R�[�h�ɓK�p����郍�b�N�̎�ނ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum LockTypeEnum
    adLockBatchOptimistic = 4 '���L�I�o�b�`�X�V�������܂��B�o�b�`�X�V���[�h�̏ꍇ�ɕK�v�ł��B
    adLockOptimistic = 3      '���R�[�h�P�ʂ̋��L�I���b�N�������܂��BUpdate ���\�b�h���Ăяo�����ꍇ�ɂ̂݁A�v���o�C�_�[�͋��L�I���b�N���g���ă��R�[�h�����b�N���܂��B
    adLockPessimistic = 2     '���R�[�h�P�ʂ̔r���I���b�N�������܂��B�v���o�C�_�[�́A���R�[�h���m���ɕҏW���邽�߂̑[�u���s���܂��B�ʏ�́A�ҏW����Ƀf�[�^ �\�[�X�Ń��R�[�h�����b�N���܂��B
    adLockReadOnly = 1        '�ǂݎ���p�̃��R�[�h�������܂��B�f�[�^�̕ύX�͂ł��܂���B
    adLockUnspecified = -1    '���b�N�̎�ނ��w�肵�܂���B�����̏ꍇ�A�������Ɠ������b�N�̎�ނ��K�p����܂��B
End Enum

'*-----------------------------------------------------------------------------
'* �T�[�o�[�ɂǂ̃��R�[�h���Ԃ���邩��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum MarshalOptionsEnum
    adMarshalAll = 0          '����l�B���ׂĂ̍s���T�[�o�[�ɕԂ��܂��B
    adMarshalModifiedOnly = 1 '�ύX�����s�̂݃T�[�o�[�ɕԂ��܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Record �I�u�W�F�N�g�� MoveRecord ���\�b�h�̓����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum MoveRecordOptionsEnum
    adMoveUnspecified = -1    '����l�B����̈ړ���������s���܂��B�����̈���t�@�C���܂��̓f�B���N�g��������ꍇ�͑��삪���s���A�n�C�p�[�e�L�X�g �����N���X�V����܂��B
    adMoveOverWrite = 1       '�����̈���t�@�C���܂��̓f�B���N�g���������Ă��㏑�����܂��B
    adMoveDontUpdateLinks = 2 '�\�[�X Record �̃n�C�p�[�e�L�X�g �����N���X�V���Ȃ����ƂŁAMoveRecord ���\�b�h�̊��蓮���ύX���܂��B���蓮��̓v���o�C�_�[�̋@�\�ɂ���ĈقȂ�܂��B�v���o�C�_�[���T�|�[�g���Ă���΁A�ړ�����Ń����N���X�V����܂��B�v���o�C�_�[�������N�̏C�����T�|�[�g���Ă��Ȃ��ꍇ�A�܂��͂��̒l���w�肳��Ă��Ȃ��ꍇ�A�����N���C�����Ȃ��Ă��ړ��͐������܂��B
    adMoveAllowEmulation = 4  '�v���o�C�_�[�ɂ��ړ� (�_�E�����[�h�A�A�b�v���[�h�A�폜�̑�����g�p) �̃V�~�����[�V������v�����܂��B���� URL ���\�[�X�Ƃ͕ʂ̃T�[�o�[�ɂ�������A�ʂ̃v���o�C�_�[���T�[�r�X��񋟂��Ă��邽�߂� Record �̈ړ������s����ƁA�v���o�C�_�[�ԂŃ��\�[�X���ړ�����Ƃ��̃v���o�C�_�[�̋@�\�̈Ⴂ�ɂ��A�x�����Ԃ̑�����f�[�^�̑������N���邱�Ƃ�����܂��B
End Enum

'*-----------------------------------------------------------------------------
'* �I�u�W�F�N�g���J���Ă��邩���Ă��邩�A�f�[�^ �\�[�X�ɐڑ������A�R�}���h��
'* ���s�����A�܂��̓f�[�^���擾������\���܂��B
'*-----------------------------------------------------------------------------
Public Enum ObjectStateEnum
    adStateClosed = 0     '�I�u�W�F�N�g�����Ă��邱�Ƃ������܂��B
    adStateOpen = 1       '�I�u�W�F�N�g���J���Ă��邱�Ƃ������܂��B
    adStateConnecting = 2 '�I�u�W�F�N�g���ڑ����ł��邱�Ƃ������܂��B
    adStateExecuting = 4  '�I�u�W�F�N�g���R�}���h�����s���ł��邱�Ƃ������܂��B
    adStateFetching = 8   '�I�u�W�F�N�g�̍s���擾���ł��邱�Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* �����܂��͏��L����ݒ肷��f�[�^�x�[�X �I�u�W�F�N�g�̎�ނ������܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum ObjectTypeEnum
    adPermObjColumn = 2             '�I�u�W�F�N�g�͗�ł��B
    adPermObjDatabase = 3           '�I�u�W�F�N�g�̓f�[�^�x�[�X�ł��B
    adPermObjProcedure = 4          '�I�u�W�F�N�g�̓v���V�[�W���ł��B
    adPermObjProviderSpecific = -1  '�I�u�W�F�N�g�̎�ނ́A�v���o�C�_�[ �ɂ���Ē�`����܂��BObjectType �p�����[�^�[�� adPermObjProviderSpecific �ŁAObjectTypeId ���w�肳��Ă��Ȃ��ꍇ�A�G���[���������܂��B
    adPermObjTable = 1              '�I�u�W�F�N�g�̓e�[�u���ł��B
    adPermObjView = 5               '�I�u�W�F�N�g�̓r���[�ł��B
End Enum


'*-----------------------------------------------------------------------------
'* Parameter �I�u�W�F�N�g�̑�����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum ParameterAttributesEnum
    adParamLong = 128    '�p�����[�^�[�Ƀ����O �o�C�i���^�̃f�[�^���w��ł��邱�Ƃ������܂��B
    adParamNullable = 64 '�p�����[�^�[�� null �l���w��ł��邱�Ƃ������܂��B
    adParamSigned = 16   '�p�����[�^�[�ɕ����t���̒l���w��ł��邱�Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Parameter �����̓p�����[�^�[�Əo�̓p�����[�^�[�̂����ꂩ�A�܂��͂��̗�����
'* �\���̂��A���邢�̓X�g�A�h �v���V�[�W������̖߂�l�ł��邩��\���܂��B
'*-----------------------------------------------------------------------------
Public Enum ParameterDirectionEnum
    adParamInput = 1       '����l�B�p�����[�^�[�����̓p�����[�^�[��\�����Ƃ������܂��B
    adParamInputOutput = 3 '�p�����[�^�[�����̓p�����[�^�[�Əo�̓p�����[�^�[�̗�����\�����Ƃ������܂��B
    adParamOutput = 2      '�p�����[�^�[���o�̓p�����[�^�[��\�����Ƃ������܂��B
    adParamReturnValue = 4 '�p�����[�^�[���߂�l��\�����Ƃ������܂��B
    adParamUnknown = 0     '�p�����[�^�[�̕������s���ł��邱�Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Recordset ��ۑ�����Ƃ��̌`����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum PersistFormatEnum
    adPersistADTG = 0             'Microsoft Advanced Data TableGram (ADTG) �`���ł��邱�Ƃ������܂��B
    adPersistXML = 1              '�g���}�[�N�A�b�v���� (XML) �`���ł��邱�Ƃ������܂��B
    adPersistADO = 1              'Indicates that ADO's own Extensible Markup Language (XML) format will be used. This value is the same as adPersistXML and is included for backwards compatibility.
    adPersistProviderSpecific = 2 'Indicates that the provider will persist the Recordset using its own format.
End Enum

'*-----------------------------------------------------------------------------
'* Recordset ���̃��R�[�h �|�C���^�[�̌��݂̈ʒu��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum PositionEnum
    adPosBOF = -2     '���݂̃��R�[�h �|�C���^�[�� BOF �ɂ��邱�Ƃ������܂� (BOF �v���p�e�B�� True �ł�)�B
    adPosEOF = -3     '���݂̃��R�[�h �|�C���^�[�� EOF �ɂ��邱�Ƃ������܂� (EOF �v���p�e�B�� True �ł�)�B
    adPosUnknown = -1 'Recordset ����ł��邩�A���݂̈ʒu���s�����A�܂��̓v���o�C�_�[�� AbsolutePage �v���p�e�B�܂��� AbsolutePosition �v���p�e�B���T�|�[�g���Ă��Ȃ����Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Property �I�u�W�F�N�g�̑�����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum PropertyAttributesEnum
    adPropNotSupported = 0 '�v���o�C�_�[���v���p�e�B���T�|�[�g���Ă��Ȃ����Ƃ������܂��B
    adPropRequired = 1     '�f�[�^ �\�[�X������������ɂ́A���[�U�[�����̃v���p�e�B�l���w�肷��K�v�����邱�Ƃ������܂��B
    adPropOptional = 2     '���[�U�[�����̃v���p�e�B�l���w�肵�Ȃ��Ă��f�[�^ �\�[�X���������ł��邱�Ƃ�\���܂��B
    adPropRead = 512       '���[�U�[���v���p�e�B��ǂݎ��\�ł��邱�Ƃ������܂��B
    adPropWrite = 1024     '���[�U�[���v���p�e�B��ݒ�ł��邱�Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Record �I�u�W�F�N�g�� Open ���\�b�h�ɑ΂��A������ Record ���J�����A�V����
'* Record ���쐬���邩��\���܂��B�����̒l�� AND ���Z�q�Ō����ł��܂��B
'*-----------------------------------------------------------------------------
Public Enum RecordCreateOptionsEnum
    adCreateCollection = 8192        '������ Record ���J�����ɁASource �p�����[�^�[�Ŏw�肵���m�[�h�ɐV���� Record ���쐬���܂��B�\�[�X�������̃m�[�h���w�肵�Ă���ꍇ�AadCreateCollection �� adOpenIfExists �܂��� adCreateOverwrite �Ƒg�ݍ��킹�Ďg�p����Ă��Ȃ�����A���s���G���[�ɂȂ�܂��B
    adCreateNonCollection = 0        '��ނ� adSimpleRecord �̐V���� Record ���쐬���܂��B
    adCreateOverwrite = 67108864     '�쐬�t���O adCreateCollection�AadCreateNonCollection�A����� adCreateStructDoc ���C�����܂��B���̒l�ƍ쐬�t���O�̒l�� 1 �� OR ���g���ĘA������Ă���ꍇ�A�\�[�X URL �������̃m�[�h�܂��� Record ���w�肵�Ă���ƁA�����̂��̂��㏑������A�V���� Record ���쐬����܂��B���̒l�́AadOpenIfExists �Ƃ͕��p�ł��܂���B
    adCreateStructDoc = -2147483648# '������ Record ���J�����ɁA��ނ� adStructDoc �̐V���� Record ���쐬���܂��B
    adFailIfNotExists = -1           '����l�BSource �����݂��Ȃ��m�[�h���w�肵�Ă���ƁA���s���G���[�ɂȂ�܂��B
    adOpenIfExists = 33554432        '�쐬�t���O adCreateCollection�AadCreateNonCollection�A����� adCreateStructDoc ���C�����܂��B���̒l�ƍ쐬�t���O�̒l�� 1 �� OR ���g���ĘA������Ă���ꍇ�A�\�[�X URL �������̃m�[�h�܂��� Record �I�u�W�F�N�g���w�肵�Ă���ƁA�v���o�C�_�[�́A�V���� Record ���쐬�����ɁA�����̂��̂��J���K�v������܂��B���̒l�́AadCreateOverwrite �Ƃ͕��p�ł��܂���B
End Enum

'*-----------------------------------------------------------------------------
'* Record ���J���Ƃ��̃I�v�V������\���܂��B �����̒l�� OR ���Z�q�Ō����ł��܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum RecordOpenOptionsEnum
    adDelayFetchFields = 32768   '�v���o�C�_�[�ɑ΂��āARecord �Ɋ֘A�t����ꂽ�t�B�[���h�́A�����͎擾����K�v���Ȃ��A�t�B�[���h�ւ̍ŏ��̃A�N�Z�X���Ɏ擾�ł��邱�Ƃ������܂��B���̃t���O���w�肳��Ă��Ȃ��ꍇ�̊��蓮��ł́ARecord �I�u�W�F�N�g�̂��ׂẴt�B�[���h���擾����܂��B
    adDelayFetchStream = 16384   '�v���o�C�_�[�ɑ΂��āARecord �Ɋ֘A�t����ꂽ����X�g���[���𓖏��͎擾����K�v���Ȃ����Ƃ������܂��B���̃t���O���w�肳��Ă��Ȃ��ꍇ�̊��蓮��ł́ARecord �I�u�W�F�N�g�Ɋ֘A�t����ꂽ����X�g���[�����擾����܂��B
    adOpenAsync = 4096           'Record �I�u�W�F�N�g���񓯊����[�h�ŊJ����邱�Ƃ������܂��B
    adOpenExecuteCommand = 65536 'Source ������ɁA���s�����R�}���h �e�L�X�g���܂܂�邱�Ƃ������܂��B���̒l�́ARecordset.Open �� adCmdText �I�v�V�����Ɠ����ł��B
    adOpenOutput = 8388608       '���s�\�X�N���v�g (�g���q�� .ASP �̃y�[�W�Ȃ�) ������m�[�h���\�[�X���w�肵�Ă���ꍇ�A���s�����X�N���v�g�̌��ʂ��A�J���Ă��� Record �Ɋ܂܂�邱�Ƃ������܂��B���̒l�́A�R���N�V�����̂Ȃ����R�[�h�ɂ̂ݗL���ł��B
    adOpenRecordUnspecified = -1 '����l�B�I�v�V�������w�肳��Ă��Ȃ����Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* �o�b�`�X�V�܂��͂��̑��̈ꊇ����Ɋւ��郌�R�[�h�̏�Ԃ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum RecordStatusEnum
    adRecCanceled = 256              '���삪�������ꂽ���߁A���R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecCantRelease = 1024          '�����̃��R�[�h�����b�N����Ă������߁A�V�������R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecConcurrencyViolation = 2048 '�I�v�e�B�~�X�e�B�b�N�������s���䂪�g�p����Ă������߁A���R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecDBDeleted = 262144          '���R�[�h�͊��Ƀf�[�^ �\�[�X����폜����Ă��邱�Ƃ������܂��B
    adRecDeleted = 4                 '���R�[�h���폜���ꂽ���Ƃ������܂��B
    adRecIntegrityViolation = 4096   '���[�U�[������������Ɉᔽ�������߁A���R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecInvalid = 16                '�u�b�N�}�[�N�������Ȃ��߁A���R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecMaxChangesExceeded = 8192   '�ۗ����̕ύX�������������߁A���R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecModified = 2                '���R�[�h���ύX���ꂽ���Ƃ������܂��B
    adRecMultipleChanges = 64        '�����̃��R�[�h�ɉe�����y�Ԃ��߁A���R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecNew = 1                     '���R�[�h���V�������Ƃ������܂��B
    adRecObjectOpen = 16384          '�J���Ă���X�g���[�W �I�u�W�F�N�g�Ƃ̋����̂��߁A���R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecOK = 0                      '���R�[�h������ɍX�V���ꂽ���Ƃ������܂��B
    adRecOutOfMemory = 32768         '�������s���̂��߂Ƀ��R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecPendingChanges = 128        '�ۗ����̑}�����Q�Ƃ��Ă��邽�߁A���R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecPermissionDenied = 65536    '���[�U�[�̌����s���ɂ��A���R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecSchemaViolation = 131072    '��ɂȂ�f�[�^�x�[�X�̍\���Ɉᔽ����̂ŁA���R�[�h���ۑ�����Ȃ��������Ƃ������܂��B
    adRecUnmodified = 8              '���R�[�h���ύX����Ȃ��������Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Record �I�u�W�F�N�g�̎�ނ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum RecordTypeEnum
    adCollectionRecord = 1 '"�R���N�V����" ���R�[�h (�q�m�[�h�����郌�R�[�h) �������܂��B
    adSimpleRecord = 0     '"�P��" ���R�[�h (�q�m�[�h���Ȃ����R�[�h) �������܂��B
    adStructDoc = 2        'COM �\�����h�L�������g��\������� "�R���N�V����" ���R�[�h�������܂��B
    adRecordUnknown = -1   'Indicates that the type of this Record is unknown.
End Enum

'*-----------------------------------------------------------------------------
'* Resync �̌Ăяo���ɂ���Ċ�ɂȂ�l���㏑������邩�ǂ�����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum ResyncEnum
    adResyncAllValues = 2        '����l�B�f�[�^�͏㏑������A�ۗ����̍X�V�͎�������܂��B
    adResyncUnderlyingValues = 1 '�f�[�^�͏㏑�����ꂸ�A�ۗ����̍X�V�͎�������܂���B
End Enum

'*-----------------------------------------------------------------------------
'* Stream �I�u�W�F�N�g����t�@�C���ɕۑ�����Ƃ��ɁA�t�@�C�����쐬���邩�A
'* �㏑�����邩��\���܂��B�����̒l�� AND ���Z�q�Ō����ł��܂��B
'*-----------------------------------------------------------------------------
Public Enum SaveOptionsEnum
    adSaveCreateNotExist = 1  '����l�BFileName �p�����[�^�[�Ŏw�肵���t�@�C�����Ȃ��ꍇ�͐V�����t�@�C�����쐬����܂��B
    adSaveCreateOverWrite = 2 'Filename �p�����[�^�[�Ŏw�肵���t�@�C��������ꍇ�́A���݊J����Ă��� Stream �I�u�W�F�N�g�̃f�[�^�Ńt�@�C�����㏑������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* OpenSchema ���\�b�h���擾����X�L�[�} Recordset �̎�ނ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum SchemaEnum
    adSchemaActions = 41               '
    adSchemaAsserts = 0                '�J�^���O�ɒ�`����A����̃��[�U�[�����L����A�T�[�V������Ԃ��܂��B (ASSERTIONS �s�Z�b�g)
    adSchemaCatalogs = 1               'DBMS ����A�N�Z�X�ł���J�^���O�Ɋ֘A�t�����Ă��镨���I������Ԃ��܂��B (CATALOGS �s�Z�b�g)
    adSchemaCharacterSets = 2          '�J�^���O�ɒ�`����A����̃��[�U�[���A�N�Z�X�ł��镶���Z�b�g��Ԃ��܂��B (CHARACTER_SETS �s�Z�b�g)
    adSchemaCheckConstraints = 5       '�J�^���O�ɒ�`����A����̃��[�U�[�����L���� CHECK �����Ԃ��܂��B (CHECK_CONSTRAINTS �s�Z�b�g)
    adSchemaCollations = 3             '�J�^���O�ɒ�`����A����̃��[�U�[���A�N�Z�X�ł��镶���ƍ�������Ԃ��܂��B (COLLATIONS �s�Z�b�g)
    adSchemaColumnPrivileges = 13      '�J�^���O�ɒ�`����A����̃��[�U�[�����p�ł���A�܂��͌��������e�[�u���̗�ɑ΂��������Ԃ��܂��B (COLUMN_PRIVILEGES �s�Z�b�g)
    adSchemaColumns = 4                '�J�^���O�ɒ�`����A����̃��[�U�[���A�N�Z�X�ł���e�[�u���̗� (�r���[���܂�) ��Ԃ��܂��B (COLUMNS �s�Z�b�g)
    adSchemaColumnsDomainUsage = 11    '�J�^���O�ɒ�`����A���̃J�^���O�ɒ�`���ꂽ�h���C���Ɉˑ����A����̃��[�U�[�����L������Ԃ��܂��B (COLUMN_DOMAIN_USAGE �s�Z�b�g)
    adSchemaCommands = 42              '
    adSchemaConstraintColumnUsage = 6  '�J�^���O�ɒ�`����A����̃��[�U�[�����L���A�Q�Ɛ���A��Ӑ���ACHECK ����A����уA�T�[�V�����Ɏg�����Ԃ��܂��B (CONSTRAINT_COLUMN_USAGE �s�Z�b�g)
    adSchemaConstraintTableUsage = 7   '�J�^���O�ɒ�`����A����̃��[�U�[�����L���A�Q�Ɛ���A��Ӑ���ACHECK ����A����уA�T�[�V�����Ɏg���e�[�u����Ԃ��܂��B (CONSTRAINT_TABLE_USAGE �s�Z�b�g)
    adSchemaCubes = 32                 '�X�L�[�} (�v���o�C�_�[���X�L�[�}���T�|�[�g���Ă��Ȃ��ꍇ�̓J�^���O) ���̗��p�ł���L���[�u�Ɋւ������Ԃ��܂��B (CUBES �s�Z�b�g \*)
    adSchemaDBInfoKeywords = 30        '�v���o�C�_�[�ŗL�̃L�[���[�h�̈ꗗ��Ԃ��܂��B (IDBInfo:: getkeywords \*)
    adSchemaDBInfoLiterals = 31        '�e�L�X�g �R�}���h�Ŏg���A�v���o�C�_�[�ŗL�̃��e�����̈ꗗ��Ԃ��܂��B (IDBInfo:: GetLiteralInfo \*)
    adSchemaDimensions = 33            '����̃L���[�u�̎����Ɋւ������Ԃ��܂��B �������Ƃ� 1 �s�����蓖�Ă��܂��B (DIMENSIONS �s�Z�b�g \*)
    adSchemaForeignKeys = 27           '����̃��[�U�[���J�^���O�ɒ�`�����O���L�[���Ԃ��܂��B (FOREIGN_KEYS �s�Z�b�g)
    adSchemaFunctions = 40             '
    adSchemaHierarchies = 34           '�����ŗ��p�ł���K�w�Ɋւ������Ԃ��܂��B (HIERARCHIES �s�Z�b�g \*)
    adSchemaIndexes = 12               '�J�^���O�ɒ�`����A����̃��[�U�[�����L����C���f�b�N�X��Ԃ��܂��B (INDEXES �s�Z�b�g)
    adSchemaKeyColumnUsage = 8         '�J�^���O�ɒ�`����A����̃��[�U�[���L�[�Ƃ��Đ��񂵂����Ԃ��܂��B (KEY_COLUMN_USAGE �s�Z�b�g)
    adSchemaLevels = 35                '�����ŗ��p�ł��郌�x���Ɋւ������Ԃ��܂��B (LEVELS �s�Z�b�g \*)
    adSchemaMeasures = 36              '���p�ł���P�ʂɊւ������Ԃ��܂��B (MEASURES �s�Z�b�g \*)
    adSchemaMembers = 38               '���p�ł��郁���o�[�Ɋւ������Ԃ��܂��B (MEMBERS �s�Z�b�g \*)
    adSchemaPrimaryKeys = 28           '����̃��[�U�[���J�^���O�ɒ�`������L�[���Ԃ��܂��B (PRIMARY_KEYS �s�Z�b�g)
    adSchemaProcedureColumns = 29      '�v���V�[�W�����Ԃ��s�Z�b�g�̗�Ɋւ������Ԃ��܂��B (PROCEDURE_COLUMNS Rowset)
    adSchemaProcedureParameters = 26   '�v���V�[�W���̃p�����[�^�[�ƃ��^�[�� �R�[�h�Ɋւ������Ԃ��܂��B (PROCEDURE_PARAMETERS �s�Z�b�g)
    adSchemaProcedures = 16            '�J�^���O�ɒ�`����A����̃��[�U�[�����L����v���V�[�W����Ԃ��܂��B (PROCEDURES �s�Z�b�g)
    adSchemaProperties = 37            '�����̊e���x���ŗ��p�ł���v���p�e�B�Ɋւ������Ԃ��܂��B (PROPERTIES �s�Z�b�g \*)
    adSchemaProviderSpecific = -1      '�v���o�C�_�[����W���̐�p�̃X�L�[�} �N�G�����`����ꍇ�Ɏg���܂��B
    adSchemaProviderTypes = 22         '�f�[�^ �v���o�C�_�[���T�|�[�g���� (��{) �f�[�^�^��Ԃ��܂��B (PROVIDER_TYPES �s�Z�b�g)
    adSchemaReferentialConstraints = 9 '�J�^���O�ɒ�`����A����̃��[�U�[�����L����Q�Ɛ����Ԃ��܂��B (REFERENTIAL_CONSTRAINTS �s�Z�b�g)
    adSchemaSchemata = 17              '����̃��[�U�[�����L����X�L�[�} (�f�[�^�x�[�X �I�u�W�F�N�g) ��Ԃ��܂��B (SCHEMATA �s�Z�b�g)
    adSchemaSets = 43                  '
    adSchemaSQLLanguages = 18          '�J�^���O�ɒ�`���ꂽ SQL ���������f�[�^���T�|�[�g���鏀�����x���A�I�v�V�����A����ь����Ԃ��܂��B (SQL_LANGUAGES �s�Z�b�g)
    adSchemaStatistics = 19            '�J�^���O�ɒ�`����A����̃��[�U�[�����L���铝�v�l��Ԃ��܂��B (STATISTICS �s�Z�b�g)
    adSchemaTableConstraints = 10      '�J�^���O�ɒ�`����A����̃��[�U�[�����L����e�[�u�������Ԃ��܂��B (TABLE_CONSTRAINTS �s�Z�b�g)
    adSchemaTablePrivileges = 14       '�J�^���O�ɒ�`����A����̃��[�U�[�����p�ł���A�܂��͌��������e�[�u���ɑ΂��������Ԃ��܂��B (TABLE_PRIVILEGES �s�Z�b�g)
    adSchemaTables = 20                '�J�^���O�ɒ�`����A����̃��[�U�[���A�N�Z�X�ł���e�[�u�� (�r���[���܂�) ��Ԃ��܂��B (TABLES �s�Z�b�g)
    adSchemaTranslations = 21          '�J�^���O�ɒ�`����A����̃��[�U�[���A�N�Z�X�ł��镶���ϊ���Ԃ��܂��B (TRANSLATIONS �s�Z�b�g)
    adSchemaTrustees = 39              '�����g�p���邽�߂ɗ\�񂳂�Ă��܂��B
    adSchemaUsagePrivileges = 15       '�J�^���O�ɒ�`����A����̃��[�U�[�����p�ł���A�܂��͌��������I�u�W�F�N�g�ɑ΂��� USAGE ������Ԃ��܂��B (USAGE_PRIVILEGES �s�Z�b�g)
    adSchemaViewColumnUsage = 24       '�J�^���O�ɒ�`����A����̃��[�U�[�����L����A�\���e�[�u�����ˑ�������Ԃ��܂��B (VIEW_COLUMN_USAGE �s�Z�b�g)
    adSchemaViews = 23                 '�J�^���O�ɒ�`����A����̃��[�U�[���A�N�Z�X�ł���r���[��Ԃ��܂��B (VIEWS �s�Z�b�g)
    adSchemaViewTableUsage = 25        '�J�^���O�ɒ�`����A����̃��[�U�[�����L���A�\���e�[�u�����ˑ�����e�[�u����Ԃ��܂��B (VIEW_TABLE_USAGE �s�Z�b�g)
End Enum

'*-----------------------------------------------------------------------------
'* Recordset ���̃��R�[�h�̌���������\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum SearchDirectionEnum
    adSearchBackward = -1 '������������ARecordset �̐擪�ŏI�����܂��B��v���郌�R�[�h��������Ȃ��ꍇ�A���R�[�h �|�C���^�[�� BOF �Ɉړ����܂��B
    adSearchForward = 1   '�O�����������ARecordset �̖����ŏI�����܂��B��v���郌�R�[�h��������Ȃ��ꍇ�A���R�[�h �|�C���^�[�� EOF �Ɉړ����܂��B
End Enum

'*-----------------------------------------------------------------------------
'* �B�����ځFRecordset ���̃��R�[�h�̌���������\���܂��B
'* �R�����g�A�E�g�B
'*-----------------------------------------------------------------------------
'Public Enum SearchDirection
'    adSearchBackward = -1
'    adSearchForward = 1
'End Enum

'*-----------------------------------------------------------------------------
'* ���s���� Seek �̎�ނ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum SeekEnum
    adSeekFirstEQ = 1   'KeyValues �ƈ�v����ŏ��̃L�[���������܂��B
    adSeekLastEQ = 2    'KeyValues �ƈ�v����Ō�̃L�[���������܂��B
    adSeekAfterEQ = 4   'KeyValues �ƈ�v����L�[�A�܂��͂��̒���̃L�[�̂����ꂩ���������܂��B
    adSeekAfter = 8     'KeyValues �ƈ�v����L�[�̒���̃L�[���������܂��B
    adSeekBeforeEQ = 16 'KeyValues �ƈ�v����L�[�A�܂��͂��̒��O�̃L�[�̂����ꂩ���������܂��B
    adSeekBefore = 32   'KeyValues �ƈ�v����L�[�̒��O�̃L�[���������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Stream �I�u�W�F�N�g���J���Ƃ��̃I�v�V������\���܂��B
'* �����̒l�� OR ���Z�q�Ō����ł��܂��B
'*-----------------------------------------------------------------------------
Public Enum StreamOpenOptionsEnum
    adOpenStreamAsync = 1        '�񓯊����[�h�� Stream �I�u�W�F�N�g���J���܂��B
    adOpenStreamFromRecord = 4   'Source �p�����[�^�[�̓��e���A���ɊJ����Ă��� Record �I�u�W�F�N�g�Ƃ��Ď��ʂ��܂��B���蓮��ł́ASource �́A�c���[�\���̃m�[�h�𒼐ڎw�肷�� URL �Ƃ��ď������܂��B���̃m�[�h�Ɋ֘A�t����ꂽ����X�g���[�����J����܂��B
    adOpenStreamUnspecified = -1 '����l�B����̃I�v�V������ Stream �I�u�W�F�N�g���J�����Ƃ�\���܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Stream �I�u�W�F�N�g����A�X�g���[���S�̂�ǂݎ�邩�A�܂��͎��̍s��ǂݎ�邩��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum StreamReadEnum
    adReadAll = -1  '����l�B���݂̈ʒu���� EOS �}�[�J�[�����ɁA���ׂẴo�C�g���X�g���[������ǂݎ��܂��B����́A�o�C�i�� �X�g���[�� (Type �� adTypeBinary) �ɗB��L���� StreamReadEnum �l�ł��B
    adReadLine = -2 '�X�g���[�����玟�̍s��ǂݎ��܂� (LineSeparator �v���p�e�B�Ŏw��)�B
End Enum

'*-----------------------------------------------------------------------------
'* Stream �I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum StreamTypeEnum
    adTypeBinary = 1 '�o�C�i�� �f�[�^�������܂��B
    adTypeText = 2   '����l�BCharset �Ŏw�肳�ꂽ�����Z�b�g�̃e�L�X�g �f�[�^�������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Stream �I�u�W�F�N�g�ɏ������ޕ�����ɁA�s��؂�L����ǉ����邩�ǂ�����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum StreamWriteEnum
    adWriteChar = 0 '����l�BStream �I�u�W�F�N�g�ɁAData �p�����[�^�[�Ŏw�肵���e�L�X�g��������������݂܂��B
    adWriteLine = 1 'Stream �I�u�W�F�N�g�ɁA�e�L�X�g������ƍs��؂�L�����������݂܂��BLineSeparator �v���p�e�B����`����Ă��Ȃ��ꍇ�́A���s���G���[��Ԃ��܂��B
End Enum

'*-----------------------------------------------------------------------------
'* ������Ƃ��� Recordset ���擾����Ƃ��̌`����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum StringFormatEnum
    adClipString = 2 '�s�� RowDelimiter �ɂ���āA�� ColumnDelimiter �ɂ���āAnull �l�� NullExpr �ɂ���ċ�؂��܂��BGetString ���\�b�h�̂����� 3 �̃p�����[�^�[�́AadClipString �� StringFormat �Ƃ̂ݕ��p�ł��܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Connection �I�u�W�F�N�g�̃g�����U�N�V����������\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum XactAttributeEnum
    adXactAbortRetaining = 262144  '���f�̕ێ������s���܂��B�܂�A RollbackTrans���Ăяo���ƁA�V�����g�����U�N�V�����������I�ɊJ�n����܂��B ���̐ݒ���T�|�[�g���Ă��Ȃ��v���o�C�_�[������܂��B
    adXactCommitRetaining = 131072 '�ێ��R�~�b�g�����s���܂��B�܂�A CommitTrans���Ăяo���ƁA�V�����g�����U�N�V�����������I�ɊJ�n����܂��B ���̐ݒ���T�|�[�g���Ă��Ȃ��v���o�C�_�[������܂��B
End Enum

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
'******************************************************************************
'* [�T  �v] �t�@�C���G���R�[�h�ꊇ�ϊ������B
'* [��  ��] �w�肵���t�H���_���̃t�@�C���̃G���R�[�h���ꊇ�ϊ�����B
'*
'* @param targetFolderName �ΏۂƂȂ�t�H���_�̃t���p�X
'* @param srcEncode �ύX���G���R�[�h
'* @param destEncode �ύX��G���R�[�h
'* @param bomInclude BOM�L���i�ȗ��B�K���False:BOM���j
'******************************************************************************
Public Sub ChangeFilesEncode(targetFolderName As String, srcEncode As String, destEncode As String, _
                            Optional bomInclude As Boolean = False)
    Dim fso, oFolder, oFiles, oFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = fso.GetFolder(targetFolderName)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Dim filePath As String
        filePath = fso.GetAbsolutePathName(AddPathSeparator(targetFolderName) & oFile.Name)
         
        ' �t�@�C���G���R�[�h�ϊ�
        Call ChangeFileEncode(filePath, srcEncode, destEncode, bomInclude)
    Next
End Sub

'******************************************************************************
'* [�T  �v] �t�@�C���G���R�[�h�ϊ������B
'* [��  ��] �w�肵���t�@�C���̃G���R�[�h��ϊ�����B
'*
'* @param filePath �ΏۂƂȂ�t�@�C���̃t���p�X
'* @param srcEncode �ύX���G���R�[�h
'* @param destEncode �ύX��G���R�[�h
'* @param bomInclude BOM�L���i�ȗ��B�K���False:BOM���j
'******************************************************************************
Public Sub ChangeFileEncode(filePath As String, srcEncode As String, destEncode As String, _
                            Optional bomInclude As Boolean = False)
    Dim adoStream1 As ADODBExStream, adoStream2 As ADODBExStream
    Set adoStream1 = New ADODBExStream
    Set adoStream2 = New ADODBExStream
         
    ' �ύX���t�@�C��Stream�Ǎ�
    With adoStream1
        .OpenStream
        .Type_ = adTypeText
        .CharSet = srcEncode
        .LoadFromFile filePath
    End With
     
    ' �ύX��t�@�C��Stream�Ǎ�
    With adoStream2
        .OpenStream
        .Type_ = adTypeText
        .CharSet = destEncode
        .BOM = bomInclude
    End With
     
    ' �G���R�[�h�ϊ�
    adoStream1.CopyTo adoStream2
    adoStream2.SaveToFile filePath, adSaveCreateOverWrite '�t�@�C���㏑�w��
     
    ' Stream�N���[�Y
    adoStream2.CloseStream
    adoStream1.CloseStream
End Sub

'******************************************************************************
'* [�T  �v] �t�@�C���Ǎ��E�������ݏ����B
'* [��  ��] �w�肵���Ǎ��t�@�C���̃f�[�^��ʃt�@�C���ɏ������ށB
'* [�Q  �l] ��e�ʃf�[�^�̓ǂݍ��݂ɂ��Ă͈ȉ��̃T�C�g���Q�l�ɂ����B
'*          <https://mussyu1204.myhome.cx/wordpress/it/?p=720>
'*
'* @param srcFilePath �Ǎ��t�@�C���̃t���p�X
'* @param srcEncode �Ǎ����G���R�[�h
'* @param srcSep �Ǎ������s�R�[�h
'* @param destFilePath �����t�@�C���̃t���p�X
'* @param destEncode ������G���R�[�h
'* @param destSep ��������s�R�[�h
'* @param funcName �s�ҏW�����p�֐����B
'*                 �ȉ��̂悤�Ɉ����ɕ�����A�߂�l�ɕ������Ԃ��֐������w��B
'*                 funcName(row As String) As String
'*                 �w�肵�Ȃ��i�󕶎��j�ꍇ�́A�s�ҏW�͍s��Ȃ��B
'* @param chunkSize �`�����N�T�C�Y�B���̃T�C�Y�𒴂���Ǎ��f�[�^�̏ꍇ�́A
'*                  �`�����N�T�C�Y���Ƃɕ������ď������s���B
'* @param bomInclude BOM�L���i�ȗ��B�K���False:BOM���j
'******************************************************************************
Public Sub ReadAndWrite(srcFilePath As String, srcEncode As String, srcSep As LineSeparatorEnum, _
                        destFilePath As String, destEncode As String, destSep As LineSeparatorEnum, _
                        Optional funcName As String = "", _
                        Optional chunkSize As Long = 2048, _
                        Optional bomInclude As Boolean = False)
                            
    Dim inStream As ADODBExStream, outStream As ADODBExStream
    Set inStream = New ADODBExStream
    Set outStream = New ADODBExStream

    With inStream
        .CharSet = srcEncode
        .LineSeparator = srcSep
        .OpenStream
        .LoadFromFile srcFilePath
    End With
     
    With outStream
        .CharSet = destEncode
        .BOM = bomInclude
        .LineSeparator = destSep
        .OpenStream
    End With

    Dim lines As Variant, lastLine As String
    
    ' �Ǎ��f�[�^�̃T�C�Y���w��T�C�Y���傫���ꍇ�͕��������i�������j���{
    If inStream.Size > chunkSize Then
        Do Until inStream.EOS
            Dim tmp As String: tmp = inStream.ReadText(chunkSize)
            lines = Split(tmp, vbLf)
             
            Dim lineCnt As Long: lineCnt = UBound(lines)
            lines(0) = lastLine + lines(0)
            
            Dim i As Long
            For i = 0 To (lineCnt - 1)
                lines(i) = Replace(lines(i), vbCr, "")
                If funcName <> "" Then
                    lines(i) = Application.Run(funcName, lines(i))
                End If
                outStream.WriteText CStr(lines(i)), adWriteLine
            Next
             
            lastLine = lines(lineCnt)
        Loop
        If lastLine <> "" Then
            outStream.WriteText lastLine, adWriteLine
        End If
    Else
        Do Until inStream.EOS
            Dim tmpLine As String: tmpLine = inStream.ReadText(adReadLine)
            If funcName <> "" Then
                tmpLine = Application.Run(funcName, tmpLine)
            End If
            outStream.WriteText tmpLine, adWriteChar
        Loop
    End If

    ' �t�@�C���ۑ�
    outStream.SaveToFile destFilePath, adSaveCreateOverWrite
     
    inStream.CloseStream
    outStream.CloseStream
End Sub

'******************************************************************************
'* [�T  �v] �f�B���N�g���p�X�������t�^�����B
'* [��  ��] �f�B���N�g���p�X�̖����ɕ������i���j���Ȃ���Εt�^���s���B
'*
'* @param strDirPath �f�B���N�g���p�X
'* @return �������t���f�B���N�g���p�X
'******************************************************************************
Private Function AddPathSeparator(strDirPath As String) As String
    If Right(strDirPath, 1) <> "\" Then
        AddPathSeparator = strDirPath & "\"
    Else
        AddPathSeparator = strDirPath
    End If
End Function
