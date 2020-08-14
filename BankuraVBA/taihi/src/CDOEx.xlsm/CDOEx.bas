Attribute VB_Name = "CDOEx"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] CDO���b�v�E�g�����W���[��
'* [��  ��] CDO��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�ACDO���g�p����
'*          ���[�e�B���e�B��񋟂���B
'*          ���b�v����CDO���C�u�����͈ȉ��̂��̂Ƃ���B
'*              [name] Microsoft CDO for Windows 2000 Library
'*              [dll] C:\Windows\System32\cdosys.dll
'* [�Q  �l]
'*  <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* CdoConfigSource
'*
'*-----------------------------------------------------------------------------
Public Enum CdoConfigSource
    cdoDefaults = -1        'Load all applicable default values from both Outlook Express and Internet Information Services.
    cdoIIS = 1              'Load configuration default values from the local Internet Information Service.
    cdoOutlookExpress = 2   'Load configuration values from the default identity of the default Outlook Express account.
End Enum

'*-----------------------------------------------------------------------------
'* CdoDSNOptions
'*
'*-----------------------------------------------------------------------------
Public Enum CdoDSNOptions
    cdoDSNDefault = 0             'No DSN commands are issued.
    cdoDSNDelay = 8               'Return a DSN if delivery is delayed.
    cdoDSNFailure = 2             'Return a DSN if delivery fails.
    cdoDSNNever = 1               'No DSNs are issued.
    cdoDSNSuccess = 4             'Return a DSN if delivery succeeds.
    cdoDSNSuccessFailOrDelay = 14 'Return a DSN if delivery succeeds, fails, or is delayed.
End Enum

'*-----------------------------------------------------------------------------
'* CdoEventStatus
'*
'*-----------------------------------------------------------------------------
Public Enum CdoEventStatus
    cdoRunNextSink = 0        'Proceed to run the next sink.
    cdoSkipRemainingSinks = 1 'Do not notify (skip) any remaining sinks for the event. This sink has consumed the event.
End Enum

'*-----------------------------------------------------------------------------
'* CdoImportanceValues
'*
'*-----------------------------------------------------------------------------
Public Enum CdoImportanceValues
    cdoHigh = 2   'The item is of high importance.
    cdoLow = 0    'The item is of low importance.
    cdoNormal = 1 'The item is of normal importance.
End Enum

'*-----------------------------------------------------------------------------
'* CdoMessageStat
'*
'*-----------------------------------------------------------------------------
Public Enum CdoMessageStat
    cdoStatAbortDelivery = 2 'Discard message and do not deliver.
    cdoStatBadMail = 3       'Do not deliver message and place it in the bad mail location.
    cdoStatSuccess = 0       'Success. Proceed to deliver message.
End Enum

'*-----------------------------------------------------------------------------
'* CdoMHTMLFlags
'*
'*-----------------------------------------------------------------------------
Public Enum CdoMHTMLFlags
    cdoSuppressAll = 31         'Do not download any resources referred to from within the page.
    cdoSuppressBGSounds = 2     'Do not download resources referred to in BGSOUND elements.
    cdoSuppressFrames = 4       'Do not download resources referred to in FRAME elements.
    cdoSuppressImages = 1       'Do not download resources referred to in IMG elements.
    cdoSuppressNone = 0         'Download all resources referred to in elements within the resource at the specified URI (not recursive).
    cdoSuppressObjects = 8      'Do not download resources referred to in OBJECT elements.
    cdoSuppressStyleSheets = 16 'Do not download resources referred to in LINK elements.
End Enum

'*-----------------------------------------------------------------------------
'* CdoNNTPProcessingField
'*
'*-----------------------------------------------------------------------------
Public Enum CdoNNTPProcessingField
    cdoPostMessage = 1      'Post the message.
    cdoProcessControl = 2   'Send message through process control.
    cdoProcessModerator = 4 'Send message to moderator.
End Enum

'*-----------------------------------------------------------------------------
'* CdoPostUsing
'*
'*-----------------------------------------------------------------------------
Public Enum CdoPostUsing
    cdoPostUsingPickup = 1 'Post the message using the local NNTP Service pickup directory.
    cdoPostUsingPort = 2   'Post the message using the NNTP protocol over the network.
End Enum

'*-----------------------------------------------------------------------------
'* CdoPriorityValues
'*
'*-----------------------------------------------------------------------------
Public Enum CdoPriorityValues
    cdoPriorityNonUrgent = -1 'The item is of non-urgent priority.
    cdoPriorityNormal = 0     'The item is of normal priority.
    cdoPriorityUrgent = 1     'The item is of urgent priority.
End Enum

'*-----------------------------------------------------------------------------
'* CdoProtocolsAuthentication
'*
'*-----------------------------------------------------------------------------
Public Enum CdoProtocolsAuthentication
    cdoAnonymous = 0 'Perform no authentication (anonymous).
    cdoBasic = 1     'Use the basic (clear text) authentication mechanism.
    cdoNTLM = 2      'Use the NTLM authentication mechanism
End Enum

'*-----------------------------------------------------------------------------
'* CdoReferenceType
'*
'*-----------------------------------------------------------------------------
Public Enum CdoReferenceType
    cdoRefTypeId = 0       'The reference parameter contains a value for the Content-ID header. The HTML body refers to the resource using this Content-ID header.
    cdoRefTypeLocation = 1 'The reference parameter contains a value for the Content-Location MIME header. The HTML body refers to this resource using this message-relative URL.
End Enum

'*-----------------------------------------------------------------------------
'* CdoSendUsing
'*
'*-----------------------------------------------------------------------------
Public Enum CdoSendUsing
    cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
    cdoSendUsingPort = 2   'Send the message using the SMTP protocol over the network.
End Enum

'*-----------------------------------------------------------------------------
'* CdoSensitivityValues
'*
'*-----------------------------------------------------------------------------
Public Enum CdoSensitivityValues
    cdoCompanyConfidential = 3 'The item is confidential to the company.
    cdoPersonal = 1            'The item is of personal sensitivity.
    cdoPrivate = 2             'The item is of private sensitivity.
    cdoSensitivityNone = 0     'The item is of no designated sensitivity.
End Enum

'*-----------------------------------------------------------------------------
'* CdoTimeZoneId
'*
'*-----------------------------------------------------------------------------
Public Enum CdoTimeZoneId
    cdoAbuDhabi = 24                      '(GMT+04:00) Abu Dhabi, Muscat
    cdoAdelaide = 19                      '(GMT+09:30) Adelaide, Central Australia
    cdoAdelaideCommonwealthGames2006 = 79 '(GMT+09:30) Adelaide (Commonwealth Games)
    cdoAlaska = 14                        '(GMT-09:00) Alaska
    cdoAlmaty = 46                        '(GMT+06:00) Almaty, North Central Asia, Novosibirsk
    cdoArab = 74                          '(GMT+03:00) Arab, Kuwait, Riyadh
    cdoArizona = 38                       '(GMT-07:00) Arizona
    cdoAthens = 7                         '(GMT+02:00) Athens, Istanbul, Minsk
    cdoAtlanticCanada = 9                 '(GMT-04:00) Atlantic Time (Canada)
    cdoAzores = 29                        '(GMT-01:00) Azores
    cdoBaghdad = 26                       '(GMT+03:00) Baghdad
    cdoBangkok = 22                       '(GMT+07:00) Bangkok, Jakarta, Hanoi
    cdoBeijing = 45                       '(GMT+08:00) Beijing, Chongqing, Hong Kong, Urumqi
    cdoBerlin = 4                         '(GMT+01:00) Berlin, Stockholm, Rome, Bern, Vienna
    cdoBogota = 35                        '(GMT-05:00) Bogota, Lima
    cdoBombay = 23                        '(GMT+05:30) Kolkota, Chennai, Mumbai, New Delhi, India Standard Time
    cdoBrasilia = 8                       '(GMT-03:00) Brasilia
    cdoBrisbane = 18                      '(GMT+10:00) Brisbane, East Australia
    cdoBuenosAires = 32                   '(GMT-03:00) Buenos Aires, Georgetown
    cdoCairo = 49                         '(GMT+02:00) Cairo
    cdoCanberraCommonwealthGames2006 = 78 '(GMT+10:00) Canberra, Melbourne, Sydney (Commonwealth Games)
    cdoCapeVerde = 53                     '(GMT-01:00) Cape Verde Is.
    cdoCaracas = 33                       '(GMT-04:00) Caracas, La Paz
    cdoCaucasus = 54                      '(GMT+04:00) Caucasus, Baku, Tbilisi, Yerevan
    cdoCentral = 11                       '(GMT-06:00) Central Time
    cdoCentralAmerica = 55                '(GMT-06:00) Central America
    cdoChihuahua = 77                     '(GMT-07:00) Chihuahua, La Paz, Mazatlan
    cdoDarwin = 44                        '(GMT+09:30) Darwin
    cdoDhaka = 71                         '(GMT+06:00) Dhaka
    cdoEastAfrica = 56                    '(GMT+03:00) East Africa, Nairobi
    cdoEastern = 10                       '(GMT-05:00) Eastern Time (US & Canada)
    cdoEasternEurope = 5                  '(GMT+02:00) Bucharest, Eastern Europe
    cdoEkaterinburg = 58                  '(GMT+05:00) Ekaterinburg
    cdoEniwetok = 39                      '(GMT-12:00) Eniwetok, Kwajalein, Dateline Time
    cdoFiji = 40                          '(GMT+12:00) Fiji, Kamchatka, Marshall Is.
    cdoFloating = 52                      'The time zone is floating.
    cdoGMT = 1                            '(GMT) Greenwich Mean Time; Dublin, Edinburgh, London
    cdoGreenland = 60                     '(GMT-03:00) Greenland
    cdoGuam = 43                          '(GMT+10:00) Guam, Port Moresby
    cdoHarare = 50                        '(GMT+02:00) Harare, Pretoria
    cdoHawaii = 15                        '(GMT-10:00) Hawaii
    cdoHelsinki = 59                      '(GMT+02:00) Helsinki
    cdoHobart = 42                        '(GMT+10:00) Hobart, Tasmania
    cdoHobartCommonwealthGames2006 = 80   '(GMT+10:00) Hobart (Commonwealth Games)
    cdoIndiana = 34                       '(GMT-05:00) Indiana (East)
    cdoInvalidTimeZone = 82               'The time zone is unrecognized or invalid.
    cdoIrkutsk = 63                       '(GMT+08:00) Irkutsk
    cdoIslamabad = 47                     '(GMT+05:00) Islamabad, Karachi, Sverdlovsk, Tashkent
    cdoIsrael = 27                        '(GMT+02:00) Israel, Jerusalem Standard Time
    cdoKabul = 48                         '(GMT+04:30) Kabul
    cdoKrasnoyarsk = 64                   '(GMT+07:00) Krasnoyarsk
    cdoMagadan = 41                       '(GMT+11:00) Magadan, Solomon Is., New Caledonia
    cdoMelbourne = 57                     '(GMT+10:00) Melbourne, Sydney
    cdoMexicoCity = 37                    '(GMT-06:00) Mexico City, Tegucigalpa
    cdoMidAtlantic = 30                   '(GMT-02:00) Mid-Atlantic
    cdoMidwayIsland = 16                  '(GMT-11:00) Midway Island, Samoa
    cdoMonrovia = 31                      '(GMT) Monrovia, Casablanca
    cdoMoscow = 51                        '(GMT+03:00) Moscow, St. Petersburg, Volgograd
    cdoMountain = 12                      '(GMT-07:00) Mountain Time (US & Canada)
    cdoNepal = 62                         '(GMT+05:45) Kathmandu, Nepal
    cdoNewfoundland = 28                  '(GMT-03:30) Newfoundland
    cdoPacific = 13                       '(GMT-08:00) Pacific Time (US & Canada)
    cdoParis = 3                          '(GMT+01:00) Paris, Madrid, Brussels, Copenhagen
    cdoPerth = 73                         '(GMT+08:00) Perth, Western Australia
    cdoPrague = 6                         '(GMT+01:00) Prague, Central Europe
    cdoRangoon = 61                       '(GMT+06:30) Rangoon
    cdoSantiago = 65                      '(GMT-04:00) Santiago
    cdoSarajevo = 2                       '(GMT+01:00) Sarajevo, Warsaw, Zagreb
    cdoSaskatchewan = 36                  '(GMT-06:00) Saskatchewan
    cdoSeoul = 72                         '(GMT+09:00) Seoul, Korea Standard Time
    cdoSingapore = 21                     '(GMT+08:00) Kuala Lumpur, Singapore
    cdoSriLanka = 66                      '(GMT+06:00) Sri Jayawardenepura, Sri Lanka
    cdoSydney2000 = 76                    '(GMT+10:00) Canberra, Melbourne, Sydney, Hobart (Year 2000 only)
    cdoTaipei = 75                        '(GMT+08:00) Taipei
    cdoTehran = 25                        '(GMT+03:30) Tehran
    cdoTijuana = 81                       '(GMT-08:00) Tijuana, Baja California
    cdoTokyo = 20                         '(GMT+09:00) Tokyo, Osaka, Sapporo
    cdoTonga = 67                         '(GMT+13:00) Tonga, Nuku'alofa
    cdoUTC = 0                            '(UTC) Universal Coordinated Time
    cdoVladivostok = 68                   '(GMT+10:00) Vladivostok
    cdoWellington = 17                    '(GMT+12:00) Wellington, Auckland
    cdoWestCentralAfrica = 69             '(GMT+01:00) West Central Africa
    cdoYakutsk = 70                       '(GMT+09:00) Yakutsk
End Enum

'*-----------------------------------------------------------------------------
'* �g��Enum��`
'*-----------------------------------------------------------------------------
'*-----------------------------------------------------------------------------
'* ADDODB���F�t�B�[���h�A�p�����[�^�[�A�܂��̓v���p�e�B�̃f�[�^�^���w�肵�܂��B
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
'* ADDODB���FField �I�u�W�F�N�g�̏�Ԃ�\���܂��B
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
'* �e�L�X�g Stream �I�u�W�F�N�g�̍s��؂�L���Ɏg���Ă��镶����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum LineSeparatorEnum
    adCR = 13   '���A�������܂��B
    adCRLF = -1 '����l�B���A���s�������܂��B
    adLF = 10   '���s�������܂��B
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
'* Stream �I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum StreamTypeEnum
    adTypeBinary = 1 '�o�C�i�� �f�[�^�������܂��B
    adTypeText = 2   '����l�BCharset �Ŏw�肳�ꂽ�����Z�b�g�̃e�L�X�g �f�[�^�������܂��B
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
'* Stream �I�u�W�F�N�g����t�@�C���ɕۑ�����Ƃ��ɁA�t�@�C�����쐬���邩�A
'* �㏑�����邩��\���܂��B�����̒l�� AND ���Z�q�Ō����ł��܂��B
'*-----------------------------------------------------------------------------
Public Enum SaveOptionsEnum
    adSaveCreateNotExist = 1  '����l�BFileName �p�����[�^�[�Ŏw�肵���t�@�C�����Ȃ��ꍇ�͐V�����t�@�C�����쐬����܂��B
    adSaveCreateOverWrite = 2 'Filename �p�����[�^�[�Ŏw�肵���t�@�C��������ꍇ�́A���݊J����Ă��� Stream �I�u�W�F�N�g�̃f�[�^�Ńt�@�C�����㏑������܂��B
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
'* Resync �̌Ăяo���ɂ���Ċ�ɂȂ�l���㏑������邩�ǂ�����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum ResyncEnum
    adResyncAllValues = 2        '����l�B�f�[�^�͏㏑������A�ۗ����̍X�V�͎�������܂��B
    adResyncUnderlyingValues = 1 '�f�[�^�͏㏑�����ꂸ�A�ۗ����̍X�V�͎�������܂���B
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
'* Stream �I�u�W�F�N�g����A�X�g���[���S�̂�ǂݎ�邩�A�܂��͎��̍s��ǂݎ�邩��\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum StreamReadEnum
    adReadAll = -1  '����l�B���݂̈ʒu���� EOS �}�[�J�[�����ɁA���ׂẴo�C�g���X�g���[������ǂݎ��܂��B����́A�o�C�i�� �X�g���[�� (Type �� adTypeBinary) �ɗB��L���� StreamReadEnum �l�ł��B
    adReadLine = -2 '�X�g���[�����玟�̍s��ǂݎ��܂� (LineSeparator �v���p�e�B�Ŏw��)�B
End Enum

'******************************************************************************
'* �萔��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* CDO.CdoCalendar �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoTimeZoneIDURN = "urn:schemas:calendar:timezoneid"

'*-----------------------------------------------------------------------------
'* CDO.CdoCharset �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoBIG5 = "big5"
Public Const cdoEUC_JP = "euc-jp"
Public Const cdoEUC_KR = "euc-kr"
Public Const cdoGB2312 = "gb2312"
Public Const cdoISO_2022_JP = "iso-2022-jp"
Public Const cdoISO_2022_KR = "iso-2022-kr"
Public Const cdoISO_8859_1 = "iso-8859-1"
Public Const cdoISO_8859_15 = "iso-8859-15"
Public Const cdoISO_8859_2 = "iso-8859-2"
Public Const cdoISO_8859_3 = "iso-8859-3"
Public Const cdoISO_8859_4 = "iso-8859-4"
Public Const cdoISO_8859_5 = "iso-8859-5"
Public Const cdoISO_8859_6 = "iso-8859-6"
Public Const cdoISO_8859_7 = "iso-8859-7"
Public Const cdoISO_8859_8 = "iso-8859-8"
Public Const cdoISO_8859_9 = "iso-8859-9"
Public Const cdoKOI8_R = "koi8-r"
Public Const cdoShift_JIS = "shift-jis"
Public Const cdoUS_ASCII = "us-ascii"
Public Const cdoUTF_7 = "utf-7"
Public Const cdoUTF_8 = "utf-8"

'*-----------------------------------------------------------------------------
'* CDO.CdoConfiguration �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoAutoPromoteBodyParts = "http://schemas.microsoft.com/cdo/configuration/autopromotebodyparts"
Public Const cdoFlushBuffersOnWrite = "http://schemas.microsoft.com/cdo/configuration/flushbuffersonwrite"
Public Const cdoHTTPCookies = "http://schemas.microsoft.com/cdo/configuration/httpcookies"
Public Const cdoLanguageCode = "http://schemas.microsoft.com/cdo/configuration/languagecode"
Public Const cdoNNTPAccountName = "http://schemas.microsoft.com/cdo/configuration/nntpaccountname"
Public Const cdoNNTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/nntpauthenticate"
Public Const cdoNNTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/nntpconnectiontimeout"
Public Const cdoNNTPServer = "http://schemas.microsoft.com/cdo/configuration/nntpserver"
Public Const cdoNNTPServerPickupDirectory = "http://schemas.microsoft.com/cdo/configuration/nntpserverpickupdirectory"
Public Const cdoNNTPServerPort = "http://schemas.microsoft.com/cdo/configuration/nntpserverport"
Public Const cdoNNTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/nntpusessl"
Public Const cdoPostEmailAddress = "http://schemas.microsoft.com/cdo/configuration/postemailaddress"
Public Const cdoPostPassword = "http://schemas.microsoft.com/cdo/configuration/postpassword"
Public Const cdoPostUserName = "http://schemas.microsoft.com/cdo/configuration/postusername"
Public Const cdoPostUserReplyEmailAddress = "http://schemas.microsoft.com/cdo/configuration/postuserreplyemailaddress"
Public Const cdoPostUsingMethod = "http://schemas.microsoft.com/cdo/configuration/postusing"
Public Const cdoSaveSentItems = "http://schemas.microsoft.com/cdo/configuration/savesentitems"
Public Const cdoSendEmailAddress = "http://schemas.microsoft.com/cdo/configuration/sendemailaddress"
Public Const cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
Public Const cdoSendUserName = "http://schemas.microsoft.com/cdo/configuration/sendusername"
Public Const cdoSendUserReplyEmailAddress = "http://schemas.microsoft.com/cdo/configuration/senduserreplyemailaddress"
Public Const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
Public Const cdoSMTPAccountName = "http://schemas.microsoft.com/cdo/configuration/smtpaccountname"
Public Const cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
Public Const cdoSMTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
Public Const cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Public Const cdoSMTPServerPickupDirectory = "http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory"
Public Const cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Public Const cdoSMTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/smtpusessl"
Public Const cdoURLGetLatestVersion = "http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion"
Public Const cdoURLProxyBypass = "http://schemas.microsoft.com/cdo/configuration/urlproxybypass"
Public Const cdoURLProxyServer = "http://schemas.microsoft.com/cdo/configuration/urlproxyserver"
Public Const cdoUseMessageResponseText = "http://schemas.microsoft.com/cdo/configuration/usemessageresponsetext"

'*-----------------------------------------------------------------------------
'* CDO.CdoContentTypeValues �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoGif = "image/gif"
Public Const cdoJpeg = "image/jpeg"
Public Const cdoMessageExternalBody = "message/external-body"
Public Const cdoMessagePartial = "message/partial"
Public Const cdoMessageRFC822 = "message/rfc822"
Public Const cdoMultipartAlternative = "multipart/alternative"
Public Const cdoMultipartDigest = "multipart/digest"
Public Const cdoMultipartMixed = "multipart/mixed"
Public Const cdoMultipartRelated = "multipart/related"
Public Const cdoTextHTML = "text/html"
Public Const cdoTextPlain = "text/plain"

'*-----------------------------------------------------------------------------
'* CDO.CdoDAV �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoContentClass = "DAV:contentclass"
Public Const cdoGetContentType = "DAV:getcontenttype"

'*-----------------------------------------------------------------------------
'* CDO.CdoEncodingType �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdo7bit = "7bit"
Public Const cdo8bit = "8bit"
Public Const cdoBase64 = "base64"
Public Const cdoBinary = "binary"
Public Const cdoMacBinHex40 = "mac-binhex40"
Public Const cdoQuotedPrintable = "quoted-printable"
Public Const cdoUuencode = "uuencode"

'*-----------------------------------------------------------------------------
'* CDO.CdoErrors �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const CDO_E_ADOSTREAM_NOT_BOUND = -2147220948
Public Const CDO_E_ARGUMENT1 = -2147205120
Public Const CDO_E_ARGUMENT2 = -2147205119
Public Const CDO_E_ARGUMENT3 = -2147205118
Public Const CDO_E_ARGUMENT4 = -2147205117
Public Const CDO_E_ARGUMENT5 = -2147205116
Public Const CDO_E_AUTHENTICATION_FAILURE = -2147220971
Public Const CDO_E_BAD_ATTENDEE_DATA = -2147220939
Public Const CDO_E_BAD_DATA = -2147220951
Public Const CDO_E_BAD_SENDER = -2147220941
Public Const CDO_E_BAD_TASKTYPE_ONASSIGN = -2147220937
Public Const CDO_E_CONNECTION_DROPPED = -2147220974
Public Const CDO_E_CONTENTPROPXML_CONVERT_FAILED = -2147220944
Public Const CDO_E_CONTENTPROPXML_NOT_FOUND = -2147220947
Public Const CDO_E_CONTENTPROPXML_PARSE_FAILED = -2147220945
Public Const CDO_E_CONTENTPROPXML_WRONG_CHARSET = -2147220946
Public Const CDO_E_DIRECTORIES_UNREACHABLE = -2147220942
Public Const CDO_E_FAILED_TO_CONNECT = -2147220973
Public Const CDO_E_FROM_MISSING = -2147220979
Public Const CDO_E_HTTP_FAILED = -2147220966
Public Const CDO_E_HTTP_FORBIDDEN = -2147220967
Public Const CDO_E_HTTP_NOT_FOUND = -2147220968
Public Const CDO_E_INACTIVE = -2147220986
Public Const CDO_E_INVALID_CHARSET = -2147220949
Public Const CDO_E_INVALID_CONTENT_TYPE = -2147220970
Public Const CDO_E_INVALID_ENCODING_FOR_MULTIPART = -2147220964
Public Const CDO_E_INVALID_ENCODING_TYPE = -2146644451
Public Const CDO_E_INVALID_POST = -2147220972
Public Const CDO_E_INVALID_POST_OPTION = -2147220959
Public Const CDO_E_INVALID_PROPERTYNAME = -2147220988
Public Const CDO_E_INVALID_SEND_OPTION = -2147220960
Public Const CDO_E_LOGON_FAILURE = -2147220969
Public Const CDO_E_MULTIPART_NO_DATA = -2147220965
Public Const CDO_E_NNTP_POST_FAILED = -2147220976
Public Const CDO_E_NNTP_SERVER_REQUIRED = -2147220981
Public Const CDO_E_NO_DEFAULT_DROP_DIR = -2147220983
Public Const CDO_E_NO_DIRECTORIES_SPECIFIED = -2147220943
Public Const CDO_E_NO_METHOD = -2147220956
Public Const CDO_E_NO_PICKUP_DIR = -2147220958
Public Const CDO_E_NO_SUPPORT_FOR_OBJECTS = -2147220985
Public Const CDO_E_NOT_ALL_DELETED = -2147220957
Public Const CDO_E_NOT_ASSIGNEDTO_USER = -2147220936
Public Const CDO_E_NOT_AVAILABLE = -2147220984
Public Const CDO_E_NOT_FOUND = -2146644475
Public Const CDO_E_NOT_OPENED = -2147220990
Public Const CDO_E_OUTOFDATE = -2147220935
Public Const CDO_E_PROP_CANNOT_DELETE = -2147220952
Public Const CDO_E_PROP_NONHEADER = -2147220950
Public Const CDO_E_PROP_NOT_FOUND = -2147220962
Public Const CDO_E_PROP_READONLY = -2147220953
Public Const CDO_E_PROP_UNSUPPORTED = -2147220987
Public Const CDO_E_RECIPIENT_MISSING = -2147220980
Public Const CDO_E_RECIPIENTS_REJECTED = -2147220977
Public Const CDO_E_ROLE_NOMORE_AVAILABLE = -2147220938
Public Const CDO_E_SELF_BINDING = -2147220940
Public Const CDO_E_SENDER_REJECTED = -2147220978
Public Const CDO_E_SMTP_SEND_FAILED = -2147220975
Public Const CDO_E_SMTP_SERVER_REQUIRED = -2147220982
Public Const CDO_E_UNCAUGHT_EXCEPTION = -2147220991
Public Const CDO_E_UNSAFE_OPERATION = -2147220963
Public Const CDO_E_UNSUPPORTED_DATASOURCE = -2147220989

'*-----------------------------------------------------------------------------
'* CDO.CdoExchange �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoSensitivity = "http://schemas.microsoft.com/exchange/sensitivity"

'*-----------------------------------------------------------------------------
'* CDO.CdoHTTPMail �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoAttachmentFilename = "urn:schemas:httpmail:attachmentfilename"
Public Const cdoBcc = "urn:schemas:httpmail:bcc"
Public Const cdoCc = "urn:schemas:httpmail:cc"
Public Const cdoContentDispositionType = "urn:schemas:httpmail:content-disposition-type"
Public Const cdoContentMediaType = "urn:schemas:httpmail:content-media-type"
Public Const cdoDate = "urn:schemas:httpmail:date"
Public Const cdoDateReceived = "urn:schemas:httpmail:datereceived"
Public Const cdoFrom = "urn:schemas:httpmail:from"
Public Const cdoHasAttachment = "urn:schemas:httpmail:hasattachment"
Public Const cdoHTMLDescription = "urn:schemas:httpmail:htmldescription"
Public Const cdoImportance = "urn:schemas:httpmail:importance"
Public Const cdoNormalizedSubject = "urn:schemas:httpmail:normalizedsubject"
Public Const cdoPriority = "urn:schemas:httpmail:priority"
Public Const cdoReplyTo = "urn:schemas:httpmail:reply-to"
Public Const cdoSender = "urn:schemas:httpmail:sender"
Public Const cdoSubject = "urn:schemas:httpmail:subject"
Public Const cdoTextDescription = "urn:schemas:httpmail:textdescription"
Public Const cdoThreadTopic = "urn:schemas:httpmail:thread-topic"
Public Const cdoTo = "urn:schemas:httpmail:to"

'*-----------------------------------------------------------------------------
'* CDO.CdoInterfaces �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoAdoStream = "_Stream"
Public Const cdoIBodyPart = "IBodyPart"
Public Const cdoIConfiguration = "IConfiguration"
Public Const cdoIDataSource = "IDataSource"
Public Const cdoIMessage = "IMessage"
Public Const cdoIStream = "IStream"

'*-----------------------------------------------------------------------------
'* CDO.CdoMailHeader �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoApproved = "urn:schemas:mailheader:approved"
Public Const cdoComment = "urn:schemas:mailheader:comment"
Public Const cdoContentBase = "urn:schemas:mailheader:content-base"
Public Const cdoContentDescription = "urn:schemas:mailheader:content-description"
Public Const cdoContentDisposition = "urn:schemas:mailheader:content-disposition"
Public Const cdoContentId = "urn:schemas:mailheader:content-id"
Public Const cdoContentLanguage = "urn:schemas:mailheader:content-language"
Public Const cdoContentLocation = "urn:schemas:mailheader:content-location"
Public Const cdoContentTransferEncoding = "urn:schemas:mailheader:content-transfer-encoding"
Public Const cdoContentType = "urn:schemas:mailheader:content-type"
Public Const cdoControl = "urn:schemas:mailheader:control"
Public Const cdoDisposition = "urn:schemas:mailheader:disposition"
Public Const cdoDispositionNotificationTo = "urn:schemas:mailheader:disposition-notification-to"
Public Const cdoDistribution = "urn:schemas:mailheader:distribution"
Public Const cdoExpires = "urn:schemas:mailheader:expires"
Public Const cdoFollowupTo = "urn:schemas:mailheader:followup-to"
Public Const cdoInReplyTo = "urn:schemas:mailheader:in-reply-to"
Public Const cdoLines = "urn:schemas:mailheader:lines"
Public Const cdoMessageId = "urn:schemas:mailheader:message-id"
Public Const cdoMIMEVersion = "urn:schemas:mailheader:mime-version"
Public Const cdoNewsgroups = "urn:schemas:mailheader:newsgroups"
Public Const cdoOrganization = "urn:schemas:mailheader:organization"
Public Const cdoOriginalRecipient = "urn:schemas:mailheader:original-recipient"
Public Const cdoPath = "urn:schemas:mailheader:path"
Public Const cdoPostingVersion = "urn:schemas:mailheader:posting-version"
Public Const cdoReceived = "urn:schemas:mailheader:received"
Public Const cdoReferences = "urn:schemas:mailheader:references"
Public Const cdoRelayVersion = "urn:schemas:mailheader:relay-version"
Public Const cdoReturnPath = "urn:schemas:mailheader:return-path"
Public Const cdoReturnReceiptTo = "urn:schemas:mailheader:return-receipt-to"
Public Const cdoSummary = "urn:schemas:mailheader:summary"
Public Const cdoThreadIndex = "urn:schemas:mailheader:thread-index"
Public Const cdoXFidelity = "urn:schemas:mailheader:x-cdostreamhighfidelity"
Public Const cdoXMailer = "urn:schemas:mailheader:x-mailer"
Public Const cdoXref = "urn:schemas:mailheader:xref"
Public Const cdoXUnsent = "urn:schemas:mailheader:x-unsent"

'*-----------------------------------------------------------------------------
'* CDO.CdoNamespace �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoNSConfiguration = "http://schemas.microsoft.com/cdo/configuration/"
Public Const cdoNSContacts = "urn:schemas:contacts:"
Public Const cdoNSHTTPMail = "urn:schemas:httpmail:"
Public Const cdoNSMailHeader = "urn:schemas:mailheader:"
Public Const cdoNSNNTPEnvelope = "http://schemas.microsoft.com/cdo/nntpenvelope/"
Public Const cdoNSSMTPEnvelope = "http://schemas.microsoft.com/cdo/smtpenvelope/"

'*-----------------------------------------------------------------------------
'* CDO.CdoNNTPEnvelope �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoNewsgroupList = "http://schemas.microsoft.com/cdo/nntpenvelope/newsgrouplist"
Public Const cdoNNTPProcessing = "http://schemas.microsoft.com/cdo/nntpenvelope/nntpprocessing"

'*-----------------------------------------------------------------------------
'* CDO.CdoOffice �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoKeywords = "urn:schemas-microsoft-com:office:office#Keywords"

'*-----------------------------------------------------------------------------
'* CDO.CdoSMTPEnvelope �̃����o�[
'*
'*-----------------------------------------------------------------------------
Public Const cdoArrivalTime = "http://schemas.microsoft.com/cdo/smtpenvelope/arrivaltime"
Public Const cdoClientIPAddress = "http://schemas.microsoft.com/cdo/smtpenvelope/clientipaddress"
Public Const cdoMessageStatus = "http://schemas.microsoft.com/cdo/smtpenvelope/messagestatus"
Public Const cdoPickupFileName = "http://schemas.microsoft.com/cdo/smtpenvelope/pickupfilename"
Public Const cdoRecipientList = "http://schemas.microsoft.com/cdo/smtpenvelope/recipientlist"
Public Const cdoSenderEmailAddress = "http://schemas.microsoft.com/cdo/smtpenvelope/senderemailaddress"

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************


