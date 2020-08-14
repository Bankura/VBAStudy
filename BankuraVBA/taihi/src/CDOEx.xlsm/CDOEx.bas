Attribute VB_Name = "CDOEx"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CDOラップ・拡張モジュール
'* [詳  細] CDOのWrapperとしての機能を提供する他、CDOを使用した
'*          ユーティリティを提供する。
'*          ラップするCDOライブラリは以下のものとする。
'*              [name] Microsoft CDO for Windows 2000 Library
'*              [dll] C:\Windows\System32\cdosys.dll
'* [参  考]
'*  <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum定義
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
'* 拡張Enum定義
'*-----------------------------------------------------------------------------
'*-----------------------------------------------------------------------------
'* ADDODBより：フィールド、パラメーター、またはプロパティのデータ型を指定します。
'*
'*-----------------------------------------------------------------------------
Public Enum DataTypeEnum
    adArray = 8192          '常に別のデータ型定数と組み合わされ、そのデータ型の配列を示すフラグ値です。
    adBigInt = 20           '8 バイトの符号付き整数を示します (DBTYPE_I8)。
    adBinary = 128          'バイナリ値を示します (DBTYPE_BYTES)。
    adBoolean = 11          'ブール値を示します (DBTYPE_BOOL)。
    adBSTR = 8              'null で終わる文字列 (Unicode) を示します (DBTYPE_BSTR)。
    adChapter = 136         '子行セットの行を識別する 4 バイト チャプター値を示します (DBTYPE_HCHAPTER)。
    adChar = 129            '文字列値を示します (DBTYPE_STR)。
    adCurrency = 6          '通貨値を示します (DBTYPE_CY)。通貨型は小数点以下 4 桁の固定小数点の数値です。スケールが 10,000 の、8 バイトの符号付き整数で格納します。
    adDate = 7              '日付値を示します (DBTYPE_DATE)。日付は倍精度浮動小数点数型 (Double) で格納され、整数部分は 1899 年 12 月 30 日からの日数を、小数部分は時刻を表します。
    adDBDate = 133          '日付値 (yyyymmdd) を示します (DBTYPE_DBDATE)。
    adDBTime = 134          '時刻値 (hhmmss) を示します (DBTYPE_DBTIME)。
    adDBTimeStamp = 135     '日付/タイム スタンプ (yyyymmddhhmmss および 10 億分の 1 桁までの分数) を示します (DBTYPE_DBTIMESTAMP)。
    adDecimal = 14          '固定精度およびスケールの正確な数値を示します (DBTYPE_DECIMAL)。
    adDouble = 5            '倍精度浮動小数点値を示します (DBTYPE_R8)。
    adEmpty = 0             '値を指定しません (DBTYPE_EMPTY)。
    adError = 10            '32 ビット エラー コードを示します (DBTYPE_ERROR)。
    adFileTime = 64         '1601 年 1 月 1 日からの時間を 100 ナノ秒単位で示す 64 ビット値を示します (DBTYPE_FILETIME)。
    adGUID = 72             'グローバル一意識別子 (GUID) を示します (DBTYPE_GUID)。
    adIDispatch = 9         'COM オブジェクトの IDispatch インターフェイスへのポインターを示します (DBTYPE_IDISPATCH)。
    adInteger = 3           '4 バイトの符号付き整数を示します (DBTYPE_I4)。
    adIUnknown = 13         'COM オブジェクトの IUnknown インターフェイスへのポインターを示します (DBTYPE_IUNKNOWN)。
    adLongVarBinary = 205   'ロング バイナリ値を示します。
    adLongVarChar = 201     '長い文字列値を示します。
    adLongVarWChar = 203    '長い、null で終わる Unicode 文字列値を示します。
    adNumeric = 131         '固定精度およびスケールの正確な数値を示します (DBTYPE_NUMERIC)。
    adPropVariant = 138     'オートメーション PROPVARIANT を示します (DBTYPE_PROP_VARIANT)。
    adSingle = 4            '単精度浮動小数点値を示します (DBTYPE_R4)。
    adSmallInt = 2          '2 バイトの符号付き整数を示します (DBTYPE_I2)。
    adTinyInt = 16          '1 バイトの符号付き整数を示します (DBTYPE_I1)。
    adUnsignedBigInt = 21   '8 バイトの符号なし整数を示します (DBTYPE_UI8)。
    adUnsignedInt = 19      '4 バイトの符号なし整数を示します (DBTYPE_UI4)。
    adUnsignedSmallInt = 18 '2 バイトの符号なし整数を示します (DBTYPE_UI2)。
    adUnsignedTinyInt = 17  '1 バイトの符号なし整数を示します (DBTYPE_UI1)。
    adUserDefined = 132     'ユーザー定義の変数を示します (DBTYPE_UDT)。
    adVarBinary = 204       'バイナリ値を示します (Parameter オブジェクトのみ)。
    adVarChar = 200         '文字列値を示します。
    adVariant = 12          'オートメーション バリアント型 (Variant) を示します (DBTYPE_VARIANT)。
    adVarNumeric = 139      '数値を示します (Parameter オブジェクトのみ)。
    adVarWChar = 202        'null で終わる Unicode 文字列を示します。
    adWChar = 130           'null で終わる Unicode 文字列を示します (DBTYPE_WSTR)。
End Enum

'*-----------------------------------------------------------------------------
'* ADDODBより：Field オブジェクトの状態を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum FieldStatusEnum
    adFieldAlreadyExists = 26             '指定したフィールドが既に存在することを示します。
    adFieldBadStatus = 12                 'ADO から OLE DB プロバイダーに無効な状態値が送信されたことを示します。原因としては、OLE DB 1.0 プロバイダーまたは 1.1 プロバイダー、あるいは不適切な組み合わせの Value と Status が考えられます。
    adFieldCannotComplete = 20            'Source で指定された URL のサーバーが操作を完了できなかったことを示します。
    adFieldCannotDeleteSource = 23        '移動操作で、ツリーまたはサブツリーを新しい位置に移動したがソースを削除できなかったことを示します。
    adFieldCantConvertValue = 2           'フィールドの取得または保存を行うときにデータが失われてしまうことを示します。
    adFieldCantCreate = 7                 'プロバイダーの限度 (許容フィールド数など) を超えたためにフィールドを追加できなかったことを示します。
    adFieldDataOverflow = 6               'プロバイダーから返されたデータがフィールドのデータ型をオーバーフローしたことを示します。
    adFieldDefault = 13                   'データの設定時にフィールドの既定値が使われたことを示します。
    adFieldDoesNotExist = 16              '指定したフィールドが存在しないことを示します。
    adFieldIgnore = 15                    'ソースでのデータ値の設定時にこのフィールドがスキップされたことを示します。プロバイダーで値が設定されませんでした。
    adFieldIntegrityViolation = 10        '計算エンティティまたは派生エンティティであるため、フィールドを編集できないことを示します。
    adFieldInvalidURL = 17                'データ ソース URL に無効な文字があることを示します。
    adFieldIsNull = 3                     'プロバイダーが種類 VT_NULL のバリアント型 (VARIANT) の値を返し、フィールドが空でないことを示します。
    adFieldOK = 0                         '既定値。フィールドの追加または削除が正常に行われたことを示します。
    adFieldOutOfSpace = 22                '移動またはコピー操作を実行するために必要な記憶域をプロバイダーが確保できないことを示します。
    adFieldPendingChange = 262144         'フィールドが削除され、異なるデータ型を指定して再度追加されたか、以前に状態が adFieldOK であったフィールドの値が変更されたことを示します。Update メソッドの呼び出し後にフィールドの最終形式によって Fields コレクションが変更されます。
    adFieldPendingDelete = 131072         'Delete 操作で状態が設定されたことを示します。フィールドは、Update メソッドの呼び出し後に Fields コレクションから削除するようマークされています。
    adFieldPendingInsert = 65536          'Append 操作で状態が設定されたことを示します。Field は、Update メソッドの呼び出し後に Fields コレクションに追加するようマークされています。
    adFieldPendingUnknown = 524288        'フィールドの状態を設定する原因となった操作をプロバイダーが判別できないことを示します。
    adFieldPendingUnknownDelete = 1048576 'フィールドの状態を設定する原因となった操作をプロバイダーが判別できず、Update メソッドの呼び出し後に Fields コレクションからフィールドが削除されることを示します。
    adFieldPermissionDenied = 9           '読み取り専用として定義されているため、フィールドを編集できないことを示します。
    adFieldReadOnly = 24                  'データ ソース内のフィールドが読み取り専用として定義されていることを示します。
    adFieldResourceExists = 19            '宛先 URL にオブジェクトが既に存在し、上書きできないため、プロバイダーが操作を実行できなかったことを示します。
    adFieldResourceLocked = 18            'データ ソースが 1 つ以上の他のアプリケーションまたはプロセスによってロックされているため、プロバイダーが操作を実行できなかったことを示します。
    adFieldResourceOutOfScope = 25        'ソースまたは宛先の URL が現在のレコードの範囲外であることを示します。
    adFieldSchemaViolation = 11           '値がフィールドのデータ ソース スキーマ制約に違反することを示します。
    adFieldSignMismatch = 5               'プロバイダーが返すデータ値が符号付きで、ADO フィールド値のデータ型が符号なしであることを示します。
    adFieldTruncated = 4                  'データ ソースからの読み取り時に可変長データが切り捨てられたことを示します。
    adFieldUnavailable = 8                'データ ソースからの読み取り時にプロバイダーが値を判別できなかったことを示します。たとえば、行が作成された直後であること、列の既定値が使用不可であること、または新しい値がまだ指定されていないことが原因として考えられます。
    adFieldVolumeNotFound = 21            'URL が示す記憶域ボリュームをプロバイダーが特定できないことを示します。
End Enum

'*-----------------------------------------------------------------------------
'* オブジェクトが開いているか閉じているか、データ ソースに接続中か、コマンドを
'* 実行中か、またはデータを取得中かを表します。
'*-----------------------------------------------------------------------------
Public Enum ObjectStateEnum
    adStateClosed = 0     'オブジェクトが閉じていることを示します。
    adStateOpen = 1       'オブジェクトが開いていることを示します。
    adStateConnecting = 2 'オブジェクトが接続中であることを示します。
    adStateExecuting = 4  'オブジェクトがコマンドを実行中であることを示します。
    adStateFetching = 8   'オブジェクトの行を取得中であることを示します。
End Enum

'*-----------------------------------------------------------------------------
'* テキスト Stream オブジェクトの行区切り記号に使われている文字を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum LineSeparatorEnum
    adCR = 13   '復帰を示します。
    adCRLF = -1 '既定値。復帰改行を示します。
    adLF = 10   '改行を示します。
End Enum

'*-----------------------------------------------------------------------------
'* Connection 内のデータの編集、Record のオープン、または Record および Stream
'*  オブジェクトの Mode プロパティの値の指定に対する権限を表します。
'*-----------------------------------------------------------------------------
Public Enum ConnectModeEnum
    adModeRead = 1            '読み取り専用の権限を表します。
    adModeReadWrite = 3       '読み取り/書き込み両方の権限を表します。
    adModeRecursive = 4194304 '他の共有拒否値 (adModeShareDenyNone、 adModeShareDenyWrite、またはadModeShareDenyRead) と共に使用して、現在のレコードのすべてのサブレコードに共有制限を伝達します。 Record に子がない場合は機能しません。adModeShareDenyNone のみと組み合わせて使用すると、実行時エラーが発生します。 ただし、その他の値と組み合わせた場合は adModeShareDenyNone と組み合わせて使用できます。
    adModeShareDenyNone = 16  '権限の種類に関係なく、他のユーザーが接続を開けるようにします。他のユーザーに対して、読み取りと書き込みの両方のアクセスを許可します。
    adModeShareDenyRead = 4   '他のユーザーが読み取り権限で接続を開くのを禁止します。
    adModeShareDenyWrite = 8  '他のユーザーが書き込み権限で接続を開くのを禁止します。
    adModeShareExclusive = 12 '他のユーザーが接続を開くのを禁止します。
    adModeUnknown = 0         '既定値。権限が設定されていないか、権限を判定できないことを示します。
    adModeWrite = 2           '書き込み専用の権限を示します。
End Enum

'*-----------------------------------------------------------------------------
'* Stream オブジェクトに保存するデータの種類を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum StreamTypeEnum
    adTypeBinary = 1 'バイナリ データを示します。
    adTypeText = 2   '既定値。Charset で指定された文字セットのテキスト データを示します。
End Enum

'*-----------------------------------------------------------------------------
'* Stream オブジェクトを開くときのオプションを表します。
'* これらの値は OR 演算子で結合できます。
'*-----------------------------------------------------------------------------
Public Enum StreamOpenOptionsEnum
    adOpenStreamAsync = 1        '非同期モードで Stream オブジェクトを開きます。
    adOpenStreamFromRecord = 4   'Source パラメーターの内容を、既に開かれている Record オブジェクトとして識別します。既定動作では、Source は、ツリー構造のノードを直接指定する URL として処理します。このノードに関連付けられた既定ストリームが開かれます。
    adOpenStreamUnspecified = -1 '既定値。既定のオプションで Stream オブジェクトを開くことを表します。
End Enum

'*-----------------------------------------------------------------------------
'* Stream オブジェクトからファイルに保存するときに、ファイルを作成するか、
'* 上書きするかを表します。これらの値は AND 演算子で結合できます。
'*-----------------------------------------------------------------------------
Public Enum SaveOptionsEnum
    adSaveCreateNotExist = 1  '既定値。FileName パラメーターで指定したファイルがない場合は新しいファイルが作成されます。
    adSaveCreateOverWrite = 2 'Filename パラメーターで指定したファイルがある場合は、現在開かれている Stream オブジェクトのデータでファイルが上書きされます。
End Enum

'*-----------------------------------------------------------------------------
'* Stream オブジェクトに書き込む文字列に、行区切り記号を追加するかどうかを表します。
'*
'*-----------------------------------------------------------------------------
Public Enum StreamWriteEnum
    adWriteChar = 0 '既定値。Stream オブジェクトに、Data パラメーターで指定したテキスト文字列を書き込みます。
    adWriteLine = 1 'Stream オブジェクトに、テキスト文字列と行区切り記号を書き込みます。LineSeparator プロパティが定義されていない場合は、実行時エラーを返します。
End Enum

'*-----------------------------------------------------------------------------
'* Field オブジェクトの 1 つ以上の属性を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum FieldAttributeEnum
    adFldCacheDeferred = 4096     'プロバイダーでフィールド値がキャッシュされ、その後の読み取りはキャッシュから行われることを示します。
    adFldFixed = 16               'フィールドが固定長データを含むことを示します。
    adFldIsChapter = 8192         'フィールドがチャプター値を含み、この親フィールドに関連付けられた特定の子レコードセットを指定していることを示します。通常、チャプター フィールドはデータ シェイプやフィルター用に使います。
    adFldIsCollection = 262144    'レコードが示すリソースが、テキスト ファイルなどの単純なリソースではなく、フォルダーなどのように他のリソースのコレクションであることを、フィールドが表していることを示します。
    adFldIsDefaultStream = 131072 'フィールドが、レコードが示すリソースの既定ストリームを含むことを示します。 たとえば、既定のストリームは、web サイトのルートフォルダーの HTML コンテンツにすることができます。これは、ルート URL が指定されたときに自動的に提供されます。
    adFldIsNullable = 32          'フィールドに null 値を指定できることを示します。
    adFldIsRowURL = 65536         'フィールドが、レコードが示すデータ ストアのリソースを指定する URL を含むことを示します。
    adFldKeyColumn = 32768
    adFldLong = 128               'フィールドがロング バイナリ型のフィールドであることを示します。また、AppendChunk メソッドと GetChunk メソッドを使用できることを示します。
    adFldMayBeNull = 64           'フィールドからの null 値の読み取りが可能であることを示します。
    adFldMayDefer = 2             'フィールドが遅延フィールドであることを示します。フィールド値は、レコード全体のデータ ソースから取得されず、明示的にアクセスした場合のみ取得されます。
    adFldNegativeScale = 16384    '負のスケール値をサポートする列の数値を、フィールドが表していることを示します。スケールは、NumericScale プロパティで指定します。
    adFldRowID = 256              'フィールドが書き込み禁止の永続化された行識別子を含み、行を識別するもの (レコード番号、一意識別子など) 以外に有効な値は持たないことを示します。
    adFldRowVersion = 512         'フィールドが更新を記録するための時刻または日付スタンプを含むことを示します。
    adFldUnknownUpdatable = 8     'フィールドへの書き込みが可能かどうかをプロバイダーが確認できないことを示します。
    adFldUnspecified = -1         'プロバイダーがフィールド属性を指定しないことを示します。
    adFldUpdatable = 4            'フィールドへの書き込みが可能であることを示します。
End Enum

'*-----------------------------------------------------------------------------
'* Resync の呼び出しによって基になる値が上書きされるかどうかを表します。
'*
'*-----------------------------------------------------------------------------
Public Enum ResyncEnum
    adResyncAllValues = 2        '既定値。データは上書きされ、保留中の更新は取り消されます。
    adResyncUnderlyingValues = 1 'データは上書きされず、保留中の更新は取り消されません。
End Enum

'*-----------------------------------------------------------------------------
'* Record オブジェクトの Open メソッドに対し、既存の Record を開くか、新しい
'* Record を作成するかを表します。これらの値は AND 演算子で結合できます。
'*-----------------------------------------------------------------------------
Public Enum RecordCreateOptionsEnum
    adCreateCollection = 8192        '既存の Record を開かずに、Source パラメーターで指定したノードに新しい Record を作成します。ソースが既存のノードを指定している場合、adCreateCollection が adOpenIfExists または adCreateOverwrite と組み合わせて使用されていない限り、実行時エラーになります。
    adCreateNonCollection = 0        '種類が adSimpleRecord の新しい Record を作成します。
    adCreateOverwrite = 67108864     '作成フラグ adCreateCollection、adCreateNonCollection、および adCreateStructDoc を修飾します。この値と作成フラグの値の 1 つが OR を使って連結されている場合、ソース URL が既存のノードまたは Record を指定していると、既存のものが上書きされ、新しい Record が作成されます。この値は、adOpenIfExists とは併用できません。
    adCreateStructDoc = -2147483648# '既存の Record を開かずに、種類が adStructDoc の新しい Record を作成します。
    adFailIfNotExists = -1           '既定値。Source が存在しないノードを指定していると、実行時エラーになります。
    adOpenIfExists = 33554432        '作成フラグ adCreateCollection、adCreateNonCollection、および adCreateStructDoc を修飾します。この値と作成フラグの値の 1 つが OR を使って連結されている場合、ソース URL が既存のノードまたは Record オブジェクトを指定していると、プロバイダーは、新しい Record を作成せずに、既存のものを開く必要があります。この値は、adCreateOverwrite とは併用できません。
End Enum

'*-----------------------------------------------------------------------------
'* Record を開くときのオプションを表します。 これらの値は OR 演算子で結合できます。
'*
'*-----------------------------------------------------------------------------
Public Enum RecordOpenOptionsEnum
    adDelayFetchFields = 32768   'プロバイダーに対して、Record に関連付けられたフィールドは、当初は取得する必要がなく、フィールドへの最初のアクセス時に取得できることを示します。このフラグが指定されていない場合の既定動作では、Record オブジェクトのすべてのフィールドが取得されます。
    adDelayFetchStream = 16384   'プロバイダーに対して、Record に関連付けられた既定ストリームを当初は取得する必要がないことを示します。このフラグが指定されていない場合の既定動作では、Record オブジェクトに関連付けられた既定ストリームが取得されます。
    adOpenAsync = 4096           'Record オブジェクトが非同期モードで開かれることを示します。
    adOpenExecuteCommand = 65536 'Source 文字列に、実行されるコマンド テキストが含まれることを示します。この値は、Recordset.Open の adCmdText オプションと等価です。
    adOpenOutput = 8388608       '実行可能スクリプト (拡張子が .ASP のページなど) があるノードをソースが指定している場合、実行したスクリプトの結果が、開いている Record に含まれることを示します。この値は、コレクションのないレコードにのみ有効です。
    adOpenRecordUnspecified = -1 '既定値。オプションが指定されていないことを示します。
End Enum

'*-----------------------------------------------------------------------------
'* Stream オブジェクトから、ストリーム全体を読み取るか、または次の行を読み取るかを表します。
'*
'*-----------------------------------------------------------------------------
Public Enum StreamReadEnum
    adReadAll = -1  '既定値。現在の位置から EOS マーカー方向に、すべてのバイトをストリームから読み取ります。これは、バイナリ ストリーム (Type は adTypeBinary) に唯一有効な StreamReadEnum 値です。
    adReadLine = -2 'ストリームから次の行を読み取ります (LineSeparator プロパティで指定)。
End Enum

'******************************************************************************
'* 定数定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* CDO.CdoCalendar のメンバー
'*
'*-----------------------------------------------------------------------------
Public Const cdoTimeZoneIDURN = "urn:schemas:calendar:timezoneid"

'*-----------------------------------------------------------------------------
'* CDO.CdoCharset のメンバー
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
'* CDO.CdoConfiguration のメンバー
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
'* CDO.CdoContentTypeValues のメンバー
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
'* CDO.CdoDAV のメンバー
'*
'*-----------------------------------------------------------------------------
Public Const cdoContentClass = "DAV:contentclass"
Public Const cdoGetContentType = "DAV:getcontenttype"

'*-----------------------------------------------------------------------------
'* CDO.CdoEncodingType のメンバー
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
'* CDO.CdoErrors のメンバー
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
'* CDO.CdoExchange のメンバー
'*
'*-----------------------------------------------------------------------------
Public Const cdoSensitivity = "http://schemas.microsoft.com/exchange/sensitivity"

'*-----------------------------------------------------------------------------
'* CDO.CdoHTTPMail のメンバー
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
'* CDO.CdoInterfaces のメンバー
'*
'*-----------------------------------------------------------------------------
Public Const cdoAdoStream = "_Stream"
Public Const cdoIBodyPart = "IBodyPart"
Public Const cdoIConfiguration = "IConfiguration"
Public Const cdoIDataSource = "IDataSource"
Public Const cdoIMessage = "IMessage"
Public Const cdoIStream = "IStream"

'*-----------------------------------------------------------------------------
'* CDO.CdoMailHeader のメンバー
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
'* CDO.CdoNamespace のメンバー
'*
'*-----------------------------------------------------------------------------
Public Const cdoNSConfiguration = "http://schemas.microsoft.com/cdo/configuration/"
Public Const cdoNSContacts = "urn:schemas:contacts:"
Public Const cdoNSHTTPMail = "urn:schemas:httpmail:"
Public Const cdoNSMailHeader = "urn:schemas:mailheader:"
Public Const cdoNSNNTPEnvelope = "http://schemas.microsoft.com/cdo/nntpenvelope/"
Public Const cdoNSSMTPEnvelope = "http://schemas.microsoft.com/cdo/smtpenvelope/"

'*-----------------------------------------------------------------------------
'* CDO.CdoNNTPEnvelope のメンバー
'*
'*-----------------------------------------------------------------------------
Public Const cdoNewsgroupList = "http://schemas.microsoft.com/cdo/nntpenvelope/newsgrouplist"
Public Const cdoNNTPProcessing = "http://schemas.microsoft.com/cdo/nntpenvelope/nntpprocessing"

'*-----------------------------------------------------------------------------
'* CDO.CdoOffice のメンバー
'*
'*-----------------------------------------------------------------------------
Public Const cdoKeywords = "urn:schemas-microsoft-com:office:office#Keywords"

'*-----------------------------------------------------------------------------
'* CDO.CdoSMTPEnvelope のメンバー
'*
'*-----------------------------------------------------------------------------
Public Const cdoArrivalTime = "http://schemas.microsoft.com/cdo/smtpenvelope/arrivaltime"
Public Const cdoClientIPAddress = "http://schemas.microsoft.com/cdo/smtpenvelope/clientipaddress"
Public Const cdoMessageStatus = "http://schemas.microsoft.com/cdo/smtpenvelope/messagestatus"
Public Const cdoPickupFileName = "http://schemas.microsoft.com/cdo/smtpenvelope/pickupfilename"
Public Const cdoRecipientList = "http://schemas.microsoft.com/cdo/smtpenvelope/recipientlist"
Public Const cdoSenderEmailAddress = "http://schemas.microsoft.com/cdo/smtpenvelope/senderemailaddress"

'******************************************************************************
'* メソッド定義
'******************************************************************************


