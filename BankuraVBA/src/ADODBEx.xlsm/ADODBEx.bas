Attribute VB_Name = "ADODBEx"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODBラップ・拡張モジュール
'* [詳  細] ADODBのWrapperとしての機能を提供する他、ADODBを使用した
'*          ユーティリティを提供する。
'*          ラップするADODBライブラリは以下のものとする。
'*              [name] Microsoft ActiveX Data Objects 6.1 Library
'*              [dll] C:\Program Files\Common Files\System\ado\msado15.dll
'* [参  考]
'*  <https://docs.microsoft.com/ja-jp/sql/ado/microsoft-activex-data-objects-ado?view=sql-server-2017>
'*  <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-enumerated-constants>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* RDS の Recordset オブジェクトに対して、データを取得する非同期スレッドの
'* 実行優先度を表します｡
'*-----------------------------------------------------------------------------
Public Enum ADCPROP_ASYNCTHREADPRIORITY_ENUM
    adPriorityLowest = 1      '優先度を可能な最低レベルに設定します。
    adPriorityBelowNormal = 2 '優先度を最低と標準の間に設定します。
    adPriorityNormal = 3      '優先度を標準に設定します。
    adPriorityAboveNormal = 4 '優先度を標準と最高の間に設定します。
    adPriorityHighest = 5     '優先度を可能な最高レベルに設定します。
End Enum

'*-----------------------------------------------------------------------------
'* 階層 Recordset の集計列と計算列を MSDataShape プロバイダーがいつ再計算するか
'* を指定します。
'*-----------------------------------------------------------------------------
Public Enum ADCPROP_AUTORECALC_ENUM
    adRecalcAlways = 1  '既定値です。計算列が依存する値が変更されたと MSDataShape プロバイダーが判断したときに再計算します。
    adRecalcUpFront = 0 '階層 Recordset の最初の作成時のみ計算します。
End Enum

'*-----------------------------------------------------------------------------
'* Recordset オブジェクトを使用してデータ ソース行の共有的更新を行う際に、競合
'* の検出に使用するフィールドを表します。
'*-----------------------------------------------------------------------------
Public Enum ADCPROP_UPDATECRITERIA_ENUM
    adCriteriaAllCols = 1   'データ ソース行の列が変更された場合に競合を検出します。
    adCriteriaKey = 0       'データ ソース行のキー列が変更された場合、つまり行が削除された場合に競合を検出します。
    adCriteriaTimeStamp = 3 'データ ソース行のタイムスタンプが変更された場合、つまり Recordset を取得した後に行にアクセスがあった場合に競合を検出します。
    adCriteriaUpdCols = 2   'Recordset の更新されたフィールドに対応するデータ ソース行の列が変更された場合に競合を検出します。
End Enum

'*-----------------------------------------------------------------------------
'* UpdateBatch メソッドに暗黙の Resync メソッド操作が続くかどうかを示し、続く
'* 場合はその操作の適用範囲を指定します。
'*-----------------------------------------------------------------------------
Public Enum ADCPROP_UPDATERESYNC_ENUM
    adResyncAll = 15          '他のすべての ADCPROP_UPDATERESYNC_ENUM メンバーの結合された値を使用して、Resync を呼び出します。
    adResyncAutoIncrement = 1 '既定値です。Microsoft Jet AutoNumber フィールドや Microsoft SQL Server の Identity 列など、データ ソースによって自動的に増分または生成される列の新しい ID 値を取得します。
    adResyncConflicts = 2     '同時実行の競合により更新操作または削除操作が失敗したすべての行について、Resync を呼び出します。
    adResyncInserts = 8       '正常に挿入されたすべての行について、Resync を呼び出します。 ただし、AutoIncrement 列の値は再同期されません。 代わりに、既存の主キーの値に基づいて、新しく挿入された行の内容が再同期されます。 主キーが AutoIncrement 値の場合、Resync では対象の行の内容を取得しません。 オートインクリメントの主キー値を自動的にインクリメントするには、 adResyncAutoIncrement + adResyncInsertsを組み合わせた値でUpdateBatchを呼び出します。
    adResyncNone = 0          'Resync を呼び出しません。
    adResyncUpdates = 4       '正常に更新されたすべての行について、Resync を呼び出します。
End Enum

'*-----------------------------------------------------------------------------
'* 操作の対象となるレコードを表します。
'*
'*-----------------------------------------------------------------------------
Public Enum AffectEnum
    adAffectAll = 3         'Recordset に適用されている Filter がない場合、すべてのレコードが対象です。 Filterプロパティが文字列抽出条件 (Author = ' Smith ' ""など) に設定されている場合、操作は現在のチャプターの表示されているレコードに影響します。 filterプロパティがfiltergroupenumのメンバーまたはブックマークの配列に設定されている場合、この操作はRecordsetのすべての行に影響します。
    adAffectAllChapters = 4 '現在適用されている Filter で非表示になっているレコードを含む、Recordset のすべての兄弟チャプターの全レコードに反映されます。
    adAffectCurrent = 1     '現在のレコードにのみ反映されます。
    adAffectGroup = 2       '現在の Filter プロパティの設定を満たすレコードにのみ反映されます。このオプションを使用するには、Filter プロパティを FilterGroupEnum 値または Bookmark の配列に設定する必要があります。
End Enum

'*-----------------------------------------------------------------------------
'* 操作の開始位置を示すブックマークを表します。
'*
'*-----------------------------------------------------------------------------
Public Enum BookmarkEnum
    adBookmarkCurrent = 0 '現在のレコードから開始します。
    adBookmarkFirst = 1   '最初のレコードから開始します。
    adBookmarkLast = 2    '最後のレコードから開始します。
End Enum

'*-----------------------------------------------------------------------------
'* コマンド引数の解釈方法を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum CommandTypeEnum
    adCmdUnspecified = -1  'Does not specify the command type argument.
    adCmdText = 1          'CommandText を、コマンドまたはストアド プロシージャのテキスト定義として評価します。
    adCmdTable = 2         'CommandText を、内部的に生成された SQL クエリから返された列のみで構成されるテーブル名として評価します。
    adCmdStoredProc = 4    'CommandText をストアド プロシージャ名として評価します。
    adCmdUnknown = 8       '既定値。CommandText プロパティのコマンドの種類が不明であることを示します。
    adCmdFile = 256        'CommandText を、保存された Recordset のファイル名として評価します。Recordset.Open または Requery と組み合わせてのみ使用できます。
    adCmdTableDirect = 512 'CommandText を、すべての列が返されたテーブル名として評価します。 Recordset.Open または Requery と組み合わせてのみ使用できます。 Seek メソッドを使用する場合、Recordset は adCmdTableDirect を指定して開く必要があります。 この値は、ExecuteOptionEnum の値 adAsyncExecute と組み合わせて使用できません。
End Enum

'*-----------------------------------------------------------------------------
'* ブックマークで表された 2 つのレコードの相対位置を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum CompareEnum
    adCompareEqual = 1         'ブックマークが等しいことを示します。
    adCompareGreaterThan = 2   '最初のブックマークが 2 番目のブックマークの後になることを示します。
    adCompareLessThan = 0      '最初のブックマークが 2 番目のブックマークの前になることを示します。
    adCompareNotComparable = 4 'ブックマークを比較できないことを示します。
    adCompareNotEqual = 3      '2 つのブックマークは異なっており、順位がないことを示します。
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
'* Connection オブジェクトの Open メソッドから制御が戻るのが、接続確立の後
'* (同期) か前 (非同期) かを表します。
'*-----------------------------------------------------------------------------
Public Enum ConnectOptionEnum
    adAsyncConnect = 16       '接続を非同期で開きます。接続可能かをどうかを判別するために、ConnectComplete イベントが使用される場合があります。
    adConnectUnspecified = -1 'Default. Opens the connection synchronously.
End Enum

'*-----------------------------------------------------------------------------
'* データ ソースとの接続を開くときに、不足しているパラメーターを要求するダイア
'* ログ ボックスを表示するかどうかを表します。
'*-----------------------------------------------------------------------------
Public Enum ConnectPromptEnum
    adPromptAlways = 1           '常に要求します。
    adPromptComplete = 2         'さらに情報が必要な場合に要求します。
    adPromptCompleteRequired = 3 'さらに情報が必要だが、任意のパラメーターが禁止されている場合に要求します。
    adPromptNever = 4            '要求しません。
End Enum

'*-----------------------------------------------------------------------------
'* CopyRecord メソッドの動作を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum CopyRecordOptionsEnum
    adCopyAllowEmulation = 4 '"コピー先" が別のサーバーにあるか、"コピー元" 以外のプロバイダーのサービスを受けているためにこのメソッドが失敗した場合、"コピー元" プロバイダーがダウンロード操作とアップロード操作を行ってコピーをシミュレートしようとすることを示します。プロバイダーの機能が異なると、パフォーマンスが低下したりデータが失われることがあります。
    adCopyNonRecursive = 2   'コピー先に現在のディレクトリをコピーしますが、サブディレクトリはコピーしません。コピー操作は再帰的ではありません。
    adCopyOverWrite = 1      '"コピー先" が既存のファイルやディレクトリを指す場合、そのファイルやディレクトリを上書きします。
    adCopyUnspecified = -1   '既定値。既定のコピー操作を実行します。コピー操作は再帰的に行われ、コピー先のファイルやディレクトリが既に存在する場合は操作が失敗します。
End Enum

'*-----------------------------------------------------------------------------
'* カーソル サービスの場所を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum CursorLocationEnum
    adUseClient = 3 'ローカルカーソルライブラリによって提供されるクライアント側カーソルを使用します。 多くの場合、ローカル カーソル サービスにはドライバーによって提供されるカーソルよりも多くのカーソル機能があるので、この設定を利用すると、より高度な機能を提供できます。 以前のバージョンとの互換性を保つために、同じ意味を持つ adUseClientBatch もサポートしています。
    adUseServer = 2 '既定値。 データ プロバイダー カーソルまたはドライバーによって提供されるカーソルを使用します。 これらのカーソルは、多くの場合柔軟性が高く、他のユーザーが行うデータ ソースへの変更を検出できます。 ただし、 Microsoft Cursor Service for OLE DB (関連付けられていないRecordsetオブジェクトなど) の一部の機能は、サーバー側カーソルを使用してシミュレートすることはできません。この設定では、これらの機能は使用できません。
    adUseNone = 1   'Does not use cursor services. (This constant is obsolete and appears solely for the sake of backward compatibility.)
End Enum

'*-----------------------------------------------------------------------------
'* Supports メソッドがテストする機能を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum CursorOptionEnum
    adAddNew = 16778240      '新しいレコードを追加する AddNew メソッドをサポートします。
    adApproxPosition = 16384 'AbsolutePosition プロパティと AbsolutePage プロパティをサポートします。
    adBookmark = 8192        '特定のレコードへのアクセスを確保する Bookmark プロパティをサポートします。
    adDelete = 16779264      'レコードを削除する Delete メソッドをサポートします。
    adFind = 524288          'Recordset 内の行の位置を確認する Find メソッドをサポートします。
    adHoldRecords = 256      '保留中のすべての変更をコミットせずに、新たなレコードを格納するか、または次の格納位置を変更します。
    adIndex = 8388608        'インデックスに名前を付ける Index プロパティをサポートします。
    adMovePrevious = 512     'ブックマークを使用せずに現在のレコードの位置を後方に移動する MoveFirst メソッドと MovePrevious メソッド、および Move メソッドと GetRows メソッドをサポートします。
    adNotify = 262144        '基になるデータ プロバイダーが通知をサポートしていることを示します (これにより Recordset イベントのサポートの有無が決まります)。
    adResync = 131072        '基になるデータベースのカーソルにある可視データを更新する Resync メソッドをサポートします。
    adSeek = 4194304         'Recordset 内の行を検索する Seek メソッドをサポートします。
    adUpdate = 16809984      '既存のデータを変更する Update メソッドをサポートします。
    adUpdateBatch = 65536    '複数の変更をグループとしてプロバイダーに送信するバッチ更新 (UpdateBatch メソッドと CancelBatch メソッド) をサポートします。
End Enum

'*-----------------------------------------------------------------------------
'* Recordset オブジェクトが使用するカーソルの種類を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum CursorTypeEnum
    adOpenDynamic = 2      '動的カーソルを使用します。他のユーザーによる追加、変更、および削除が表示され、プロバイダーがブックマークをサポートしていない場合を除き、Recordset 内でのすべての種類の移動を許可します。
    adOpenForwardOnly = 0  '既定値。前方スクロール カーソルを使用します。レコードのスクロール方向が前方向に限定されていることを除き、静的カーソルと同じ働きをします。Recordset のスクロールが 1 回だけで十分な場合は、これによってパフォーマンスを向上できます。
    adOpenKeyset = 1       'キーセット カーソルを使います。自分の Recordset から他のユーザーが削除したレコードはアクセスできませんが、他のユーザーが追加したレコードは表示できない点を除いて動的カーソルと同じです。他のユーザーが変更したデータは表示できます。
    adOpenStatic = 3       '静的カーソルを使用します。データの検索やレポートの生成に使用できるの静的コピーです。他のユーザーによる追加、変更、または削除は表示されません。
    adOpenUnspecified = -1 'カーソルの種類を指定しません。
End Enum

'*-----------------------------------------------------------------------------
'* フィールド、パラメーター、またはプロパティのデータ型を指定します。
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
'* レコードの編集状況を示します。
'*
'*-----------------------------------------------------------------------------
Public Enum EditModeEnum
    adEditNone = 0       '進行中の編集操作がないことを示します。
    adEditInProgress = 1 '現在のレコードのデータが変更されたが、保存されていないことを示します。
    adEditAdd = 2        'AddNew メソッドが呼び出され、コピー バッファー内の現在のレコードが、データベースに保存されていない新しいレコードであることを示します。
    adEditDelete = 4     '現在のレコードが削除されたことを示します。
End Enum

'*-----------------------------------------------------------------------------
'* ADO 実行時エラーの種類を表します。
'* https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/errorvalueenum
'*-----------------------------------------------------------------------------
Public Enum ErrorValueEnum
    adErrBoundToCommand = 3707           'Command オブジェクトをソースに持つ Recordset オブジェクトの ActiveConnection プロパティを変更できません。
    adErrCannotComplete = 3732           'サーバーは操作を完了できません。
    adErrCantChangeConnection = 3748     '接続が拒否されました。 要求された新規接続の特性が現在使用中の特性と異なります。
    adErrCantChangeProvider = 3220       '指定されたプロバイダーが既に使用されているものと異なります。
    adErrCantConvertvalue = 3724         '符号の不一致またはデータ オーバーフロー以外の理由により、データ値を変換できません。 たとえば、変換によりデータの一部が切り捨てられる場合などです。
    adErrCantCreate = 3725               'フィールドのデータ型が不明であったか、プロバイダーが操作を実行するのに十分なリソースを持っていなかったため、データ値を設定または取得できません。
    adErrCatalogNotSet = 3747            '操作には有効な ParentCatalog が必要です。
    adErrColumnNotOnThisRow = 3726       'レコードにこのフィールドが存在しません。
    adErrConnectionStringTooLong = 3754
    adErrDataConversion = 3421           '現在の操作に対して、間違った型の値を使用しています。
    adErrDataOverflow = 3721             'データ値が大きすぎるために、フィールドのデータ型で表現できません。
    adErrDelResOutOfScope = 3738         '削除されるオブジェクトの URL は現在のレコードの範囲外です。
    adErrDenyNotSupported = 3750         'プロバイダーが共有の制約をサポートしていません。
    adErrDenyTypeNotSupported = 3751     'プロバイダーが、要求された種類の共有の制約をサポートしていません。
    adErrFeatureNotAvailable = 3251      'オブジェクトまたはプロバイダーは要求された操作を実行できません。
    adErrFieldsUpdateFailed = 3749       'フィールドを更新できませんでした。 詳細については、各 Field オブジェクトの Status プロパティを参照してください。
    adErrIllegalOperation = 3219         'このコンテキストで操作は許可されていません。
    adErrIntegrityViolation = 3719       'データの値がフィールドの整合性制約に反しています。
    adErrInTransaction = 3246            'トランザクションの実行中に Connection オブジェクトを明示的に閉じることができません。
    adErrInvalidArgument = 3001          '間違った種類または許容範囲外の引数を使用しているか、使用している引数が競合しています。
    adErrInvalidConnection = 3709        'この操作を実行するために接続を使用できません。 このコンテキストで閉じているかあるいは無効です。
    adErrInvalidParamInfo = 3708         'Parameter オブジェクトが適切に定義されていません。 矛盾した、または不完全な情報が指定されました。
    adErrInvalidTransaction = 3714       '調整トランザクションが無効であるか、開始されていません。
    adErrInvalidURL = 3729               'URL に無効な文字が含まれています。 URL が正しく入力されているか確認してください。
    adErrItemNotFound = 3265             '要求された名前、または序数に対応する項目がコレクションで見つかりません。
    adErrNoCurrentRecord = 3021          'BOF または EOF が True であるか、現在のレコードが削除されています。 要求された操作には現在のレコードが必要です。
    adErrNotReentrant = 3710             'イベント処理中に操作を行うことはできません。
    adErrObjectClosed = 3704             'オブジェクトが閉じている場合は、操作は許可されません。
    adErrObjectInCollection = 3367       'オブジェクトは既にコレクションに存在します。 追加できません。
    adErrObjectNotSet = 3420             'オブジェクトが無効になっています。
    adErrObjectOpen = 3705               'オブジェクトが開いている場合は、操作は許可されません。
    adErrOpeningFile = 3002              'ファイルを開くことができませんでした。
    adErrOperationCancelled = 3712       'ユーザーにより操作が取り消されました。
    adErrOutOfSpace = 3734               '操作を実行できません。 プロバイダーによって十分な記憶域が確保できません。
    adErrPermissionDenied = 3720         '権限不足のためフィールドの書き込みはできません。
    adErrPropConflicting = 3742
    adErrPropInvalidColumn = 3739
    adErrPropInvalidOption = 3740
    adErrPropInvalidValue = 3741
    adErrPropNotAllSettable = 3743
    adErrPropNotSet = 3744
    adErrPropNotSettable = 3745
    adErrPropNotSupported = 3746
    adErrProviderFailed = 3000           'プロバイダーが要求された操作を実行できませんでした。
    adErrProviderNotFound = 3706         'プロバイダーが見つかりません。 正しくインストールされていない可能性があります。
    adErrProviderNotSpecified = 3753
    adErrReadFile = 3003                 'ファイルを読み込むことができませんでした。
    adErrResourceExists = 3731           'コピー操作を実行できません。 宛先の URL で指定されたオブジェクトは既に存在します。 オブジェクトを置き換えるためには adCopyOverwrite を指定してください。
    adErrResourceLocked = 3730           '指定された URL によって表されたオブジェクトは 1 つ以上の他のプロセスによってロックされています。プロセスが終了するまで待って、操作を再度実行してください。
    adErrResourceOutOfScope = 3735       'ソースまたは宛先の URL が、現在のレコードの範囲外です。
    adErrSchemaViolation = 3722          'データ値がフィールドのデータ型と一致していないか、フィールドの制約に反しています。
    adErrSignMismatch = 3723             'データの値は符号付きですが、プロバイダーによって使用されるフィールド データ型は符号なしのため、変換に失敗しました。
    adErrStillConnecting = 3713          '非同期操作の保留中に、操作を行うことはできません。
    adErrStillExecuting = 3711           '非同期実行中に操作を行うことはできません。
    adErrTreePermissionDenied = 3728     '権限が不十分なために、ツリーまたはサブツリーにアクセスできません。
    adErrUnavailable = 3736              '操作の完了に失敗し、状態は利用できません。 フィールドが利用できないか、操作が実行されなかった可能性があります。
    adErrUnsafeOperation = 3716          'このコンピューターの安全性の設定により、他のドメインのデータ ソースへのアクセスが禁止されています。
    adErrURLDoesNotExist = 3727          'ソース URL または宛先の URL の親が存在しません。
    adErrURLNamedRowDoesNotExist = 3737  'この URL によって名前を付けられたレコードが存在しません。
    adErrVolumeNotFound = 3733           'プロバイダーが、URL で示された記憶装置の場所を特定できません。 URL が正しく入力されているか確認してください。
    adErrWriteFile = 3004                'ファイルへの書き込みに失敗しました。
    adwrnSecurityDialog = 3717           '内部使用のために用意されています。 使用しないでください。
    adwrnSecurityDialogHeader = 3718     '内部使用のために用意されています。 使用しないでください。
End Enum

'*-----------------------------------------------------------------------------
'* イベントが発生した理由を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum EventReasonEnum
    adRsnAddNew = 1        '新しいレコードが追加されました。
    adRsnClose = 9         'Recordset が閉じられました。
    adRsnDelete = 2        'レコードが削除されました。
    adRsnFirstChange = 11  'レコードに初めての変更が加えられました。
    adRsnMove = 10         'Recordset 内のレコード ポインターが移動しました。
    adRsnMoveFirst = 12    'レコード ポインターが Recordset の最初のレコードに移動しました。
    adRsnMoveLast = 15     'レコード ポインターが Recordset の最後のレコードに移動しました。
    adRsnMoveNext = 13     'レコード ポインターが Recordset の次のレコードに移動しました。
    adRsnMovePrevious = 14 'レコード ポインターが Recordset の前のレコードに移動しました。
    adRsnRequery = 7       'Recordset が再クエリされました。
    adRsnResynch = 8       'Recordset がデータベースと再同期しました。
    adRsnUndoAddNew = 5    '新しいレコードの追加が取り消されました。
    adRsnUndoDelete = 6    'レコードの削除が取り消されました。
    adRsnUndoUpdate = 4    'レコードの更新が取り消されました。
    adRsnUpdate = 3        '既存のレコードが更新されました。
End Enum

'*-----------------------------------------------------------------------------
'* イベントの実行の現在の状態を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum EventStatusEnum
    adStatusCancel = 4         'イベントを発生させた操作の取り消しを要求します。
    adStatusCantDeny = 3       '保留中の操作の取り消しを要求できないことを示します。
    adStatusErrorsOccurred = 2 'イベントを発生させた操作がエラーによって失敗したことを示します。
    adStatusOK = 1             'イベントを発生させた操作が成功したことを示します。
    adStatusUnwantedEvent = 5  'イベント メソッドの実行が終了するまで、後続の通知が行われません。
End Enum

'*-----------------------------------------------------------------------------
'* プロバイダーによるコマンドの実行方法を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum ExecuteOptionEnum
    adAsyncExecute = 16          'コマンドを非同期に実行することを示します。 この値は、CommandTypeEnum の値 adCmdTableDirect と組み合わせて使用できません。
    adAsyncFetch = 32            'CacheSize プロパティで指定した初期量の残りの行を非同期に取得することを示します。
    adAsyncFetchNonBlocking = 64 '取得中にメイン スレッドがブロックしないことを示します。 要求された行がまだ取得されていない場合、現在の行が自動的にファイルの最後に移動します。永続的に保存された Recordset を持つ Stream から Recordset を開いた場合、adAsyncFetchNonBlocking は無効になり、操作は同期で実行され、ブロッキングが発生します。 adCmdTableDirect オプションを使用して Recordset を開いた場合、adAsynchFetchNonBlocking は無効になります。
    adExecuteNoRecords = 128     'コマンド テキストが、行を返さないコマンドまたはストアド プロシージャ (たとえば、データの挿入のみを行うコマンド) であることを示します。 取得した行があっても削除されるので、コマンドからは返されません。 adExecuteNoRecordsは、コマンドまたはConnectionのExecuteメソッドに、省略可能なパラメーターとしてのみ渡すことができます。
    adExecuteStream = 1024       'コマンドの実行結果がストリームとして返されることを示します。 adExecuteStreamは、 Command Executeメソッドにオプションのパラメーターとして渡すことができます。
    adExecuteRecord = 512        'Indicates that the CommandText is a command or stored procedure that returns a single row which should be returned as a Record object.
    adOptionUnspecified = -1     'Indicates that the command is unspecified.
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
'* Record オブジェクトの Fields コレクションで参照される特定のフィールドを
'* 表します。
'*-----------------------------------------------------------------------------
Public Enum FieldEnum
    adDefaultStream = -1 'Record に関連付けられた既定の Stream オブジェクトを含むフィールドを参照します。
    adRecordURL = -2     '現在の Record の絶対 URL 文字列を含むフィールドを参照します。
End Enum

'*-----------------------------------------------------------------------------
'* Field オブジェクトの状態を表します。
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
'* Recordset でフィルターの対象となるレコード グループを表します。
'*
'*-----------------------------------------------------------------------------
Public Enum FilterGroupEnum
    adFilterAffectedRecords = 2    '最後に行った Delete、Resync、UpdateBatch、または CancelBatch 呼び出しで影響を受けたレコードのみを表示するようにフィルター処理します。
    adFilterConflictingRecords = 5 '最後に行ったバッチ更新が失敗したレコードを表示するようにフィルター処理します。
    adFilterFetchedRecords = 3     'データベースから最後に取得されたレコードである現在のキャッシュ内のレコードを表示するようにフィルター処理します。
    adFilterNone = 0               '現在のフィルターを削除し、すべてのレコードを復元して表示します。
    adFilterPendingRecords = 1     '変更が行われたが、変更内容がサーバーにまだ送信されていないレコードのみを表示するようにフィルター処理します。バッチ更新モードの場合のみ使用できます。
End Enum

'*-----------------------------------------------------------------------------
'* Recordset から取得するレコード数を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum GetRowsOptionEnum
    adGetRowsRest = -1 '現在の位置または GetRows メソッドの Start パラメーターで指定されたブックマークから、Recordset 内の残りのレコードを取得します。
End Enum

'*-----------------------------------------------------------------------------
'* Connection オブジェクトのトランザクション分離レベルを表します
'*
'*-----------------------------------------------------------------------------
Public Enum IsolationLevelEnum
    adXactUnspecified = -1       'プロバイダーが指定されたものとは異なる分離レベルを使用していますが、レベルを特定できないことを示します。
    adXactChaos = 16             '分離度の高いトランザクションからの保留中の変更を上書きできないことを示します。
    adXactBrowse = 256           '1つのトランザクションから、他のトランザクションのコミットされていない変更を表示できることを示します。
    adXactReadUncommitted = 256  'adXactBrowseと同じです。
    adXactCursorStability = 4096 '1つのトランザクションから、コミットされた後にのみ他のトランザクションの変更を表示できることを示します。
    adXactReadCommitted = 4096   'adXactCursorStabilityと同じです。
    adXactRepeatableRead = 65536 '1つのトランザクションから、他のトランザクションで行われた変更を表示できないが、再クエリで新しいRecordsetオブジェクトを取得できることを示します。
    adXactIsolated = 1048576     'トランザクションが他のトランザクションとは分離して実行されることを示します。
    adXactSerializable = 1048576 'adXactIsolatedと同じです。
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
'* 編集時にレコードに適用されるロックの種類を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum LockTypeEnum
    adLockBatchOptimistic = 4 '共有的バッチ更新を示します。バッチ更新モードの場合に必要です。
    adLockOptimistic = 3      'レコード単位の共有的ロックを示します。Update メソッドを呼び出した場合にのみ、プロバイダーは共有的ロックを使ってレコードをロックします。
    adLockPessimistic = 2     'レコード単位の排他的ロックを示します。プロバイダーは、レコードを確実に編集するための措置を行います。通常は、編集直後にデータ ソースでレコードをロックします。
    adLockReadOnly = 1        '読み取り専用のレコードを示します。データの変更はできません。
    adLockUnspecified = -1    'ロックの種類を指定しません。複製の場合、複製元と同じロックの種類が適用されます。
End Enum

'*-----------------------------------------------------------------------------
'* サーバーにどのレコードが返されるかを表します。
'*
'*-----------------------------------------------------------------------------
Public Enum MarshalOptionsEnum
    adMarshalAll = 0          '既定値。すべての行をサーバーに返します。
    adMarshalModifiedOnly = 1 '変更した行のみサーバーに返します。
End Enum

'*-----------------------------------------------------------------------------
'* Record オブジェクトの MoveRecord メソッドの動作を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum MoveRecordOptionsEnum
    adMoveUnspecified = -1    '既定値。既定の移動操作を実行します。既存の宛先ファイルまたはディレクトリがある場合は操作が失敗し、ハイパーテキスト リンクが更新されます。
    adMoveOverWrite = 1       '既存の宛先ファイルまたはディレクトリがあっても上書きします。
    adMoveDontUpdateLinks = 2 'ソース Record のハイパーテキスト リンクを更新しないことで、MoveRecord メソッドの既定動作を変更します。既定動作はプロバイダーの機能によって異なります。プロバイダーがサポートしていれば、移動操作でリンクが更新されます。プロバイダーがリンクの修正をサポートしていない場合、またはこの値が指定されていない場合、リンクを修正しなくても移動は成功します。
    adMoveAllowEmulation = 4  'プロバイダーによる移動 (ダウンロード、アップロード、削除の操作を使用) のシミュレーションを要求します。宛先 URL がソースとは別のサーバーにあったり、別のプロバイダーがサービスを提供しているために Record の移動が失敗すると、プロバイダー間でリソースを移動するときのプロバイダーの機能の違いにより、遅延時間の増加やデータの損失が起きることがあります。
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
'* 権限または所有権を設定するデータベース オブジェクトの種類を示します。
'*
'*-----------------------------------------------------------------------------
Public Enum ObjectTypeEnum
    adPermObjColumn = 2             'オブジェクトは列です。
    adPermObjDatabase = 3           'オブジェクトはデータベースです。
    adPermObjProcedure = 4          'オブジェクトはプロシージャです。
    adPermObjProviderSpecific = -1  'オブジェクトの種類は、プロバイダー によって定義されます。ObjectType パラメーターが adPermObjProviderSpecific で、ObjectTypeId が指定されていない場合、エラーが発生します。
    adPermObjTable = 1              'オブジェクトはテーブルです。
    adPermObjView = 5               'オブジェクトはビューです。
End Enum


'*-----------------------------------------------------------------------------
'* Parameter オブジェクトの属性を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum ParameterAttributesEnum
    adParamLong = 128    'パラメーターにロング バイナリ型のデータを指定できることを示します。
    adParamNullable = 64 'パラメーターに null 値を指定できることを示します。
    adParamSigned = 16   'パラメーターに符号付きの値を指定できることを示します。
End Enum

'*-----------------------------------------------------------------------------
'* Parameter が入力パラメーターと出力パラメーターのいずれか、またはその両方を
'* 表すのか、あるいはストアド プロシージャからの戻り値であるかを表します。
'*-----------------------------------------------------------------------------
Public Enum ParameterDirectionEnum
    adParamInput = 1       '既定値。パラメーターが入力パラメーターを表すことを示します。
    adParamInputOutput = 3 'パラメーターが入力パラメーターと出力パラメーターの両方を表すことを示します。
    adParamOutput = 2      'パラメーターが出力パラメーターを表すことを示します。
    adParamReturnValue = 4 'パラメーターが戻り値を表すことを示します。
    adParamUnknown = 0     'パラメーターの方向が不明であることを示します。
End Enum

'*-----------------------------------------------------------------------------
'* Recordset を保存するときの形式を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum PersistFormatEnum
    adPersistADTG = 0             'Microsoft Advanced Data TableGram (ADTG) 形式であることを示します。
    adPersistXML = 1              '拡張マークアップ言語 (XML) 形式であることを示します。
    adPersistADO = 1              'Indicates that ADO's own Extensible Markup Language (XML) format will be used. This value is the same as adPersistXML and is included for backwards compatibility.
    adPersistProviderSpecific = 2 'Indicates that the provider will persist the Recordset using its own format.
End Enum

'*-----------------------------------------------------------------------------
'* Recordset 内のレコード ポインターの現在の位置を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum PositionEnum
    adPosBOF = -2     '現在のレコード ポインターが BOF にあることを示します (BOF プロパティが True です)。
    adPosEOF = -3     '現在のレコード ポインターが EOF にあることを示します (EOF プロパティが True です)。
    adPosUnknown = -1 'Recordset が空であるか、現在の位置が不明か、またはプロバイダーが AbsolutePage プロパティまたは AbsolutePosition プロパティをサポートしていないことを示します。
End Enum

'*-----------------------------------------------------------------------------
'* Property オブジェクトの属性を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum PropertyAttributesEnum
    adPropNotSupported = 0 'プロバイダーがプロパティをサポートしていないことを示します。
    adPropRequired = 1     'データ ソースを初期化するには、ユーザーがこのプロパティ値を指定する必要があることを示します。
    adPropOptional = 2     'ユーザーがこのプロパティ値を指定しなくてもデータ ソースを初期化できることを表します。
    adPropRead = 512       'ユーザーがプロパティを読み取り可能であることを示します。
    adPropWrite = 1024     'ユーザーがプロパティを設定できることを示します。
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
'* バッチ更新またはその他の一括操作に関するレコードの状態を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum RecordStatusEnum
    adRecCanceled = 256              '操作が取り消されたため、レコードが保存されなかったことを示します。
    adRecCantRelease = 1024          '既存のレコードがロックされていたため、新しいレコードが保存されなかったことを示します。
    adRecConcurrencyViolation = 2048 'オプティミスティック同時実行制御が使用されていたため、レコードが保存されなかったことを示します。
    adRecDBDeleted = 262144          'レコードは既にデータ ソースから削除されていることを示します。
    adRecDeleted = 4                 'レコードが削除されたことを示します。
    adRecIntegrityViolation = 4096   'ユーザーが整合性制約に違反したため、レコードが保存されなかったことを示します。
    adRecInvalid = 16                'ブックマークが無効なため、レコードが保存されなかったことを示します。
    adRecMaxChangesExceeded = 8192   '保留中の変更が多すぎたため、レコードが保存されなかったことを示します。
    adRecModified = 2                'レコードが変更されたことを示します。
    adRecMultipleChanges = 64        '複数のレコードに影響が及ぶため、レコードが保存されなかったことを示します。
    adRecNew = 1                     'レコードが新しいことを示します。
    adRecObjectOpen = 16384          '開いているストレージ オブジェクトとの競合のため、レコードが保存されなかったことを示します。
    adRecOK = 0                      'レコードが正常に更新されたことを示します。
    adRecOutOfMemory = 32768         'メモリ不足のためにレコードが保存されなかったことを示します。
    adRecPendingChanges = 128        '保留中の挿入を参照しているため、レコードが保存されなかったことを示します。
    adRecPermissionDenied = 65536    'ユーザーの権限不足により、レコードが保存されなかったことを示します。
    adRecSchemaViolation = 131072    '基になるデータベースの構造に違反するので、レコードが保存されなかったことを示します。
    adRecUnmodified = 8              'レコードが変更されなかったことを示します。
End Enum

'*-----------------------------------------------------------------------------
'* Record オブジェクトの種類を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum RecordTypeEnum
    adCollectionRecord = 1 '"コレクション" レコード (子ノードがあるレコード) を示します。
    adSimpleRecord = 0     '"単純" レコード (子ノードがないレコード) を示します。
    adStructDoc = 2        'COM 構造化ドキュメントを表す特殊な "コレクション" レコードを示します。
    adRecordUnknown = -1   'Indicates that the type of this Record is unknown.
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
'* Stream オブジェクトからファイルに保存するときに、ファイルを作成するか、
'* 上書きするかを表します。これらの値は AND 演算子で結合できます。
'*-----------------------------------------------------------------------------
Public Enum SaveOptionsEnum
    adSaveCreateNotExist = 1  '既定値。FileName パラメーターで指定したファイルがない場合は新しいファイルが作成されます。
    adSaveCreateOverWrite = 2 'Filename パラメーターで指定したファイルがある場合は、現在開かれている Stream オブジェクトのデータでファイルが上書きされます。
End Enum

'*-----------------------------------------------------------------------------
'* OpenSchema メソッドが取得するスキーマ Recordset の種類を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum SchemaEnum
    adSchemaActions = 41               '
    adSchemaAsserts = 0                'カタログに定義され、所定のユーザーが所有するアサーションを返します。 (ASSERTIONS 行セット)
    adSchemaCatalogs = 1               'DBMS からアクセスできるカタログに関連付けられている物理的属性を返します。 (CATALOGS 行セット)
    adSchemaCharacterSets = 2          'カタログに定義され、所定のユーザーがアクセスできる文字セットを返します。 (CHARACTER_SETS 行セット)
    adSchemaCheckConstraints = 5       'カタログに定義され、所定のユーザーが所有する CHECK 制約を返します。 (CHECK_CONSTRAINTS 行セット)
    adSchemaCollations = 3             'カタログに定義され、所定のユーザーがアクセスできる文字照合順序を返します。 (COLLATIONS 行セット)
    adSchemaColumnPrivileges = 13      'カタログに定義され、所定のユーザーが利用できる、または権限を持つテーブルの列に対する特権を返します。 (COLUMN_PRIVILEGES 行セット)
    adSchemaColumns = 4                'カタログに定義され、所定のユーザーがアクセスできるテーブルの列 (ビューも含む) を返します。 (COLUMNS 行セット)
    adSchemaColumnsDomainUsage = 11    'カタログに定義され、そのカタログに定義されたドメインに依存し、所定のユーザーが所有する列を返します。 (COLUMN_DOMAIN_USAGE 行セット)
    adSchemaCommands = 42              '
    adSchemaConstraintColumnUsage = 6  'カタログに定義され、所定のユーザーが所有し、参照制約、一意制約、CHECK 制約、およびアサーションに使う列を返します。 (CONSTRAINT_COLUMN_USAGE 行セット)
    adSchemaConstraintTableUsage = 7   'カタログに定義され、所定のユーザーが所有し、参照制約、一意制約、CHECK 制約、およびアサーションに使うテーブルを返します。 (CONSTRAINT_TABLE_USAGE 行セット)
    adSchemaCubes = 32                 'スキーマ (プロバイダーがスキーマをサポートしていない場合はカタログ) 内の利用できるキューブに関する情報を返します。 (CUBES 行セット \*)
    adSchemaDBInfoKeywords = 30        'プロバイダー固有のキーワードの一覧を返します。 (IDBInfo:: getkeywords \*)
    adSchemaDBInfoLiterals = 31        'テキスト コマンドで使う、プロバイダー固有のリテラルの一覧を返します。 (IDBInfo:: GetLiteralInfo \*)
    adSchemaDimensions = 33            '所定のキューブの次元に関する情報を返します。 次元ごとに 1 行が割り当てられます。 (DIMENSIONS 行セット \*)
    adSchemaForeignKeys = 27           '所定のユーザーがカタログに定義した外部キー列を返します。 (FOREIGN_KEYS 行セット)
    adSchemaFunctions = 40             '
    adSchemaHierarchies = 34           '次元で利用できる階層に関する情報を返します。 (HIERARCHIES 行セット \*)
    adSchemaIndexes = 12               'カタログに定義され、所定のユーザーが所有するインデックスを返します。 (INDEXES 行セット)
    adSchemaKeyColumnUsage = 8         'カタログに定義され、所定のユーザーがキーとして制約した列を返します。 (KEY_COLUMN_USAGE 行セット)
    adSchemaLevels = 35                '次元で利用できるレベルに関する情報を返します。 (LEVELS 行セット \*)
    adSchemaMeasures = 36              '利用できる単位に関する情報を返します。 (MEASURES 行セット \*)
    adSchemaMembers = 38               '利用できるメンバーに関する情報を返します。 (MEMBERS 行セット \*)
    adSchemaPrimaryKeys = 28           '所定のユーザーがカタログに定義した主キー列を返します。 (PRIMARY_KEYS 行セット)
    adSchemaProcedureColumns = 29      'プロシージャが返す行セットの列に関する情報を返します。 (PROCEDURE_COLUMNS Rowset)
    adSchemaProcedureParameters = 26   'プロシージャのパラメーターとリターン コードに関する情報を返します。 (PROCEDURE_PARAMETERS 行セット)
    adSchemaProcedures = 16            'カタログに定義され、所定のユーザーが所有するプロシージャを返します。 (PROCEDURES 行セット)
    adSchemaProperties = 37            '次元の各レベルで利用できるプロパティに関する情報を返します。 (PROPERTIES 行セット \*)
    adSchemaProviderSpecific = -1      'プロバイダーが非標準の専用のスキーマ クエリを定義する場合に使います。
    adSchemaProviderTypes = 22         'データ プロバイダーがサポートする (基本) データ型を返します。 (PROVIDER_TYPES 行セット)
    adSchemaReferentialConstraints = 9 'カタログに定義され、所定のユーザーが所有する参照制約を返します。 (REFERENTIAL_CONSTRAINTS 行セット)
    adSchemaSchemata = 17              '所定のユーザーが所有するスキーマ (データベース オブジェクト) を返します。 (SCHEMATA 行セット)
    adSchemaSets = 43                  '
    adSchemaSQLLanguages = 18          'カタログに定義された SQL 実装処理データがサポートする準拠レベル、オプション、および言語を返します。 (SQL_LANGUAGES 行セット)
    adSchemaStatistics = 19            'カタログに定義され、所定のユーザーが所有する統計値を返します。 (STATISTICS 行セット)
    adSchemaTableConstraints = 10      'カタログに定義され、所定のユーザーが所有するテーブル制約を返します。 (TABLE_CONSTRAINTS 行セット)
    adSchemaTablePrivileges = 14       'カタログに定義され、所定のユーザーが利用できる、または権限を持つテーブルに対する特権を返します。 (TABLE_PRIVILEGES 行セット)
    adSchemaTables = 20                'カタログに定義され、所定のユーザーがアクセスできるテーブル (ビューも含む) を返します。 (TABLES 行セット)
    adSchemaTranslations = 21          'カタログに定義され、所定のユーザーがアクセスできる文字変換を返します。 (TRANSLATIONS 行セット)
    adSchemaTrustees = 39              '将来使用するために予約されています。
    adSchemaUsagePrivileges = 15       'カタログに定義され、所定のユーザーが利用できる、または権限を持つオブジェクトに対する USAGE 特権を返します。 (USAGE_PRIVILEGES 行セット)
    adSchemaViewColumnUsage = 24       'カタログに定義され、所定のユーザーが所有する、表示テーブルが依存する列を返します。 (VIEW_COLUMN_USAGE 行セット)
    adSchemaViews = 23                 'カタログに定義され、所定のユーザーがアクセスできるビューを返します。 (VIEWS 行セット)
    adSchemaViewTableUsage = 25        'カタログに定義され、所定のユーザーが所有し、表示テーブルが依存するテーブルを返します。 (VIEW_TABLE_USAGE 行セット)
End Enum

'*-----------------------------------------------------------------------------
'* Recordset 内のレコードの検索方向を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum SearchDirectionEnum
    adSearchBackward = -1 '後方検索をし、Recordset の先頭で終了します。一致するレコードが見つからない場合、レコード ポインターは BOF に移動します。
    adSearchForward = 1   '前方検索をし、Recordset の末尾で終了します。一致するレコードが見つからない場合、レコード ポインターは EOF に移動します。
End Enum

'*-----------------------------------------------------------------------------
'* 隠し項目：Recordset 内のレコードの検索方向を表します。
'* コメントアウト。
'*-----------------------------------------------------------------------------
'Public Enum SearchDirection
'    adSearchBackward = -1
'    adSearchForward = 1
'End Enum

'*-----------------------------------------------------------------------------
'* 実行する Seek の種類を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum SeekEnum
    adSeekFirstEQ = 1   'KeyValues と一致する最初のキーを検索します。
    adSeekLastEQ = 2    'KeyValues と一致する最後のキーを検索します。
    adSeekAfterEQ = 4   'KeyValues と一致するキー、またはその直後のキーのいずれかを検索します。
    adSeekAfter = 8     'KeyValues と一致するキーの直後のキーを検索します。
    adSeekBeforeEQ = 16 'KeyValues と一致するキー、またはその直前のキーのいずれかを検索します。
    adSeekBefore = 32   'KeyValues と一致するキーの直前のキーを検索します。
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
'* Stream オブジェクトから、ストリーム全体を読み取るか、または次の行を読み取るかを表します。
'*
'*-----------------------------------------------------------------------------
Public Enum StreamReadEnum
    adReadAll = -1  '既定値。現在の位置から EOS マーカー方向に、すべてのバイトをストリームから読み取ります。これは、バイナリ ストリーム (Type は adTypeBinary) に唯一有効な StreamReadEnum 値です。
    adReadLine = -2 'ストリームから次の行を読み取ります (LineSeparator プロパティで指定)。
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
'* Stream オブジェクトに書き込む文字列に、行区切り記号を追加するかどうかを表します。
'*
'*-----------------------------------------------------------------------------
Public Enum StreamWriteEnum
    adWriteChar = 0 '既定値。Stream オブジェクトに、Data パラメーターで指定したテキスト文字列を書き込みます。
    adWriteLine = 1 'Stream オブジェクトに、テキスト文字列と行区切り記号を書き込みます。LineSeparator プロパティが定義されていない場合は、実行時エラーを返します。
End Enum

'*-----------------------------------------------------------------------------
'* 文字列として Recordset を取得するときの形式を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum StringFormatEnum
    adClipString = 2 '行が RowDelimiter によって、列が ColumnDelimiter によって、null 値が NullExpr によって区切られます。GetString メソッドのこれらの 3 つのパラメーターは、adClipString の StringFormat とのみ併用できます。
End Enum

'*-----------------------------------------------------------------------------
'* Connection オブジェクトのトランザクション属性を表します。
'*
'*-----------------------------------------------------------------------------
Public Enum XactAttributeEnum
    adXactAbortRetaining = 262144  '中断の保持を実行します。つまり、 RollbackTransを呼び出すと、新しいトランザクションが自動的に開始されます。 この設定をサポートしていないプロバイダーもあります。
    adXactCommitRetaining = 131072 '保持コミットを実行します。つまり、 CommitTransを呼び出すと、新しいトランザクションが自動的に開始されます。 この設定をサポートしていないプロバイダーもあります。
End Enum

'******************************************************************************
'* メソッド定義
'******************************************************************************
'******************************************************************************
'* [概  要] ファイルエンコード一括変換処理。
'* [詳  細] 指定したフォルダ内のファイルのエンコードを一括変換する。
'*
'* @param targetFolderName 対象となるフォルダのフルパス
'* @param srcEncode 変更元エンコード
'* @param destEncode 変更先エンコード
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
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
         
        ' ファイルエンコード変換
        Call ChangeFileEncode(filePath, srcEncode, destEncode, bomInclude)
    Next
End Sub

'******************************************************************************
'* [概  要] ファイルエンコード変換処理。
'* [詳  細] 指定したファイルのエンコードを変換する。
'*
'* @param filePath 対象となるファイルのフルパス
'* @param srcEncode 変更元エンコード
'* @param destEncode 変更先エンコード
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub ChangeFileEncode(filePath As String, srcEncode As String, destEncode As String, _
                            Optional bomInclude As Boolean = False)
    Dim adoStream1 As ADODBExStream, adoStream2 As ADODBExStream
    Set adoStream1 = New ADODBExStream
    Set adoStream2 = New ADODBExStream
         
    ' 変更元ファイルStream読込
    With adoStream1
        .OpenStream
        .Type_ = adTypeText
        .CharSet = srcEncode
        .LoadFromFile filePath
    End With
     
    ' 変更先ファイルStream読込
    With adoStream2
        .OpenStream
        .Type_ = adTypeText
        .CharSet = destEncode
        .BOM = bomInclude
    End With
     
    ' エンコード変換
    adoStream1.CopyTo adoStream2
    adoStream2.SaveToFile filePath, adSaveCreateOverWrite 'ファイル上書指定
     
    ' Streamクローズ
    adoStream2.CloseStream
    adoStream1.CloseStream
End Sub

'******************************************************************************
'* [概  要] ファイル読込・書き込み処理。
'* [詳  細] 指定した読込ファイルのデータを別ファイルに書き込む。
'* [参  考] 大容量データの読み込みについては以下のサイトを参考にした。
'*          <https://mussyu1204.myhome.cx/wordpress/it/?p=720>
'*
'* @param srcFilePath 読込ファイルのフルパス
'* @param srcEncode 読込元エンコード
'* @param srcSep 読込元改行コード
'* @param destFilePath 書込ファイルのフルパス
'* @param destEncode 書込先エンコード
'* @param destSep 書込先改行コード
'* @param funcName 行編集処理用関数名。
'*                 以下のように引数に文字列、戻り値に文字列を返す関数名を指定。
'*                 funcName(row As String) As String
'*                 指定しない（空文字）場合は、行編集は行わない。
'* @param chunkSize チャンクサイズ。このサイズを超える読込データの場合は、
'*                  チャンクサイズごとに分割して処理を行う。
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
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
    
    ' 読込データのサイズが指定サイズより大きい場合は分割処理（高速化）実施
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

    ' ファイル保存
    outStream.SaveToFile destFilePath, adSaveCreateOverWrite
     
    inStream.CloseStream
    outStream.CloseStream
End Sub

'******************************************************************************
'* [概  要] ディレクトリパス分離符付与処理。
'* [詳  細] ディレクトリパスの末尾に分離符（￥）がなければ付与を行う。
'*
'* @param strDirPath ディレクトリパス
'* @return 分離符付きディレクトリパス
'******************************************************************************
Private Function AddPathSeparator(strDirPath As String) As String
    If Right(strDirPath, 1) <> "\" Then
        AddPathSeparator = strDirPath & "\"
    Else
        AddPathSeparator = strDirPath
    End If
End Function
