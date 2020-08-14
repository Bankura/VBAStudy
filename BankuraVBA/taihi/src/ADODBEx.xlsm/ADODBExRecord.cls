VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODB.Recordクラスのラップ・拡張クラス
'* [詳  細] ADODB.RecordクラスのWrapperとしての機能を提供する他、
'*          ADODB.Recordクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Recordクラスについて。
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Represents a single row of data, either from a Recordset or from the provider.
' This record could represent a database record or some other type of object
' such as a file or directory, depending upon your provider.
'
' レコードセットまたはプロバイダーからのデータの単一行を表す。
' このレコードは、プロバイダーに応じて、データベースレコードまたはファイルや
' ディレクトリなどの他のタイプのオブジェクトを表す場合がある。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Record
Private mOrigin As Object
    
'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* ActiveConnection プロパティ
'*
'*   ※オブジェクトブラウザでCommandのActiveConnection プロパティを確認すると
'*     型はConnectionとなっているが、接続情報文字列（String）も受け取ることが
'*     可能のため、Variant型としています。
'*
'* [補  足]
'* ・ADODBExConnection/Connectionオブジェクト、または接続情報文字列（String）
'*   を設定する。
'* ・Nothingを設定すると、Commandオブジェクトに現在設定されているConnection
'*   が解放される。
'* ・接続が閉じている場合には接続の定義が格納された文字列型 ( String ) の値
'*   を、接続が開いている場合には現在の Connection オブジェクトが格納された
'*   バリアント型 ( Variant ) の値を設定または取得する。
'*   既定は、Null オブジェクト参照。
'*   ConnectionString プロパティの説明を参照すること。
'* ・このプロパティは、Record オブジェクトが閉じている場合には値の設定および
'*   取得が可能で、接続文字列または開いている Connection オブジェクトの参照を
'*   格納できる。
'*   Record オブジェクトが開いている場合は値の取得のみ可能で、開いている
'*   Connection オブジェクトの参照が格納されている。
'* ・Connection オブジェクトは、URL から Record オブジェクトが開かれたときに
'*   暗黙的に作成される。
'*   既存の、開いている Connection オブジェクトで Record を開くには、
'*   Connection オブジェクトをこのプロパティに代入するか、または Connection
'*   オブジェクトを Open メソッド呼び出しのパラメーターとして使用する。
'*   recordを既存のrecordまたはrecordsetから開いた場合は、そのレコードまたは
'*   recordsetオブジェクトのConnectionオブジェクトに自動的に関連付けられる。
'*-----------------------------------------------------------------------------
Property Get ActiveConnection() As Variant
    If IsObject(mOrigin.ActiveConnection) Then
        Dim wrap As ADODBExConnection: Set wrap = New ADODBExConnection
        Set wrap.Origin = mOrigin.ActiveConnection
        Set ActiveConnection = wrap
    Else
        ActiveConnection = mOrigin.ActiveConnection
    End If
End Property
Property Let ActiveConnection(arg As Variant)
    If IsObject(arg) Then
        If TypeName(arg) = "ADODBExConnection" Then
            Set mOrigin.ActiveConnection = arg.Origin
        Else
            Set mOrigin.ActiveConnection = arg
        End If
    Else
        If TypeName(arg) = "Nothing" Then
            Set mOrigin.ActiveConnection = Nothing
        Else
            mOrigin.ActiveConnection = arg
        End If
    End If
End Property

'*-----------------------------------------------------------------------------
'* Fields プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・１つ以上の列（フィールド）を表すオブジェクトを返却する。
'*-----------------------------------------------------------------------------
Property Get Fields() As ADODBExFields
Attribute Fields.VB_UserMemId = 0
    Dim wrap As ADODBExFields: Set wrap = New ADODBExFields
    Set wrap.Origin = mOrigin.Fields
    Set Fields = wrap
End Property

'*-----------------------------------------------------------------------------
'* Mode プロパティ
'*
'* [補  足]
'* ・使用可能なデータ変更権限を示す。
'*   Record オブジェクトの既定値は adModeRead 。
'* ・このプロパティは、オブジェクトが閉じているときは読み取り/書き込み可能で、
'*   オブジェクトが開いているときは読み取り専用になる。
'*-----------------------------------------------------------------------------
Property Get Mode() As ConnectModeEnum
    Mode = mOrigin.Mode
End Property
Property Let Mode(arg As ConnectModeEnum)
    mOrigin.Mode = arg
End Property

'*-----------------------------------------------------------------------------
'* ParentURL プロパティ
'*
'* [補  足]
'* ・現在の Record オブジェクトの親 Record を示す絶対 URL 文字列を示す。
'* ・ParentURL プロパティは、 Record オブジェクトを開くために使用されるソース
'*   に依存する。
'*   たとえば、 ActiveConnection プロパティによって参照されるディレクトリの相対
'*   パスを含むソースを使用して Record を開くことができる。
'* ・ここで、"second" が "first" の下のフォルダーであるとする。
'*   Record オブジェクトを開くには、次のようにする。
'*       Record.ActiveConnection = "https://first"
'*       Record.Open "second"
'*   この時点で、 parenturlプロパティの値は "https://first" で、
'*   ActiveConnectionと同じである。
'* ・ソースには、"https://first/second" などの絶対 URL を指定することもできる。
'*   parenturlプロパティは、上記のhttps://firstレベルの "" である。
'*   parenturlプロパティは、"https://first"、上の "秒" のレベルである。
'* ・次のような場合、このプロパティは Null 値になる。
'*   ・現在のレコードに親がない場合 ( Record オブジェクトがディレクトリのルート
'*     を表す場合など)
'*   ・Record オブジェクトが URL で指定できないエンティティである場合
'* ・このプロパティは値の取得のみ可能｡
'* ・[!メモ] このプロパティは、Microsoft OLE DB Provider for Internet Publishing
'*   などのドキュメント ソース プロバイダーでのみサポートされる。
'* ・[!メモ] http スキームを使用している URL は、Microsoft OLE DB Provider for
'*   Internet Publishing を自動的に呼び出す。
'* ・[!メモ] 現在のレコードに ADO Recordset のデータ レコードが含まれている
'*   場合、 ParentURL プロパティにアクセスすると、URL を取得できないことを示す
'*   実行時エラーが発生する。
'*-----------------------------------------------------------------------------
Property Get ParentURL() As String
    ParentURL = mOrigin.ParentURL
End Property
Property Let ParentURL(arg As String)
    mOrigin.ParentURL = arg
End Property

'*-----------------------------------------------------------------------------
'* Properties プロパティ
'*
'* [補  足]
'* ・Commandに設定されているPropertiesオブジェクトを取得する。
'*-----------------------------------------------------------------------------
Property Get Properties() As ADODBExProperties
    Dim wrap As ADODBExProperties: Set wrap = New ADODBExProperties
    Set wrap.Origin = mOrigin.Properties
    Set Properties = wrap
End Property
Property Set Properties(arg As ADODBExProperties)
    Set mOrigin.Properties = arg.Origin
End Property

'*-----------------------------------------------------------------------------
'* RecordType プロパティ
'*
'* [補  足]
'* ・Record オブジェクトの種類を示す。
'*-----------------------------------------------------------------------------
Property Get RecordType() As RecordTypeEnum
    RecordType = mOrigin.RecordType
End Property
Property Let RecordType(arg As RecordTypeEnum)
    mOrigin.RecordType = arg
End Property

'*-----------------------------------------------------------------------------
'* Source プロパティ
'*
'* [補  足]
'* ・Record で表されるデータソースまたはオブジェクトを示す。
'* ・Source プロパティは、Record オブジェクトの Open メソッドの Source 引数を
'*   返す。
'*   このプロパティには絶対 URL または相対 URL の文字列を格納できる。
'*   絶対 URL を使用すると、 ActiveConnection プロパティを設定せずに、直接
'*   Record オブジェクトを開くことができる。
'*   この場合、暗黙の Connection オブジェクトが作成される。
'* ・Source プロパティには、既に開いている Recordset への参照も格納でき、この
'*   参照は Recordset 内の現在の行を表す Record オブジェクトを開く。
'* ・また、 Source プロパティには、プロバイダーから 1 つのデータ行を返す Command
'*   オブジェクトへの参照も格納できる。
'* ・ActiveConnection プロパティも設定する場合は、 Source プロパティはその接続
'*   範囲内に存在する同じオブジェクトを指している必要がある。たとえば、ツリ
'*   ー構造の名前空間では、 Source プロパティに絶対 URL が格納されている場合、
'*   接続文字列の URL で指定されたノードの範囲内に存在するノードを指している必要
'*   がある。
'*   Source プロパティに相対 URL が格納されている場合、 ActiveConnection プロパ
'*   ティで設定されたコンテキスト内でのみ有効になる。
'* ・Source プロパティは、 Record オブジェクトが閉じている場合は値の取得および
'*   設定が可能で、 Record オブジェクトが開いている場合は値の取得のみ可能。
'* ・[!メモ] http スキーマを使用している URL は自動的に Microsoft OLE DB
'*   Provider for Internet Publishing を呼び出す。
'*-----------------------------------------------------------------------------
Property Get Source() As Variant
    If IsObject(mOrigin.Source) Then
        Set Source = mOrigin.Source
    Else
        Source = mOrigin.Source
    End If
End Property
Property Let Source(arg As Variant)
    If IsObject(arg) Then
        Set mOrigin.Source = arg
    Else
        mOrigin.Source = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* State プロパティ（読み取り専用）
'*
'* [補  足]
'* ・オブジェクトの状態が開いているか、閉じているかを示す。
'* ・非同期メソッドを実行する対象になるすべてのオブジェクトについて､オブジェク
'*   トの状態が接続、実行、取得のいずれであるかを示す｡
'* ・ObjectStateEnum の値になる長整数型 ( Long ) の値を返す。
'*   既定値は adStateClosed 。
'* ・State プロパティを使用して､特定のオブジェクトの現在の状態をいつでも調べる
'*   ことができる｡
'* ・オブジェクトの State プロパティは、値の組み合わせになる場合がある。
'*   たとえば、ステートメントが実行中である場合、このプロパティの値は
'*   adStateOpen と adStateExecuting の組み合わせになる。
'* ・State プロパティは値の取得のみ可能｡
'*-----------------------------------------------------------------------------
Property Get State() As ObjectStateEnum
    State = mOrigin.State
End Property
Property Let State(arg As ObjectStateEnum)
    mOrigin.State = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("ADODB.Record")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Cancelメソッド
'* [詳  細] Cancelのラッパーメソッド。
'*          保留中の非同期メソッド呼び出しの実行を取り消す。
'*          非同期メソッド呼び出し (つまり、adAsyncConnect、adAsyncExecute、
'*          または adAsyncFetch オプションを指定して呼び出されたメソッド) の
'*          実行を中止する。
'*          Recordの場合、CopyRecord、DeleteRecord、MoveRecord、または Openが
'*          最後の非同期呼び出しが中止されるメソッド｡
'*
'******************************************************************************
Sub Cancel()
    Call mOrigin.Cancel
End Sub

'******************************************************************************
'* [概  要] CloseRメソッド
'* [詳  細] Closeのラッパーメソッド。
'*          開いているオブジェクトおよびそれに関連するすべてのオブジェクトを
'*          閉じる。
'*
'*          関連するすべてのシステム リソースを解放する場合に使用する。
'*          オブジェクトを閉じてもメモリからは削除されず、プロパティ設定を変更
'*          してもう一度開くことができる。
'*          オブジェクトをメモリから完全に削除するには、オブジェクトを閉じた後
'*          に、オブジェクト変数をNothingに設定する。
'*          オブジェクトを閉じると、関連するデータ、およびデータに対するその
'*          オブジェクトからの排他アクセスが、すべて解放される。
'*          後で Open メソッドを呼び出し、同じ属性で、または属性を変更して、
'*          オブジェクトを再度開くことができる。
'*
'******************************************************************************
Sub CloseR()
    Call mOrigin.Close
End Sub

'******************************************************************************
'* [概  要] CopyRecordメソッド
'* [詳  細] CopyRecordのラッパーメソッド。
'*          Record で表されるエンティティを別の場所にコピーする。
'*          Source と Destination の値は異なっている必要がある。
'*          値が等しい場合、実行時エラーが発生する。
'*          サーバー名、パス名、リソース名のうち、少なくとも 1 つが異なっている
'*          必要がある。
'*          Source の子 (サブディレクトリなど) はすべて、adCopyNonRecursive が
'*          指定されている場合を除き、そのとおりにコピーされる。
'*          サブディレクトリもコピーする場合は、Destination に Source のサブデ
'*          ィレクトリを指定しないこと。
'*          指定した場合はコピー操作が完了しない。
'*          Destination に既存のエンティティ (ファイル、ディレクトリなど) を
'*          指定する場合、adCopyOverWrite を指定していなければ、このメソッドは
'*          失敗する。
'*          adCopyOverWriteオプションを慎重に使用すること｡たとえば､ファイル
'*          をディレクトリにコピーするときにこのオプションを指定すると､ディレク
'*          トリが削除され､ファイルに置き換えられる｡
'*          [!メモ] http スキームを使用している URL は、Microsoft OLE DB
'*          Provider for Internet Publishing を自動的に呼び出す。
'*
'* @param Source      省略可能。コピー元のエンティティ (ファイル、ディレクトリ
'*                    など) を指定する URL。
'*                    Source を省略した場合、または空の文字列を指定した
'*                    場合、現在の Record で表されているファイルまたはディレク
'*                    トリがコピーされる。
'* @param Destination 省略可能。Source のコピー先の場所を指定する URL。
'* @param UserName    省略可能。Destination へのアクセス権が設定されている
'*                    場合、アクセス権を持つユーザーID。
'* @param Password    省略可能。UserName を確認するためのパスワード。
'* @param Options     省略可能。このメソッドの動作を指定。
'* @param Async       省略可能。ブール型 ( Boolean ) の値を指定。値が True の
'*                    場合、この動作は非同期で実行される。
'* @return 通常はDestination の値が返却。
'*         ただし、実際に返される値はプロバイダーによって異なる。
'******************************************************************************
Function CopyRecord(Optional Source As String, _
                    Optional Destination As String, _
                    Optional UserName As String, _
                    Optional Password As String, _
                    Optional Options As CopyRecordOptionsEnum = adCopyUnspecified, _
                    Optional Async As Boolean = False) As String
    CopyRecord = mOrigin.CopyRecord(Source, Destination, UserName, Password, Options, Async)
End Function
                    
'******************************************************************************
'* [概  要] DeleteRecordメソッド
'* [詳  細] DeleteRecordのラッパーメソッド。
'*          Record で表されるエンティティを削除する。
'*          このメソッドの終了後、Record で表されているオブジェクトに対する操作
'*          が失敗する場合がある。
'*          DeleteRecord を呼び出した後は Record を閉じる必要がある。
'*          これは、プロバイダーがデータ ソースで Record をいつ更新するかにより、
'*          Record の動作が予期できなくなることがあるためである。
'*          このレコードがレコードセットから取得された場合、この操作の結果はすぐ
'*          にレコードセットに反映されない。
'*          recordset を最新の情報に更新するには、レコードセットをいったん閉じて
'*          から開き直すか、またはrecordsetの再クエリを実行するか、 Updateメソッ
'*          ドとResyncメソッドを実行する。
'*          [!メモ] http スキームを使用している URL は、Microsoft OLE DB
'*          Provider for Internet Publishing を自動的に呼び出す。
'*
'* @param Source 省略可能。削除するエンティティ (ファイル、ディレクトリなど) を
'*               表す URL。Source を省略するか、または空の文字列を指定した場合、
'*               現在 Record で表されているエンティティが削除される。
'*               Record がコレクション レコード (ディレクトリなど、RecordType
'*               が adCollectionRecord であるもの) である場合は、すべての子
'*               (サブディレクトリなど) も削除される。
'* @param Async  省略可能。ブール型 ( Boolean ) の値を指定する。
'*               True のときは、削除操作が非同期で実行されることを示す。
'* @return 返される値はプロバイダーによって異なる。
'******************************************************************************
Function DeleteRecord(Optional Source As String, _
                      Optional Async As Boolean) As String

    DeleteRecord = mOrigin.DeleteRecord(Source, Async)
End Function


'******************************************************************************
'* [概  要] GetChildrenメソッド
'* [詳  細] GetChildrenのラッパーメソッド。
'*          各行がコレクション Record の子を表す Recordset を返却する。
'*          返される Recordset にどのような列があるかは、プロバイダーによって
'*          決まる。たとえば、ドキュメント ソース プロバイダーは、常にリソース
'*          の Recordset を返す。
'*
'* @return 各行が現在の Record オブジェクトの子を表す Recordset オブジェクト。
'*         たとえば、ディレクトリを表す Record の子は、親ディレクトリに含まれる
'*         ファイルとサブディレクトリとなる。
'******************************************************************************
Function GetChildren() As ADODBExRecordset
    Dim wrap As ADODBExRecordset: Set wrap = New ADODBExRecordset
    Set wrap.Origin = mOrigin.GetChildren
    Set GetChildren = wrap
End Function

'******************************************************************************
'* [概  要] MoveRecordメソッド
'* [詳  細] MoveRecordのラッパーメソッド。
'*          Record で表されるエンティティを別の場所に移動する。
'*          Source と Destination の値は異なっている必要がある。
'*          値が等しい場合、実行時エラーが発生する。
'*          サーバー名、パス名、およびリソース名のうち、少なくとも 1 つが異なっ
'*          ている必要がある。
'*          Internet Publishing Provider を使用して移動されるファイルでは、
'*          Options で特に指定のない限り、移動されるファイル内のすべてのハイパ
'*          ーテキストリンクが更新される。
'*          Destination に既存のオブジェクト (たとえば、ファイルまたはディレク
'*          トリ) を指定する場合、adMoveOverWrite が指定されていないと、この
'*          メソッドは失敗する。
'*          [!メモ] adMoveOverWrite オプションは十分に注意して使用すること。
'*          たとえば、ファイルをディレクトリに移動するときにこのオプションを指
'*          定していると、移動先のディレクトリが "削除" され、移動元のファイル
'*          に置き換えられる。
'*          この操作の終了後、 Record オブジェクトの一部の属性 ( ParentURL
'*          プロパティなど) は更新されなくなる。
'*          Record オブジェクトのプロパティを更新するには、 Record を閉じ、そ
'*          のファイルまたはディレクトリが移動された場所の URL でもう一度開く
'*          こと。
'*          Recordset から取得した Record の場合、ファイルまたはディレクトリの
'*          移動後の場所は、 Recordset にすぐには反映されない。
'*          反映するには、 Recordset をいったん閉じてもう一度開くこと。
'*          [!メモ] http スキームを使用している URL は、Microsoft OLE DB
'*          Provider for Internet Publishing を自動的に呼び出す。
'*
'* @param Source      省略可能。移動する Record を示す URL。
'*                    Source を省略するか、または空文字列を指定すると、この
'*                    Record で表されるオブジェクトが移動さる。
'*                    たとえば、Record がファイルを表している場合は、ファイル
'*                    の内容が Destination で指定した場所に移動される。
'* @param Destination 省略可能。 Source の移動先の場所を指定する URL。
'* @param UserName    省略可能。 Destination へのアクセス権が設定されている
'*                    場合、アクセス権を持つユーザー ID。
'* @param Password    省略可能。 UserName を確認するためのパスワード。
'* @param Options     省略可能。このメソッドの動作を指定。
'* @param Async       省略可能。 ブール値。Trueの場合、この操作は非同期である
'*                    ことを指定。
'* @return 通常はDestination の値が返却。
'*         ただし、実際に返される値はプロバイダーによって異なる。
'******************************************************************************
Function MoveRecord(Optional Source As String, _
                    Optional Destination As String, _
                    Optional UserName As String, _
                    Optional Password As String, _
                    Optional Options As MoveRecordOptionsEnum = adMoveUnspecified, _
                    Optional Async As Boolean = False) As String
    MoveRecord = mOrigin.MoveRecord(Source, Destination, UserName, Password, Options, Async)
End Function


'******************************************************************************
'* [概  要] OpenRメソッド
'* [詳  細] Openのラッパーメソッド。
'*          既存の Record オブジェクトを開くか、または Record で表される新しい
'*          アイテム (ファイル、ディレクトリなど) を作成する。
'*          Source には､次の値を指定できる｡
'*            ・URL。URL のプロトコルが http の場合、既定ではインターネット
'*              プロバイダーが呼び出される。URL が実行可能スクリプト (拡張子
'*              が .ASP のページなど) を含むノードを指している場合、既定では実
'*              行されたコンテンツではなく、そのソースを含む Record が開かれる。
'*              この動作は Options 引数を使用して修正する。
'*            ・Record オブジェクト。別の Record から開かれた Record オブジェク
'*              トは、元の Record オブジェクトを複製する。
'*            ・Command オブジェクト。開かれた Record オブジェクトは、 Command
'*              を実行することによって返された単一の行を表す。
'*              結果に複数の行が含まれる場合、レコードには最初の行の内容が入り、
'*              Errors コレクションにエラーが追加されることがある。
'*            ・SQL SELECT ステートメント。開かれた Record オブジェクトは、文字
'*              列の内容を実行することによって返された単一の行を表す。
'*              結果に複数の行が含まれる場合、レコードには最初の行の内容が入り、
'*              Errors コレクションにエラーが追加されることがある。
'*            ・テーブル名｡
'*          Record オブジェクトが、URL でアクセスできないエンティティ (データ
'*          ベースから派生した Recordset の行など) を表す場合、 ParentURL
'*          プロパティの値、および adRecordURL 定数でアクセスするフィールド
'*          の値はいずれも Null になる。
'*          [!メモ] http 体系を使用している URL は、Microsoft OLE DB
'*          Provider for Internet Publishing を自動的に呼び出す。
'*
'* @param Source           省略可能。Record オブジェクトで表されるエンティティ
'*                         の URL、 Command 、開かれた Recordset または別の
'*                         Record オブジェクト、SQL SELECT ステートメントを含む
'*                         文字列、またはテーブル名を表す値を指定。
'* @param ActiveConnection 省略可能。接続文字列または開かれた Connection オブ
'*                         ジェクト。
'* @param Mode             省略可能。結果の Record オブジェクトのアクセスモード
'*                         を ConnectModeEnum 値で指定。
'* @param CreateOptions    省略可能。既存のファイルまたはディレクトリを開くか、
'*                         新しいファイルまたはディレクトリを作成するかを
'*                         RecordCreateOptionsEnum 値で指定。
'*                         既定値の場合、アクセス モードは Mode プロパティから
'*                         取得される。
'*                         このパラメーターは、Source パラメーターに URL が含
'*                         まれていないと無視される。
'* @param Options          省略可能。Record を開くときのオプションを
'*                         RecordOpenOptionsEnum 値で指定。これらの値は組み合
'*                         わせることも可能。
'* @param UserName         省略可能。Source へのアクセス権が設定されている場合
'*                         アクセス権を持つユーザー ID 。
'* @param Password         省略可能。UserName を確認するためのパスワード。
'*
'******************************************************************************
Sub OpenR(Optional Source, _
          Optional ActiveConnection, _
          Optional Mode As ConnectModeEnum = adModeUnknown, _
          Optional CreateOptions As RecordCreateOptionsEnum = adFailIfNotExists, _
          Optional Options As RecordOpenOptionsEnum = adOpenRecordUnspecified, _
          Optional UserName As String, _
          Optional Password As String)
          
    If IsMissing(Source) Then
        If IsMissing(ActiveConnection) Then
            Call mOrigin.Open(, , Mode, CreateOptions, Options, UserName, Password)
        Else
            Call mOrigin.Open(, ActiveConnection, Mode, CreateOptions, Options, UserName, Password)
        End If
    Else
        If IsMissing(ActiveConnection) Then
            Call mOrigin.Open(Source, , Mode, CreateOptions, Options, UserName, Password)
        Else
            Call mOrigin.Open(Source, ActiveConnection, Mode, CreateOptions, Options, UserName, Password)
        End If
    End If
End Sub
