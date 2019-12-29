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
'* [機能名] ADODB.Connectionクラスのラップ・拡張クラス
'* [詳  細] ADODB.ConnectionクラスのWrapperとしての機能を提供する他、
'*          ADODB.Connectionクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Connectionクラスについて。
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
' Connectionオブジェクトは、データソースとの一意のセッションを表す。
' クライアント/サーバーデータベースシステムでは、サーバーへの実際の
' ネットワーク接続と同じ接続を表す場合がある｡
' プロバイダーのサポート機能によっては、コレクションオブジェクトのいくつかの
' コレクション、メソッド、プロパティを使用できない可能性がある｡
'
' <イベント情報>
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
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Connection
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
    If Not mOrigin Is Nothing Then
        If mOrigin.State <> adStateClosed Then
            Call mOrigin.Close
        End If
    End If
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* Attributes プロパティ
'*
'* [補  足]
'* ・オブジェクトの 1 つまたは複数の属性を示す。
'* ・Attributes プロパティは値の取得および設定が可能で、その値は1つまたは複数の
'*   XactAttributeEnum 値の合計になる。既定値は 0 。
'* ・複数の属性を設定する場合は、該当する定数の合計を使用できる。
'*   プロパティの値を､互換性のない定数を含む合計に設定すると､エラーが発生する｡
'* ・このプロパティは､クライアント側のConnectionオブジェクトでは使用できない｡
'*-----------------------------------------------------------------------------
Property Get Attributes() As Long
    Attributes = mOrigin.Attributes
End Property
Property Let Attributes(arg As Long)
    mOrigin.Attributes = arg
End Property

'*-----------------------------------------------------------------------------
'* CommandTimeout プロパティ
'*
'* [補  足]
'* ・コマンドが実行されるまでの待機時間を秒単位で示す値を設定または取得する。
'*   既定値は 30 。
'* ・ネットワーク トラフィックやサーバーの過負荷により実行が遅れている Execute
'*   メソッドの呼び出しを取り消すことができるようにするには、 Connection オブジ
'*   ェクトまたは Command オブジェクトの CommandTimeout プロパティを使用する。
'*   コマンドの実行が完了する前に CommandTimeout プロパティで設定された時間が
'*   経過すると、エラーが発生してコマンドが取り消される。
'*   プロパティを 0 に設定すると、コマンド実行が完了するまで無限に待機する。
'*   コードを書き込むプロバイダーとデータ ソースが CommandTimeout 機能をサポー
'*   トしていることを確認すること。
'* ・Connection オブジェクトの CommandTimeout 設定は、同じ Connection 上の
'*   Command オブジェクトの CommandTimeout 設定に影響しない。
'*   つまり、 Command オブジェクトの CommandTimeout プロパティは、 Connection
'*   オブジェクトの CommandTimeout の値を継承しない。
'* ・Connection オブジェクトでは、 CommandTimeout プロパティは Connection が開
'*   かれた後も、読み取り/書き込みが可能。
'*-----------------------------------------------------------------------------
Property Get CommandTimeout() As Long
    CommandTimeout = mOrigin.CommandTimeout
End Property
Property Let CommandTimeout(arg As Long)
    mOrigin.CommandTimeout = arg
End Property

'*-----------------------------------------------------------------------------
'* ConnectionString プロパティ（既定メンバー）
'*
'* [補  足]
'* ・データソースとの接続を確立するために使用される情報を示す。
'* ・一連の引数 = valueステートメントを含む詳細な接続文字列をセミコロンで区切っ
'*   て渡して、データソースを指定する。
'* ・ADO では、 ConnectionString プロパティに対して 5 種類の引数がサポートされ
'*   る。その他の引数は ADO で処理されずに直接プロバイダーに渡される。
'*   ADO でサポートされる引数を次に示す。
'*
'*   Provider =       接続に使用するプロバイダーの名前を指定する。
'*   File Name=       設定済みの接続情報を格納したプロバイダー固有のファイル
'*                    (たとえば、持続的なデータ ソース オブジェクト) の名前を
'*                    指定する。
'*   Remote Provider= クライアント側の接続を開くときに使用するプロバイダーの
'*                    名前を指定する。 (リモートデータサービスのみ)
'*   Remote Server=   クライアント側の接続を開くときに使用するサーバーのパス名
'*                    を指定する (リモート データ サービスのみ)。
'*   URL =            接続文字列を、ファイルやディレクトリなどのリソースを識別
'*                    する絶対 URL として指定する。
'*
'* ・ConnectionString プロパティを設定して Connection オブジェクトを開いた後、
'*   ADO によって定義された引数名がプロバイダーの対応する引数名にマップされる
'*   など、プロバイダーによってプロパティの内容が変更される場合がある。
'* ・ConnectionString プロパティは、Open メソッドの ConnectionString 引数に
'*   使用された値を自動的に継承する。
'*   したがって、Open メソッドの呼び出し中に現在の ConnectionString プロパティ
'*   を上書きできる。
'* ・File Name 引数により関連のあるプロバイダーが呼び出されるため、Provider
'*   引数と File Name 引数の両方を渡すことはできない。
'* ・ConnectionString プロパティは、接続が閉じているときは読み取り/書き込み
'*   可能で、開いているときは読み取り専用になる。
'* ・ConnectionString プロパティにおいて重複している引数は無視される。
'*   引数の最後のインスタンスが使用される。
'* ・リモートデータサービスの使用状況クライアント側のConnectionオブジェクトで
'*   使用する場合、 ConnectionStringプロパティに含めることができるのは、リモー
'*   トプロバイダーおよびリモートサーバーのパラメーターだけである。
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
'* ConnectionTimeout プロパティ
'*
'* [補  足]
'* ・接続が開くまでの待機時間を秒単位で示す。既定値は 15 。
'* ・ネットワーク トラフィックやサーバーの過度の使用が原因で接続を開く試みを
'*   中止する必要がある場合は、Connection オブジェクトで ConnectionTimeout
'*   プロパティを使用する。
'*   接続が開かれる前に ConnectionTimeout プロパティで設定した時間が経過した
'*   場合は、エラーが発生して接続の試みが取り消される。
'*   このプロパティを 0 に設定した場合は、ADO は接続が開かれるまで無限に待機
'*   する。コードを書き込むプロバイダーが、 ConnectionTimeout 機能をサポート
'*   していることを確認すること。
'* ・ConnectionTimeout プロパティは､接続が閉じているときは読み取り / 書き込み
'*   可能で､開いているときは読み取り専用になる｡
'*-----------------------------------------------------------------------------
Property Get ConnectionTimeout() As Long
    ConnectionTimeout = mOrigin.ConnectionTimeout
End Property
Property Let ConnectionTimeout(arg As Long)
    mOrigin.ConnectionTimeout = arg
End Property

'*-----------------------------------------------------------------------------
'* CursorLocation プロパティ
'*
'* [補  足]
'* ・カーソルサービスの場所を示す。
'* ・このプロパティで、プロバイダーにアクセス可能なさまざまなカーソル ライブラ
'*   リの中からカーソル サービスを選択する。
'*   通常は、クライアント側カーソル ライブラリ、またはサーバー側カーソル ライブ
'*   ラリから選択する。
'* ・このプロパティ設定は、プロパティが設定された後に確立された接続のみに作用
'*   する。CursorLocation プロパティを変更しても既存の接続には影響しない。
'* ・Execute メソッドが返すカーソルは、この設定を継承する。
'*   Recordset オブジェクトは、関連付けられた接続からこの設定を自動的に継承する。
'* ・このプロパティは、Connection または閉じている Recordset 上では読み取り/書き
'*   込み可能だが、開いている Recordset 上では読み取り専用になる。
'* ・リモートデータサービスの使用状況クライアント側のRecordsetオブジェクトまた
'*   はConnectionオブジェクトで使用する場合､カーソル位置プロパティはadUseClient
'*   にのみ設定できる｡
'*-----------------------------------------------------------------------------
Property Get CursorLocation() As CursorLocationEnum
    CursorLocation = mOrigin.CursorLocation
End Property
Property Let CursorLocation(arg As CursorLocationEnum)
    mOrigin.CursorLocation = arg
End Property

'*-----------------------------------------------------------------------------
'* DefaultDatabase プロパティ
'*
'* [補  足]
'* ・Connection オブジェクトの既定のデータベースを示す。
'* ・既定のデータベースがある場合、そのデータベースのオブジェクトにアクセスする
'*   SQL 文の構文が不適切な場合がある。
'*   DefaultDatabase プロパティで指定されたデータベース以外のデータベースのオブ
'*   ジェクトにアクセスするには、オブジェクト名を目的のデータベース名で修飾する
'*   必要がある。
'*   接続時に、プロバイダーは DefaultDatabase プロパティに既定のデータベース情
'*   報を書き込む。
'*   プロバイダーの中には 1 つの接続に 1 つのデータベースしか許可しないものがあ
'*   り、その場合は DefaultDatabase プロパティを変更できない。
'* ・データ ソースとプロバイダーによっては､この機能をサポートせず､エラーまたは
'*   空の文字列を返す場合がある｡
'* ・このプロパティは､クライアント側のConnectionオブジェクトでは使用できない｡
'*-----------------------------------------------------------------------------
Property Get DefaultDatabase() As String
    DefaultDatabase = mOrigin.DefaultDatabase
End Property
Property Let DefaultDatabase(arg As String)
    mOrigin.DefaultDatabase = arg
End Property

'*-----------------------------------------------------------------------------
'* Errors プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Connectionに設定されているErrorsオブジェクトを取得する。
'*-----------------------------------------------------------------------------
Property Get Errors() As ADODBExErrors
    Dim wrap As ADODBExErrors: Set wrap = New ADODBExErrors
    Set wrap.Origin = mOrigin.Errors
    Set Errors = wrap
End Property

'*-----------------------------------------------------------------------------
'* IsolationLevel プロパティ
'*
'* [補  足]
'* ・Connection オブジェクトの分離レベルを示す。
'*   既定値は adXactChaos 。
'* ・IsolationLevel プロパティでは、 Connection オブジェクトの分離レベルを設定
'*   する。
'*   設定値は、次に BeginTrans メソッドを呼び出すまで有効にならない。
'*   要求した分離レベルを使用できない場合、プロバイダーはその次に高い分離レベル
'*   を返す。
'* ・IsolationLevel プロパティは読み取り / 書き込み可能｡
'* ・リモートデータサービスの使用状況クライアント側の Connection オブジェクトで
'*   使用する場合、 IsolationLevelプロパティはadXactUnspecifiedにのみ設定できる。
'* ・ユーザーは、クライアント側のキャッシュ上の接続されていない Recordset オブ
'*   ジェクトで作業するため、マルチユーザーの場合は問題になることがある。
'*   たとえば、2 人のユーザーが同じレコードを更新しようとした際、リモート データ
'*   サービスは単純に先に操作を行ったユーザーの更新を受け付ける。
'*   2 番目のユーザーの更新要求はエラーになって失敗する。
'*-----------------------------------------------------------------------------
Property Get IsolationLevel() As IsolationLevelEnum
    IsolationLevel = mOrigin.IsolationLevel
End Property
Property Let IsolationLevel(arg As IsolationLevelEnum)
    mOrigin.IsolationLevel = arg
End Property

'*-----------------------------------------------------------------------------
'* Mode プロパティ
'*
'* [補  足]
'* ・使用可能なデータ変更権限を示す。
'*   Connection の既定値は adModeUnknown 。
'* ・Mode プロパティは、プロバイダーが現在の接続で使用しているアクセス権を設定
'*   または取得するために使用する。
'*   Mode プロパティは、 Connection オブジェクトが閉じているときにのみ設定できる。
'* ・このプロパティは、オブジェクトが閉じているときは読み取り/書き込み可能で、
'*   オブジェクトが開いているときは読み取り専用になる。
'* ・リモートデータサービスの使用状況クライアント側の Connection オブジェクトで
'*   使用する場合、 Modeプロパティはadmodeunknownにのみ設定できる。
'*-----------------------------------------------------------------------------
Property Get Mode() As ConnectModeEnum
    Mode = mOrigin.Mode
End Property
Property Let Mode(arg As ConnectModeEnum)
    mOrigin.Mode = arg
End Property

'*-----------------------------------------------------------------------------
'* Properties プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Connectionに設定されているPropertiesオブジェクトを取得する。
'*-----------------------------------------------------------------------------
Property Get Properties() As ADODBExProperties
    Dim wrap As ADODBExProperties: Set wrap = New ADODBExProperties
    Set wrap.Origin = mOrigin.Properties
    Set Properties = wrap
End Property

'*-----------------------------------------------------------------------------
'* Provider プロパティ
'*
'* [補  足]
'* ・Connection オブジェクトのプロバイダー名を示す。
'* ・接続用のプロバイダーの名前を設定または取得するには、Provider プロパティを
'*   使用する。
'*   このプロパティは、ConnectionString プロパティまたは Open メソッドの
'*   ConnectionString 引数の内容によって設定することもできる。
'*   ただし、Open メソッドを呼び出すときに複数の箇所でプロバイダーを指定すると
'*   予期しない結果が生じる可能性がある。
'*   プロバイダーを指定しない場合、このプロパティは既定値の MSDASQL (Microsoft
'*   OLE DB Provider for ODBC) になる。
'* ・Provider プロパティは、接続が閉じているときは値の設定および取得が可能で、
'*   接続が開いているときは値の取得のみが可能。
'*   設定値は Connection オブジェクトを開くか、Connection オブジェクトの
'*   Properties コレクションにアクセスするまで有効にならない。
'*   設定が無効である場合は、エラーが発生する。
'*-----------------------------------------------------------------------------
Property Get Provider() As String
    Provider = mOrigin.Provider
End Property
Property Let Provider(arg As String)
    mOrigin.Provider = arg
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
Property Get State() As Long
    State = mOrigin.State
End Property

'*-----------------------------------------------------------------------------
'* Version プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ADO のバージョン番号を示す。
'* ・ADO 実装のバージョン番号を返すには、Version プロパティを使用する。
'* ・プロバイダーのバージョンは､Properties コレクションの動的プロパティとして
'*   使用可能｡
'*-----------------------------------------------------------------------------
Property Get Version() As String
    Version = mOrigin.Version
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
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
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] BeginTransメソッド
'* [詳  細] BeginTransのラッパーメソッド。
'*          新しいトランザクションを開始する。
'*          いったん BeginTrans メソッドを呼び出すと、 CommitTrans メソッド
'*          またはRollbackTrans メソッドを呼び出してトランザクションを終了する
'*          まで、変更がコミットされることはない。
'*          入れ子になったトランザクションをサポートするプロバイダーの場合、
'*          開いているトランザクションで BeginTrans メソッドを呼び出すと、
'*          入れ子になった新しいトランザクションが開始される。
'*          戻り値は、入れ子のレベルを示す。
'*          戻り値 "1" は最上位レベルのトランザクション (他のトランザクション
'*          の入れ子になっていないトランザクション) が開いたことを示し、
'*          "2" は第 2 レベルのトランザクション (最上位レベルのトランザクショ
'*          ンの入れ子になっているトランザクション) が開いたことを示す ("3"
'*          以下も同様)。
'*          CommitTrans メソッドまたは RollbackTrans メソッドを呼び出すと、
'*          最後に開いたトランザクションだけが処理される。
'*          さらに上のレベルのトランザクションを処理するには、現在のトランザ
'*          クションを閉じるか、またはロールバックする必要がある。
'*
'* @return トランザクションの入れ子レベルを示す値
'******************************************************************************
Function BeginTrans() As Long
    Call mOrigin.BeginTrans
End Function

'******************************************************************************
'* [概  要] Cancelメソッド
'* [詳  細] Cancelのラッパーメソッド。
'*          保留中の非同期メソッド呼び出しの実行を取り消す。
'*          非同期メソッド呼び出し (つまり、adAsyncConnect、adAsyncExecute、
'*          または adAsyncFetch オプションを指定して呼び出されたメソッド) の
'*          実行を中止する。
'*          Connectionの場合、ExecuteまたはOpenが最後の非同期呼び出しが中止され
'*          るメソッド｡
'*
'******************************************************************************
Sub Cancel()
    Call mOrigin.Cancel
End Sub

'******************************************************************************
'* [概  要] CloseCnメソッド
'* [詳  細] Closeのラッパーメソッド。
'*          開いているオブジェクトおよびそれに関連するすべてのオブジェクトを
'*          閉じる。
'*
'*          関連するすべてのシステム リソースを解放する場合に使用する。
'*          オブジェクトを閉じてもメモリからは削除されず、プロパティ設定を変更
'*          してもう一度開くことができる。
'*          オブジェクトをメモリから完全に削除するには、オブジェクトを閉じた後
'*          に、オブジェクト変数をNothingに設定する。
'*          Close メソッドを使用して Connection オブジェクトを閉じると、その
'*          接続に関連するアクティブな Recordset オブジェクトもすべて閉じる。
'*          閉じた Connection オブジェクトに関連する Command オブジェクトは
'*          そのまま維持されるが、 Connection オブジェクトとの関連はなくなる。
'*          つまり、 ActiveConnection プロパティは Nothing に設定される。
'*          また、 Command オブジェクトの Parameters コレクション内のプロバイ
'*          ダー定義のパラメーターはすべてクリアされる。
'*          後で Open メソッドを呼び出して、同じデータ ソースまたは別のデータ
'*          ソースへの接続を再度確立することができる。
'*          Connection オブジェクトが閉じている間に、データ ソースに対する開い
'*          た接続を必要とするメソッドを呼び出すと、エラーが発生する。
'*          接続上に開いた Recordset オブジェクトがある状態で Connection オブ
'*          ジェクトを閉じると、すべての Recordset オブジェクトの保留中の変更
'*          がすべてロールバックされる。
'*          トランザクションの処理中に Close メソッドを呼び出して明示的に
'*          Connection オブジェクトを閉じると、エラーが発生する。
'*          トランザクションの処理中に Connection オブジェクトが適用範囲を外れ
'*          ると、トランザクションは自動的にロールバックされる。
'*
'******************************************************************************
Sub CloseCn()
    Call mOrigin.Close
End Sub

'******************************************************************************
'* [概  要] CommitTransメソッド
'* [詳  細] CommitTransのラッパーメソッド。
'*          すべての変更を保存して現在のトランザクションを終了する｡
'*          新しいトランザクションを開始する場合もある｡
'*
'*          Connection オブジェクトの Attributes プロパティによっては、
'*          CommitTrans メソッドまたは RollbackTrans メソッドを呼び出すと、
'*          新しいトランザクションが自動的に開始する場合がある。
'*          Attributes プロパティが adXactCommitRetaining に設定されている場合
'*          は、 CommitTrans メソッドを呼び出すと、プロバイダーが新しいトラン
'*          ザクションを自動的に開始する。
'*
'******************************************************************************
Sub CommitTrans()
    Call mOrigin.CommitTrans
End Sub

'******************************************************************************
'* [概  要] Executeメソッド
'* [詳  細] Executeのラッパーメソッド。
'*          指定されたクエリ、SQL ステートメント、ストアド プロシージャ、または
'*          プロバイダー固有のテキストを実行する。
'*
'*          Connection オブジェクトで Execute メソッドを使用すると、CommandText
'*          引数でメソッドに渡したクエリが、その接続で実行される。
'*          CommandText 引数で行を返すクエリを指定すると、実行結果はすべて新しい
'*          Recordset オブジェクトに格納される。
'*          コマンドが結果を返すことを意図していない場合 (SQL UPDATE クエリな
'*          ど)、プロバイダーはオプションadExecuteNoRecordsが指定されている限
'*          り、 Nothingを返す。
'*          それ以外の場合、Execute は閉じられたRecordsetを返す。
'*          返される Recordset オブジェクトは常に、読み取り専用で前方スクロール
'*          タイプのカーソルである。
'*          より多くの機能を備えた Recordset オブジェクトが必要な場合は、まず
'*          目的のプロパティ設定を持つ Recordset オブジェクトを作成し、次に
'*          Recordset オブジェクトの Open メソッドを使用してクエリを実行して、
'*          目的のカーソルの種類を返す。
'*          CommandText引数の内容はプロバイダーに固有のものであり、標準の SQL
'*          構文またはプロバイダーがサポートする特別なコマンド形式を使用できる。
'*          この操作が完了すると､ExecuteComplete イベントが発生する｡
'*
'* @param CommandText     実行する SQL ステートメント、ストアド プロシージャ、
'*                        URL、またはプロバイダー固有のテキスト。
'*                        テーブル名を使用することもできるが、プロバイダーが
'*                        SQL 対応である場合に限られる。
'*                        たとえば、"Customers" というテーブル名を使用すると、
'*                        SQL の標準の Select 構文が自動的に前に付加されて
'*                        "SELECT * FROM Customers" というステートメントが作成
'*                        され、T-SQLステートメントとしてプロバイダーに渡される。
'* @param RecordsAffected 省略可能。この操作の影響を受けたレコード数。
'* @param Options         省略可能。MoveRecordOptionsEnum 値を指定。
'*                        プロバイダーが CommandText 引数を評価する方法を示す。
'*                        1つまたは複数の CommandTypeEnum 値または
'*                        ExecuteOptionEnum 値のビットマスクを指定できる。
'*                        [!メモ] 内部処理を最小限に抑えてパフォーマンスを向上
'*                        させるには、 ExecuteOptionEnum の値 adExecuteNoRecords
'*                        を使用する。
'*                        Execute では、 CommandTypeEnum の値 adCmdFile または
'*                        adCmdTableDirect を使用しないこと。
'*                        これらの値は、Recordset の Open メソッドおよび Requery
'*                        メソッドのオプションとしてのみ使用できる。
'*
'* @return Recordset オブジェクトの参照
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
'* [概  要] OpenCnメソッド
'* [詳  細] Openのラッパーメソッド。
'*          データソースへの接続を開く｡
'*
'*          Connection オブジェクトで Open メソッドを使用すると、データソース
'*          への物理的な接続が確立される。
'*          このメソッドが正常に終了すると、接続が有効になり、接続に対してコマ
'*          ンドを発行して、その結果を処理することができる。
'*          省略可能なConnectionString引数を使用して、セミコロンで区切られた
'*          一連の引数 = valueステートメントを含む接続文字列を指定するか、
'*          または URL で識別されるファイルまたはディレクトリリソースを指定する。
'*          connectionstringプロパティは、 connectionstring引数に使用された値を
'*          自動的に継承する。
'*          そのため、 connectionオブジェクトを開く前に、 connectionstring
'*          プロパティを設定するか、またはconnectionstring引数を使用して、
'*          Openメソッドの呼び出し中に現在の接続パラメーターを設定または上書き
'*          することができる。
'*          ユーザーとパスワードの情報を、ConnectionString 引数と、省略可能な
'*          UserID 引数および Password 引数の両方で渡すと、UserID 引数と
'*          Password 引数の方が、ConnectionString で指定した値より優先される。
'*          開いている Connection オブジェクトに対する操作が完了したら、Close
'*          メソッドを使用して関連するすべてのシステム リソースを解放する。
'*          オブジェクトを閉じても、オブジェクトはメモリから削除されない。
'*          オブジェクトのプロパティ設定を変更し、Open メソッドを使用して、
'*          後で再度開くことができる。
'*          オブジェクトをメモリから完全に削除するには、オブジェクト変数に
'*          Nothing を設定する。
'*          リモートデータサービスの使用状況クライアント側のconnectionオブジェ
'*          クトで使用する場合、 connectionオブジェクトでRecordsetが開かれる
'*          まで、 Openメソッドは実際にはサーバーへの接続を確立しない。
'*
'* @param ConnectionString 省略可能。接続情報 の値。
'*                         有効な設定値の詳細については、 ConnectionString
'*                         プロパティを参照。
'* @param UserID           省略可能。接続を確立するときに使用するユーザー名。
'* @param Password         省略可能。接続を確立するときに使用するパスワード。
'* @param Options          省略可能。接続が確立された後 (同期) または接続が
'*                         確立される前 (非同期) のどちらでこのメソッドが返る
'*                         かを指定する ConnectOptionEnum 値。
'******************************************************************************
Sub OpenCn(Optional ConnectionString As String, _
           Optional UserID As String, _
           Optional Password As String, _
           Optional Options As Long = -1)
    Call mOrigin.Open(ConnectionString, UserID, Password, Options)
End Sub
         
'******************************************************************************
'* [概  要] OpenSchemaメソッド
'* [詳  細] OpenSchemaのラッパーメソッド。
'*          プロバイダーからデータベースのスキーマ情報を取得する｡
'*
'*          OpenSchema メソッドは、データソースに含まれるテーブル、テーブルに
'*          含まれる列、サポートされているデータ型などのデータソースに関する
'*          情報を返す。
'*          QueryType 引数は、返される列 (スキーマ) を示す GUID である。OLE DB
'*          の仕様には、すべてのスキーマの一覧がある。
'*          引数Restrictionsは、スキーマクエリの結果を制限する。
'*          Restrictionsには、結果のRecordsetで、対応する列の列のサブセット
'*           (制約列)に必要な値の配列を指定する。
'*          OLE DB 仕様以外の非標準スキーマ クエリをプロバイダーが独自に定義して
'*          いる場合は、QueryType 引数に adSchemaProviderSpecific を使用する。
'*          この定数を使用する場合は、SchemaID 引数に、実行するスキーマ クエリ
'*          の GUID を指定する必要がある。
'*          QueryType が adSchemaProviderSpecific に設定され、SchemaID が指定
'*          されていない場合、エラーが発生しない。
'*          プロバイダーは、すべての OLE DB 標準スキーマ クエリをサポートする
'*          必要はない。
'*          OLE DB の仕様では、adSchemaTables、adSchemaColumns、および
'*          adSchemaProviderTypes のみが必要とされる。
'*          ただし、これらのスキーマ クエリでは、プロバイダーは Criteria の
'*          制約をサポートする必要はない。
'*          リモートデータサービスの使用状況 OpenSchemaメソッドは、クライアン
'*          ト側のConnectionオブジェクトでは使用できない。
'*
'* @param Schema       実行するスキーマクエリの種類を表す SchemaEnum 値を指定
'* @param Restrictions 省略可能。 SchemaEnum の指定に従って、各QueryTypeオプ
'*                     ションのクエリ制約の配列を指定
'* @param SchemaID     OLE DB の仕様で定義されていない、プロバイダースキーマ
'*                     のクエリの GUID (グローバル一意識別子) を指定。
'*                     このパラメーターは、QueryType が adSchemaProviderSpecific
'*                     に設定されている場合は必須。
'*                     それ以外の場合、このパラメーターは使用しない。
'* @return スキーマ情報を含む Recordset オブジェクト。 Recordset は読み取り
'*         専用の静的カーソルとして開かれる。QueryType により、Recordset に
'*         表示される列が決定される。
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
'* [概  要] RollbackTransメソッド
'* [詳  細] RollbackTransのラッパーメソッド。
'*          現在のトランザクションの間に行われたすべての変更をキャンセルし、
'*          トランザクションを終了する。
'*          新しいトランザクションを開始する場合もある。
'*
'*          Connection オブジェクトの Attributes プロパティによっては、
'*          CommitTrans メソッドまたは RollbackTrans メソッドを呼び出すと、
'*          新しいトランザクションが自動的に開始する場合がある。
'*          Attributes プロパティが adXactAbortRetaining に設定されている場合
'*          は、 RollbackTrans メソドを呼び出すと、プロバイダーが新しいトラン
'*          ザクションを自動的に開始する。
'*
'******************************************************************************
Sub RollbackTrans()
    Call mOrigin.RollbackTrans
End Sub


'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------
'******************************************************************************
'* [概  要] OpenPostgreSql メソッド
'* [詳  細] PostgreSqlデータベースへの接続を開く｡
'*          PostgreSql接続文字列を作成し、Connection.Open メソッドを実行する。
'*
'* @param Server   接続先サーバ名
'* @param Port     接続先ポート番号
'* @param Database 接続先データベース名
'* @param UserID   省略可能。接続を確立するときに使用するユーザー名。
'* @param Password 省略可能。接続を確立するときに使用するパスワード。
'* @param Options  省略可能。接続が確立された後 (同期) または接続が
'*                 確立される前 (非同期) のどちらでこのメソッドが返る
'*                 かを指定する ConnectOptionEnum 値。
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
