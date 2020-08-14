VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODB.Commandクラスのラップ・拡張クラス
'* [詳  細] ADODB.CommandクラスのWrapperとしての機能を提供する他、
'*          ADODB.Commandクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Commandクラスについて。
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Used to define a specific command, such as a SQL query, intended to run against
' a data source.
'
' データソースに対して実行することを目的とした、SQLクエリなどの特定のコマンドを
' 定義するために使用されます。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Command
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
'* ・開いている Connection オブジェクトまたは有効な接続文字列をこのプロパティ
'*   に設定する前に、 Command オブジェクトに対して Execute を呼び出そうとする
'*   と、エラーが発生する。
'* ・ActiveConnectionプロパティをNothingに設定すると現在の接続からCommandオブ
'*   ジェクトを切り離し、プロバイダーがデータに関連付けられているリソースを
'*   解放する。その後は、その Command オブジェクトを同じまたは別の Connection
'*   オブジェクトに関連付けることができる。
'*   プロバイダーによっては、プロパティをNothingに設定しなくても、ある接続か
'*   ら別の接続にプロパティ設定を変更することができる。
'* ・Command オブジェクトの Parameters コレクションにプロバイダーから供給され
'*   たパラメーターが格納されている場合は、ActiveConnection プロパティを
'*   Nothing に設定したり、他の Connection オブジェクトに設定したりすると、
'*   コレクションがクリアされる。
'*   手作業で Parameter オブジェクトを作成し、それを使用して Command オブジェ
'*   クトの Parameters コレクションにデータを格納した場合は、ActiveConnection
'*   プロパティを Nothing や他の Connection オブジェクトに設定しても、Parameters
'*   コレクションはそのまま残される。
'* ・Command オブジェクトが関連付けられている Connection オブジェクトを閉じる
'*   と、ActiveConnection プロパティが Nothing に設定さる。
'*   このプロパティに閉じている Connection オブジェクトを設定すると、
'*   エラーが発生する。
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
'* CommandStream プロパティ
'*
'* [補  足]
'* ・Commandオブジェクトの入力として使用されるストリームを示す。
'* ・Commandオブジェクトの入力として使用されるストリームを設定または返す。
'*   このストリームの形式はプロバイダー固有である。
'*   詳細については、プロバイダーのドキュメントを参照。
'*   このプロパティは、CommandTextプロパティに似ている。
'*   CommandTextプロパティは、コマンドの入力に文字列を指定するために使用される。
'* ・CommandStreamとCommandTextは相互に排他的である。
'*   ユーザーがCommandStreamプロパティを設定すると、CommandTextプロパティは空の
'*   文字列（ ""）に設定される。
'*   ユーザーがCommandTextプロパティを設定すると、CommandStreamプロパティは
'*   Nothingに設定される。
'* ・Command.Parameters.RefreshおよびCommand.Prepareメソッドの動作は､プロバイダ
'*   ーによって定義される｡
'*   ストリーム内のパラメーターの値は更新できない｡
'* ・入力ストリームは、コマンドのソースを返す他のADOオブジェクトでは使用でき
'*   ない。 たとえば、RecordsetのSourceが入力としてストリームを持つCommand
'*   オブジェクトに設定されている場合、Recordset.SourceはCommandTextプロパティ
'*   を返す。
'*   このプロパティには、ストリームのコンテンツの代わりに空の文字列（ ""）が含
'*   まれる。
'* ・（CommandStreamで指定された）コマンドストリームを使用する場合、CommandType
'*   プロパティの有効なCommandTypeEnum値はadCmdTextとadCmdUnknownのみである。
'*   その他の値はエラーの原因になる。
'*-----------------------------------------------------------------------------
Property Get CommandStream() As Variant
    If IsObject(mOrigin.CommandStream) Then
        Set CommandStream = mOrigin.CommandStream
    Else
        CommandStream = mOrigin.CommandStream
    End If
End Property
Property Set CommandStream(arg As Variant)
    If IsObject(arg) Then
        Set mOrigin.CommandStream = arg
    Else
        mOrigin.CommandStream = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* CommandText プロパティ
'*
'* [補  足]
'* ・SQL ステートメント、テーブル名、相対 URL、またはストアド プロシージャの
'*   呼び出しなど、プロバイダーのコマンドを含む文字列型 ( String ) の値を設定
'*   または取得する。既定値は "" (長さ 0 の文字列)。
'* ・Command オブジェクトで表されるコマンド テキストを設定または取得するには、
'*   CommandText プロパティを使用する。
'*   通常は SQL ステートメントを使うが、ストアド プロシージャの呼び出しなど、
'*   プロバイダーが認識する、他の種類のコマンド ステートメントでもかまわない。
'*   SQL ステートメントは、特定の文法またはプロバイダーのクエリ プロセッサが
'*   サポートするバージョンである必要がある。
'* ・Command オブジェクトの Prepared プロパティが True に設定されていて、
'*   CommandText プロパティを設定するときに開いていた接続に Command オブジェクト
'*   がバインドされている場合、 Execute メソッドまたは Open メソッドを呼び出す
'*   と、クエリ (つまり、プロバイダーが保存するコンパイルされたクエリ) が準備
'*   される。
'* ・CommandType プロパティの設定値によっては、 CommandText プロパティが変更
'*   される場合がある。
'*   CommandText プロパティはいつでも読み出すことができ、ADO がコマンド実行中
'*   に使う実際のコマンド テキストの参照も可能。
'* ・ファイルやディレクトリなどのリソースを指定する相対 URL を設定したり取得
'*   したりするには、 CommandText プロパティを使う。
'*   リソースは、絶対 URL で明示的に指定された位置や、開かれた Connection
'*   オブジェクトで暗黙的に指定された位置に対して相対的である。
'*-----------------------------------------------------------------------------
Property Get CommandText() As String
    CommandText = mOrigin.CommandText
End Property
Property Let CommandText(arg As String)
    mOrigin.CommandText = arg
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
'* CommandType プロパティ
'*
'* [補  足]
'* ・Command オブジェクトの型を示す。1 つまたは複数の CommandTypeEnum 値を設定
'*   または取得する。
'* ・[!メモ] CommandType では、 adCmdFile または adCmdTableDirect の
'*   CommandTypeEnum 値を使用しないこと。
'*   これらの値は、 Recordset の Open メソッドと Requery メソッドのオプション
'*   としてのみ使用することができる。
'* ・CommandType プロパティは、 CommandText プロパティの評価を最適化するために
'*   使用する。
'* ・CommandType プロパティの値が adCmdUnknown (既定値) と等しい場合、パフォー
'*   マンスが低下することがある。
'*   これは、 CommandText プロパティの型が SQL ステートメント、ストアドプロシー
'*   ジャ、またはテーブル名であるかを調べるためにプロバイダーを呼び出す必要があ
'*   るためである。
'*   使っているコマンドの種類がわかっている場合は、 CommandType プロパティを
'*   設定することにより、該当するコードに直接移動できる。
'*   CommandType プロパティが CommandText プロパティのコマンドの種類と一致しな
'*   い場合に Execute メソッドを呼び出すと、エラーが発生する。
'*-----------------------------------------------------------------------------
Property Get CommandType() As CommandTypeEnum
    CommandType = mOrigin.CommandType
End Property
Property Let CommandType(arg As CommandTypeEnum)
    mOrigin.CommandType = arg
End Property

'*-----------------------------------------------------------------------------
'* Dialect プロパティ
'*
'* [補  足]
'* ・CommandTextまたはCommandStreamプロパティの方言を示す。
'*   方言は、文字列またはストリームを解析するためにプロバイダーが使用する構文
'*   と一般的な規則を定義する。
'* ・Dialectプロパティには、コマンドテキストまたはストリームの方言を表す有効な
'*   GUIDが含まれている。
'*   このプロパティのデフォルト値は{C8B521FB-5CF3-11CE-ADE5-00AA0044773D}。
'*   これは、プロバイダーがコマンドテキストまたはストリームの解釈方法を選択する
'*   必要があることを示す。
'* ・ユーザーがこのプロパティの値を読み取っても､ADOはプロバイダーにクエリを実行
'*   しない｡ 現在Commandオブジェクトに格納されている値の文字列表現を返す｡
'* ・ユーザーがDialectプロパティを設定すると､ADOはGUIDを検証し､指定された値が
'*   有効なGUIDでない場合はエラーを発生させる｡
'*   プロバイダのドキュメントを参照して､Dialectプロパティでサポートされている
'*   GUID値を確認すること｡
'*-----------------------------------------------------------------------------
Property Get Dialect() As String
    Dialect = mOrigin.Dialect
End Property
Property Let Dialect(arg As String)
    mOrigin.Dialect = arg
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ
'*
'* [補  足]
'* ・オブジェクトの名前を示す文字列型 (String) の値を設定または取得する。
'* ・値は、 Command オブジェクトでは読み取り/書き込み可能。
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property
Property Let Name(arg As String)
    mOrigin.Name = arg
End Property

'*-----------------------------------------------------------------------------
'* NamedParameters プロパティ
'*
'* [補  足]
'* ・パラメーター名をプロバイダーに渡すかどうかを示す｡
'* ・このプロパティがtrueの場合、ADOはCommandオブジェクトのParameterコレクショ
'*   ンの各パラメーターのNameプロパティの値を渡す。
'*   プロバイダーは、パラメーター名を使用して、CommandTextまたはCommandStream
'*   プロパティのパラメーターと一致する。
'*   このプロパティがfalse（デフォルト）の場合、パラメーター名は無視され、プロ
'*   バイダーはパラメーターの順序を使用して、CommandTextまたはCommandStreamプ
'*   ロパティのパラメーターに値を一致させる。
'*-----------------------------------------------------------------------------
Property Get NamedParameters() As Boolean
    NamedParameters = mOrigin.NamedParameters
End Property
Property Let NamedParameters(arg As Boolean)
    mOrigin.NamedParameters = arg
End Property

'*-----------------------------------------------------------------------------
'* Parameters プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・Commandに設定されているParametersオブジェクトを取得する。
'*-----------------------------------------------------------------------------
Property Get Parameters() As ADODBExParameters
Attribute Parameters.VB_UserMemId = 0
    Dim wrap As ADODBExParameters: Set wrap = New ADODBExParameters
    Set wrap.Origin = mOrigin.Parameters
    Set Parameters = wrap
End Property

'*-----------------------------------------------------------------------------
'* Prepared プロパティ
'*
'* [補  足]
'* ・コンパイルされたバージョンのコマンドを実行前に保存するかどうかを示す。
'* ・True に設定されている場合は、コマンドの準備が必要であることを示す。
'* ・Command オブジェクトを最初に実行する前に、 CommandText プロパティで指定さ
'*   れたクエリの準備済み (コンパイル済み) バージョンをプロバイダーで保存する
'*   には、 Prepared プロパティを使用する。
'*   これによって、コマンドの最初の実行は遅くなることがあるが、プロバイダーで
'*   コマンドをコンパイルした後はコンパイル済みのコマンドが使用されるので、
'*   パフォーマンスが向上する。
'* ・このプロパティが False の場合、プロバイダーはコンパイル済みバージョンを
'*   作成せずに、直接 Command オブジェクトを実行する。
'* ・プロバイダーがコマンドの準備をサポートしていない場合、このプロパティを
'*   True に設定すると、すぐにエラーが返されることがある。
'*   エラーを返さない場合、プロバイダーは単にコマンドの準備の要求を無視し、
'*   Prepared プロパティを False に設定する。
'*-----------------------------------------------------------------------------
Property Get Prepared() As Boolean
    Prepared = mOrigin.Prepared
End Property
Property Let Prepared(arg As Boolean)
    mOrigin.Prepared = arg
End Property

'*-----------------------------------------------------------------------------
'* Properties プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Commandに設定されているPropertiesオブジェクトを取得する。
'*-----------------------------------------------------------------------------
Property Get Properties() As ADODBExProperties
    Dim wrap As ADODBExProperties: Set wrap = New ADODBExProperties
    Set wrap.Origin = mOrigin.Properties
    Set Properties = wrap
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
    
'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("ADODB.Command")
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
'*          Commandの場合、Executeが最後の非同期呼び出しが中止されるメソッド。
'*
'******************************************************************************
Sub Cancel()
    Call mOrigin.Cancel
End Sub

'******************************************************************************
'* [概  要] CreateParameterメソッド
'* [詳  細] CreateParameterのラッパーメソッド。
'*          指定したプロパティを使用して新規Parameterオブジェクトを作成する。
'*          名前、種類、方向、サイズ、および値を指定して新規 Parameter オブジェ
'*          クトを作成する。引数に指定した値は、対応する Parameter プロパティに
'*          書き込まれる。
'*          このメソッドは、CommandオブジェクトのParametersコレクションに
'*          Parameterオブジェクトを自動的に追加しない。
'*          これにより、コレクションにParameterオブジェクトを追加したときに
'*          ADOが検証する値を持つ追加のプロパティを設定できる。
'*          Type 引数で可変長データ型を指定した場合、Parameters コレクションに
'*          オブジェクトを追加する前に、Size 引数を渡すか、または Parameter
'*          オブジェクトの Size プロパティを設定する必要がある。
'*          そのようにしなかった場合はエラーが発生する。
'*          Type 引数に数値型 (adNumeric または adDecimal) を指定する場合は、
'*          同時に NumericScale プロパティと Precision プロパティも設定する必要
'*          がある。
'*
'* @param Name      省略可能。 Parameter オブジェクト名。
'* @param Type_     省略可能。 Parameter オブジェクトのデータ型。
'* @param Direction 省略可能。 Parameter オブジェクトの方向。
'* @param Size      省略可能。パラメーター値の最大長。
'* @param Value     省略可能。 Parameter オブジェクトの値。
'* @return Parameterオブジェクト
'******************************************************************************
Function CreateParameter(Optional Name As String, _
                         Optional Type_ As DataTypeEnum = adEmpty, _
                         Optional Direction As ParameterDirectionEnum = adParamInput, _
                         Optional Size As Long, _
                         Optional Value) As ADODBExParameter
    Dim wrap As ADODBExParameter: Set wrap = New ADODBExParameter
    
    If IsMissing(Name) Then
        If IsMissing(Size) Then
            If IsMissing(Value) Then
                Set wrap.Origin = mOrigin.CreateParameter(, Type_, Direction)
            Else
                Set wrap.Origin = mOrigin.CreateParameter(, Type_, Direction, , Value)
            End If
        ElseIf IsMissing(Value) Then
            Set wrap.Origin = mOrigin.CreateParameter(, Type_, Direction, Size)
        Else
            Set wrap.Origin = mOrigin.CreateParameter(, Type_, Direction, Size, Value)
        End If
    Else
        If IsMissing(Size) Then
            If IsMissing(Value) Then
                Set wrap.Origin = mOrigin.CreateParameter(Name, Type_, Direction)
            Else
                Set wrap.Origin = mOrigin.CreateParameter(Name, Type_, Direction, , Value)
            End If
        ElseIf IsMissing(Value) Then
            Set wrap.Origin = mOrigin.CreateParameter(Name, Type_, Direction, Size)
        Else
            Set wrap.Origin = mOrigin.CreateParameter(Name, Type_, Direction, Size, Value)
        End If
    End If
    Set CreateParameter = wrap
End Function
                         
'******************************************************************************
'* [概  要] Executeメソッド
'* [詳  細] Executeのラッパーメソッド。
'*          CommandText プロパティで指定されたクエリ、SQL ステートメント、
'*          またはストアドプロシージャを実行する。
'*
'*          Command オブジェクトに対して Execute メソッドを使用すると、オブジェ
'*          クトの CommandText プロパティで指定されているクエリが実行される。
'*          行を返すクエリが CommandText プロパティで指定されていると、実行が
'*          生成するすべての結果は、新しい Recordset オブジェクトに格納される。
'*          コマンドが行を返さないクエリである場合、プロバイダーは閉じた
'*          Recordset オブジェクトを返す。
'*          アプリケーション言語によっては、 Recordset が必要ないときはこの
'*          戻り値を無視してもかまわない。
'*          クエリにパラメーターがある場合、Execute呼び出しで渡されるパラメー
'*          ター値でこれらをオーバーライドしない限り、Commandオブジェクトの
'*          パラメーターの現在の値が使用される。
'*          Executeメソッドを呼び出すときに一部のパラメーターの新しい値を省略
'*          することにより、パラメーターのサブセットをオーバーライドできる。
'*          パラメーターを指定する順序は、メソッドがパラメーターを渡す順序と
'*          同じである。たとえば、4つ (またはそれ以上) のパラメーターがあり、
'*          第1と第4のパラメーターに対してのみ新しい値を渡す場合は、
'*          parameters引数として渡す。
'*
'* @param RecordsAffected 省略可能。この操作の影響を受けたレコード数。
'*                        アクションクエリまたはストアドプロシージャに対して
'*                        のみ適用される。
'*                        結果を返すクエリまたはストアドプロシージャによって
'*                        返されるレコードの数を返さない。
'*                        この情報を取得するには、 RecordCountプロパティを使用
'*                        する。 Executeメソッドは、 adasyncexecuteで使用されて
'*                        いる場合は正しい情報を返さない。
'*                        コマンドが非同期的に実行された場合、影響を受けるレコ
'*                        ードの数がメソッドから戻るときにまだ認識されない可能
'*                        性があるためである。
'* @param Parameters      省略可能。 SQL ステートメントに渡すパラメーター値。
'*                        出力パラメーターをこの引数で渡しても､正しい値は返
'*                        されない。
'* @param Options         省略可能です。MoveRecordOptionsEnum 値を指定する。
'*                        Command オブジェクトの CommandText プロパティを
'*                        プロバイダーが評価する方法を示す の値。
'*                        CommandTypeEnum 値および ExecuteOptionEnum 値のビット
'*                        マスク値を使用できる。たとえば、CommandText プロパティ
'*                        の値をテキストとして評価するようADOに指示し、コマンド
'*                        テキストの実行時に生成されるすべてのレコードをコマン
'*                        ドが破棄して返さないようにする場合は、adCmdText と
'*                        adExecuteNoRecords を組み合わせて使用できる。
'* @return Recordset オブジェクトの参照または Nothing
'******************************************************************************
Function Execute(Optional RecordsAffected, _
                 Optional Parameters, _
                 Optional Options As Long = -1) As ADODBExRecordset
    Dim wrap As ADODBExRecordset: Set wrap = New ADODBExRecordset

    If IsMissing(RecordsAffected) Then
        If IsMissing(Parameters) Then
            Set wrap.Origin = mOrigin.Execute(, , Options)
        Else
            Set wrap.Origin = mOrigin.Execute(, Parameters, Options)
        End If
    ElseIf IsMissing(Parameters) Then
        Set wrap.Origin = mOrigin.Execute(RecordsAffected, , Options)
    Else
        Set wrap.Origin = mOrigin.Execute(RecordsAffected, Parameters, Options)
    End If
    Set Execute = wrap
End Function

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] SQL表示処理。
'* [詳  細] SQLの埋込パラメータを「?」から置き換える。
'* [参  考] <http://kagen88.blogspot.com/2012/12/adodbcommand.html>
'*
'* @return SQL文字列
'******************************************************************************
' adArray
' adBigInt
' adBinary
' adBoolean
' adBSTR
' adChapter
' adChar
' adCurrency
' adDate
' adDBDate
' adDBTime
' adDBTimeStamp
' adDecimal
' adDouble
' adEmpty
' adError
' adFileTime
' adGUID
' adIDispatch
' adInteger
' adIUnknown
' adLongVarBinary
' adLongVarChar
' adLongVarWChar
' adNumeric
' adPropVariant
' adSingle
' adSmallInt
' adTinyInt
' adUnsignedBigInt
' adUnsignedInt
' adUnsignedSmallInt
' adUnsignedTinyInt
' adUserDefined
' adVarBinary
' adVarChar
' adVariant
' adVarNumeric
' adVarWChar
' adWChar
Public Function PrintSql() As String
    Dim p As ADODBExParameter
    Dim v As Variant
    Dim i As Integer
    Dim s As String
    i = 1
    ReDim v(Me.Parameters.Count)
    For Each p In mOrigin.Parameters
        Select Case p.Type_
            'ここはニーズに応じてタイプを追加
            Case adInteger
                v(i) = p.Value
            Case Else
                v(i) = "'" & p.Value & "'"
        End Select
        i = i + 1
    Next
    s = mOrigin.CommandText
    For i = 1 To UBound(v)
        s = Replace(s, "?", v(i), 1, 1)
    Next
    PrintSql = s
End Function
