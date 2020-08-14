VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODB.Streamクラスのラップ・拡張クラス
'* [詳  細] ADODB.StreamクラスのWrapperとしての機能を提供する他、
'*          ADODB.Streamクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Streamクラスについて。
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Represents a stream of binary or text data.
' For example, an XML document can be loaded into a stream for command input
' or returned from certain providers as the results of a query.
' A Stream object can be used to manipulate fields or records containing these
' streams of data.
'
' バイナリデータまたはテキストデータのストリームを表す｡
' たとえば､XMLドキュメントをコマンド入力用のストリームにロードしたり､クエリの
' 結果として特定のプロバイダーから返すことができる｡
' Streamオブジェクトを使用して､これらのデータストリームを含むフィールドまたは
' レコードを操作できる｡
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Stream
Private mOrigin As Object
Private mBOM As Boolean

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
'* Charset プロパティ
'*
'* [補  足]
'* ・テキスト Stream の内容を変換して Stream オブジェクト内部バッファーに格納
'*   するための文字セットを示す。
'* ・Stream の内容の変換に使用される文字セットを指定する文字列型 ( String ) の
'*   値を設定または取得する。 既定値は "Unicode" 。
'*   指定可能な値は、インターフェイスを介してインターネット文字セット文字列と
'*   して渡すことのできる通常の文字列 ("iso-8859-1"、"Windows-1252" など) 。
'*   システムによって認識される文字セット文字列の一覧については、Windows レジス
'*   トリの_HKEY_クラス\の\ルート\MIME データベース文字セットのサブキーを参照。
'* ・テキスト Stream オブジェクトでは、 Charset プロパティで指定された文字セッ
'*   トでテキスト データが格納される。既定値は Unicode 。
'*   Charset プロパティは、 Stream に送るデータの変換、または Stream から受け取
'*   るデータの変換に使用される。
'*   たとえば、 Stream に ISO-8859-1 データが格納されており、そのデータが BSTR
'*   にコピーされる場合、 Stream オブジェクトはデータを Unicode に変換する。
'*   逆の変換も同じように行われる。
'* ・開いている Stream の場合、 Charset を設定するには、現在の Position が Stream
'*   の先頭 (0) になっている必要がある。
'* ・Charset は、テキスト Stream オブジェクト (Type が adTypeText) でのみ使用
'*   される。Type が adTypeBinary の場合、このプロパティは無視される。
'*-----------------------------------------------------------------------------
Property Get CharSet() As String
    CharSet = mOrigin.CharSet
End Property
Property Let CharSet(arg As String)
    mOrigin.CharSet = arg
End Property

'*-----------------------------------------------------------------------------
'* EOS プロパティ（読み取り専用）
'*
'* [補  足]
'* ・現在の位置がストリームの末尾にあるかどうかを示す。
'* ・現在の位置がストリームの末尾かどうかを示すブール型 ( Boolean ) の値を返す。
'*   ストリームにバイトがなくなると EOS は True を返し、現在の位置の後にもバイ
'*   トがあると False を返す。
'* ・ストリームの末尾の位置を設定するには、SetEOS メソッドを使用する。
'*   現在の位置を確認するには、 Position プロパティを使用する。
'*-----------------------------------------------------------------------------
Property Get EOS() As Boolean
    EOS = mOrigin.EOS
End Property

'*-----------------------------------------------------------------------------
'* LineSeparator プロパティ
'*
'* [補  足]
'* ・テキスト Stream オブジェクトの行区切り文字として使用するバイナリ文字を示す。
'* ・Stream で使う行区切り文字を示す LineSeparatorsEnum 値を設定または取得する。
'*   既定値は adCRLF 。
'* ・LineSeparator は、テキスト Stream の内容を読むときに行を解釈するために使う。
'*   行は、SkipLine メソッドで読み飛ばすこともできる。
'* ・LineSeparator は、テキスト Stream オブジェクト (Type が adTypeText) だけに
'*   使用する。 Type が adTypeBinary の場合、このプロパティは無視される。
'*-----------------------------------------------------------------------------
Property Get LineSeparator() As LineSeparatorEnum
    LineSeparator = mOrigin.LineSeparator
End Property
Property Let LineSeparator(arg As LineSeparatorEnum)
    mOrigin.LineSeparator = arg
End Property

'*-----------------------------------------------------------------------------
'* Mode プロパティ
'*
'* [補  足]
'* ・使用可能なデータ変更権限を示す。
'*   基になるソースに関連付けられている Stream (URL により、ソース、または
'*   Record の既定 Stream として開かれたもの) の既定値は、adModeRead 。
'*   基になるソースに関連付けられていない (メモリ内でインスタンス化された)
'*   Streamの既定値は adModeUnknown 。
'* ・アクセス モードを指定しない場合、 Stream オブジェクトを開くときに使う
'*   ソースのモードが継承される｡
'*   たとえば、 Record オブジェクトから Stream を開くと、既定では Record と同じ
'*   モードで開く。
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
'* Position プロパティ
'*
'* [補  足]
'* ・Stream オブジェクト内の現在の位置を示す。
'* ・ストリームの先頭から現在位置までのオフセットをバイト単位で示す長整数型
'*   ( Long ) の値を設定または取得する。
'*   既定値は 0 で、ストリームの最初のバイトを表す。
'* ・現在位置はストリームの末尾よりも後ろに移動できる。
'*   ストリームの末尾を越えて現在位置を指定すると、 Stream オブジェクトの Size
'*   も拡張される。この方法で追加された新しいバイトはすべて Null になる。
'* ・[!メモ] Position は常にバイト単位である。マルチバイト文字セットを使用する
'*   テキスト ストリームでは、位置と文字サイズを乗算することによって文字数が
'*   わかる。たとえば、2 バイト文字セットの場合、最初の文字の位置は 0、2 番
'*   目の文字の位置は 2、3 番目の文字の位置は 4 のようになる。
'* ・負の値を使用して Stream 内の現在位置を変更することはできない。
'*   Position で使用できるのは正の値だけである。
'* ・読み取り専用の Stream オブジェクトの場合、 Position が Stream の Size より
'*   も大きい値に設定されていても ADO はエラーを返さない。
'*   これによって、 Stream のサイズや、 Stream の内容が変更されることはない。
'*   ただし、 Position の値が意味のない値になるので、このような操作はしないこと。
'*-----------------------------------------------------------------------------
Property Get Position() As Long
    Position = mOrigin.Position
End Property
Property Let Position(arg As Long)
    mOrigin.Position = arg
End Property

'*-----------------------------------------------------------------------------
'* Size プロパティ（読み取り専用）
'*
'* [補  足]
'* ・サイズを示す。
'*-----------------------------------------------------------------------------
Property Get Size() As Long
    Size = mOrigin.Size
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

'*-----------------------------------------------------------------------------
'* Type プロパティ（Typeが予約語のためType_に変更）
'*
'* [補  足]
'* ・Stream (バイナリまたはテキスト) に格納されているデータの種類を示す。
'* ・Stream オブジェクトに格納されたデータの種類を示す StreamTypeEnum 値を設定
'*   または取得する。既定値は adTypeText 。
'*   ただし、バイナリ データが新しい空の Stream に書き込まれる場合、 Type は
'*   adTypeBinary に変更される。
'* ・Type プロパティは、現在の位置が Stream の先頭 (Position が 0) である場合
'*   にのみ値の取得および設定が可能であり、その他の位置である場合は値の取得のみ
'*   可能。
'* ・Type プロパティによって、Stream の読み取りと書き込みに使用するメソッドが
'*   決まる。
'*   テキスト Stream の場合は、ReadText および WriteText を使用する。
'*   バイナリ Stream の場合は、Read および Write を使用する。
'*-----------------------------------------------------------------------------
Property Get Type_() As StreamTypeEnum
    Type_ = mOrigin.Type
End Property
Property Let Type_(arg As StreamTypeEnum)
    mOrigin.Type = arg
End Property

'*-----------------------------------------------------------------------------
'* 拡張プロパティ:BOM プロパティ
'*
'* [補  足]
'* BOMの有無を指定する。書込ファイルのCharsetを「UTF-8」としている場合に有効。
'*-----------------------------------------------------------------------------
Property Get BOM() As Boolean
    BOM = mBOM
End Property
Property Let BOM(arg As Boolean)
    mBOM = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Public Sub Class_Initialize()
    Set mOrigin = CreateObject("ADODB.Stream")
End Sub
Public Sub Class_Terminate()
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
'* [概  要] Cancelメソッド
'* [詳  細] Cancelのラッパーメソッド。
'*          保留中の非同期メソッド呼び出しの実行を取り消す。
'*          非同期メソッド呼び出し (つまり、adAsyncConnect、adAsyncExecute、
'*          または adAsyncFetch オプションを指定して呼び出されたメソッド) の
'*          実行を中止する。
'*          Streamの場合、Openが最後の非同期呼び出しが中止されるメソッド｡
'*
'******************************************************************************
Sub Cancel()
    Call mOrigin.Cancel
End Sub

'******************************************************************************
'* [概  要] CloseStreamメソッド
'* [詳  細] Closeのラッパーメソッド。
'*          開いているオブジェクトおよびそれに関連するすべてのオブジェクトを
'*          閉じる。
'*
'*          関連するすべてのシステム リソースを解放する場合に使用する。
'*          オブジェクトを閉じてもメモリからは削除されず、プロパティ設定を変更
'*          してもう一度開くことができる。
'*          オブジェクトをメモリから完全に削除するには、オブジェクトを閉じた後
'*          に、オブジェクト変数をNothingに設定する。
'******************************************************************************
Sub CloseStream()
    Call mOrigin.Close
End Sub

'******************************************************************************
'* [概  要] CopyToメソッド
'* [詳  細] CopyToのラッパーメソッド。
'*          Stream オブジェクトから、指定した文字数またはバイト数 (Type によっ
'*          て決まる) のデータを別の Stream オブジェクトにコピーする。
'*
'*          このメソッドは、Position プロパティで指定された現在の位置から指定
'*          数の文字またはバイトをコピーする。
'*          指定した数が、現在の位置から EOS までのバイト数を超える場合は、
'*          現在の位置から EOS までの文字またはバイトのみがコピーされる。
'*          CharNumber の値が -1 の場合、または省略されている場合は、現在の位
'*          置から始まるすべての文字またはバイトがコピーされる｡
'*          コピー先のストリームに既存の文字またはバイトが存在する場合、コピー
'*          の終了ポイント以降の内容はすべてそのまま残り、削除されない。
'*          Position は、コピーされた最後のバイトの次のバイトになる。
'*          これらのバイトを削除する場合は、 SetEOS を呼び出すこと。
'*          CopyTo メソッドは、コピー先の Stream とコピー元の Stream の種類が
'*          同じ場合 (Type プロパティの設定値がいずれも adTypeText であるか、
'*          またはいずれも adTypeBinary である場合) のコピーに使用する。
'*          Stream オブジェクトが文字列型の場合、コピー先の Stream の Charset
'*          プロパティの設定値を変更すると、文字セットを別の文字セットに変換
'*          することができる。
'*          また、文字列型の Stream オブジェクトをバイナリ型の Stream オブジェ
'*          クトにコピーすることはできるが、バイナリ型の Stream オブジェクトを
'*          文字列型の Stream オブジェクトにコピーすることはできない。
'*
'* @param DestStream 開いている Stream オブジェクトへの参照を含むオブジェクト
'*                   変数の値を指定。DestStream で指定した Stream に現在の
'*                   Stream がコピーされる。コピー先の Stream は既に開かれて
'*                   いる必要がある。開かれていない場合は、実行時エラーが発生
'*                   する。
'*                   注: deststreamパラメーターは､クライアントにリモート接続
'*                   できないstreamオブジェクト上のプライベートインターフェイス
'*                   にアクセスする必要があるため､ streamオブジェクトのプロキシ
'*                   ではない場合がある｡
'* @param CharNumber 省略可能。コピー元 Stream の現在の位置からコピー先の
'*                   Stream にコピーするバイト数または文字数を指定する整数型
'*                   (Integer) の値を指定する。既定値は -1 で、現在の位置から
'*                   EOS までのすべての文字またはバイトをコピーするよう指定する。
'******************************************************************************
Sub CopyTo(DestStream As ADODBExStream, Optional CharNumber As Long = -1)
    Call mOrigin.CopyTo(DestStream.Origin, CharNumber)
End Sub

'******************************************************************************
'* [概  要] Flushメソッド
'* [詳  細] Flushのラッパーメソッド。
'*          ADO バッファーに残っている Stream の内容を、Stream に関連付けられた
'*          基になるオブジェクトに反映します。
'*          このメソッドを使用すると、基になるオブジェクト (たとえば、 Stream
'*          オブジェクトのソースの URL で表されるノードやファイル) に、ストリー
'*          ム バッファーの内容を送ることができます。
'*          Stream の内容に加えたすべての変更が確実に書き込まれるようにするには
'*          このメソッドを呼び出す必要があります。
'*          ただし、ADO ではバックグラウンドで可能な限り継続してバッファーのフ
'*          ラッシュが行われるので、通常は Flush を呼び出す必要はありません。
'*          Stream の内容の更新は自動的に行われ、 Flush を呼び出すまでキャッシュ
'*          されるわけではありません。
'*          Close メソッドで Stream を閉じると、 Stream の内容が自動的にフラッ
'*          シュされるので、 Close の直前に明示的に Flush を呼び出す必要はあり
'*          ません。
'*
'******************************************************************************
Sub Flush()
    Call mOrigin.Flush
End Sub

'******************************************************************************
'* [概  要] LoadFromFileメソッド
'* [詳  細] LoadFromFileのラッパーメソッド。
'*          既存のファイルの内容を Stream に読み込みます。
'*
'*          このメソッドを使用して、ローカル ファイルの内容を Stream オブジェ
'*          クトに読み込むことができます。
'*          この方法で、ローカル ファイルの内容をサーバーにアップロードできます。
'*          Stream オブジェクトは LoadFromFile を呼び出す前に開かれている必要が
'*          あります。
'*          Stream オブジェクトのバインディングはこのメソッドによって変更されず
'*          その Stream を開いた URL または Record で指定されたオブジェクトに
'*          バインドされたままです。
'*          LoadFromFile では、 Stream オブジェクトの現在の内容を、ファイルから
'*          読み込んだデータで上書きします。
'*          Stream の既存のすべてのバイトは、ファイルの内容で上書きされます。
'*          LoadFromFile メソッドで作成された EOS の後に残っている既存のバイト
'*          は、すべて削除されます。
'*          LoadFromFile を呼び出した後、現在の位置は Stream の先頭に設定され、
'*          Position プロパティが 0 になります。
'*          ストリームの先頭にエンコード用の 2 バイトを追加できるので、ストリー
'*          ムのサイズは読み込み元のファイルのサイズと完全には一致しない場合が
'*          あります。
'*
'* @param FileName Stream に読み込むファイルの名前を含む文字列型 (String) の値
'*                 を指定します。FileName には UNC 形式のパスとファイル名を指定
'*                 できます。指定したファイルが存在しない場合は、実行時エラー
'*                 が発生します。
'******************************************************************************
Sub LoadFromFile(FileName As String)
    Call mOrigin.LoadFromFile(FileName)
End Sub

'******************************************************************************
'* [概  要] OpenStreamメソッド
'* [詳  細] Openのラッパーメソッド。
'*          バイナリ データまたはテキスト データのストリームを操作するために
'*          Stream オブジェクトを開きます。
'*
'*          Record オブジェクトをソース パラメーターとして渡した場合、その
'*          Record オブジェクトへのアクセスは既に有効になっているため、UserID
'*          パラメーターと Password パラメーターは使用されません。
'*          同様に、Record オブジェクトの Mode が、Stream オブジェクトに転送
'*          されます。
'*          Source を指定しない場合、開かれた Stream にはデータが含まれず、Size
'*          はゼロ (0) になります。
'*          Stream を閉じたときに、この Stream に書き込まれたデータが失われるの
'*          を防ぐには、CopyTo メソッドまたは SaveToFile メソッドを使用して
'*          Stream を保存するか、またはメモリの別の場所に保存します。
'*          OpenOptions 値の adOpenStreamFromRecord は、Source パラメーターの内
'*          容を、既に開かれた Record オブジェクトとして識別します。
'*          既定の動作では、ファイルなどのツリー構造内のノードを直接指し示す URL
'*          として Source を扱います。
'*          そのノードに関連付けられる既定のストリームが開かれます。
'*          Stream が開かれていないときは、 Stream の読み取り専用プロパティをす
'*          べて読み取ることができます。
'*          Stream が非同期で開かれている場合、以降のすべての操作 ( State プロ
'*          パティやその他の読み取り専用プロパティの確認を除く) は、 Open 操作
'*          が完了するまでブロックされます。
'*          上で説明したオプションに加えて、Source を指定しないことにより、基に
'*          なるソースに関連付けることなく、Stream オブジェクトのインスタンスを
'*          メモリ内に作成することができます。ストリームにデータを動的に追加する
'*          には、Write または WriteText を使用してバイナリ データまたはテキスト
'*          データを Stream に書き込むか、または LoadFromFile を使用してデータを
'*          ファイルから読み込みます。
'*
'* @param Source   省略可能。 ストリームのデータソースを指定するバリアント型
'*                 (Variant) の値です。
'*                 Sourceには、電子メールやファイルシステムなど、既知のツリー
'*                 構造内の既存のノードを指す絶対 URL 文字列を含めることがで
'*                 きます。 url は、url キーワード ("url =scheme//サーバー/フ
'*                 ォルダー") を使用して指定する必要があります。
'*                 または、 Sourceには既に開いているrecordオブジェクトへの参照が
'*                 含まれている場合があります。
'*                 これにより、 recordに関連付けられている既定のストリームが開
'*                 きます。 sourceが指定されていない場合、既定では、基になるソ
'*                 ースに関連付けられていないストリームがインスタンス化されて
'*                 開きます。
'* @param Mode     省略可能。結果の Stream に対するアクセス モード (読み取り/書
'*                 き込み、読み取り専用など) を ConnectModeEnum 値で指定します。
'*                 既定値は adModeUnknown です。アクセス モードの詳細については
'*                 「Mode プロパティ (ADO)」を参照してください。
'*                 Mode を指定しない場合は、ソース オブジェクトから継承されます。
'*                 たとえば、ソース Record が読み取り専用モードで開かれている場
'*                 合、既定では、その Stream も読み取り専用モードで開かれます。
'* @param Options  省略可能。StreamOpenOptionsEnum 値を指定します。既定値は
'*                 adOpenStreamUnspecified です。
'* @param UserName 省略可能。必要に応じて、 Stream オブジェクトにアクセスするた
'*                 めのユーザー ID を含む、文字列型 ( String ) の値を指定します。
'* @param Password 省略可能。必要に応じて、 Stream オブジェクトにアクセスするた
'*                 めのパスワードを含む、文字列型 ( String ) の値を指定します。
'******************************************************************************
Sub OpenStream(Optional Source, _
               Optional Mode As ConnectModeEnum = adModeUnknown, _
               Optional Options As StreamOpenOptionsEnum = adOpenStreamUnspecified, _
               Optional UserName As String, _
               Optional Password As String)
    If IsMissing(Source) Then
        Call mOrigin.Open(, Mode, Options, UserName, Password)
    Else
        Call mOrigin.Open(Source, Mode, Options, UserName, Password)
    End If
End Sub

'******************************************************************************
'* [概  要] Readメソッド
'* [詳  細] Readのラッパーメソッド。
'*          バイナリ型の Stream オブジェクトから、指定されたバイト数を読み取り
'*          ます。
'*
'*          NumBytes が Stream に残っているバイト数よりも大きい場合、残っている
'*          バイトのみが返されます。
'*          読み取ったデータに、NumBytes で指定した長さに合うようにスペースが補
'*          充されることはありません。
'*          読み取るデータが残っていない場合は、Null 値のバリアント型が返されま
'*          す。Read を使用して逆方向の読み取りを行うことはできません。
'*          NumBytes は常にバイトを表します。 テキストの Stream オブジェクト
'*          (Type が adTypeText) の場合は、ReadText を使用してください。
'*
'* @param NumBytes 省略可能。ファイルから読み取るバイト数を指定する長整数型
'*                 ( Long ) の値、または StreamReadEnum 値の adReadAll (既定
'*                 値) を指定します。
'* @return Read メソッドは、指定したバイト数またはストリーム全体を Stream オブ
'*              ジェクトから読み取り、結果のデータをバリアント型 (Variant) と
'*              して返します。
'******************************************************************************
Function Read(Optional NumBytes As Long = -1)
    Read = mOrigin.Read(NumBytes)
End Function

'******************************************************************************
'* [概  要] ReadTextメソッド
'* [詳  細] ReadTextのラッパーメソッド。
'*          文字列型の Stream オブジェクトから、指定された文字数を読み取ります。
'*
'*          NumChar がストリームに残っている文字の数よりも大きい場合、残ってい
'*          る文字のみが返されます。
'*          読み取った文字列に、NumChar で指定した長さに合うようにスペースが補
'*          充されることはありません。
'*          読み取る文字が残っていない場合は、値が Null のバリアント型が返され
'*          ます。
'*          ReadText は、逆方向の読み取りに使用することはできません。
'*          ReadText メソッドは、文字列型ストリーム (Type が adTypeText) で使用
'*          します。 バイナリ型のストリーム (Type が adTypeBinary) の場合は、
'*          Read を使用してください。
'*
'* @param NumChars 省略可能。ファイルから読み取る文字の数を指定する長整数型
'*                 ( Long ) の値、または StreamReadEnum 値を指定します。
'*                 既定値は adReadAll です。
'* @return 指定した文字数、行全体、またはストリーム全体を Stream オブジェクト
'*         から読み取り、結果の文字列を返します。
'******************************************************************************
Function ReadText(Optional NumChars As Long = -1) As String
    ReadText = mOrigin.ReadText(NumChars)
End Function

'******************************************************************************
'* [概  要] SaveToFileメソッド
'* [詳  細] SaveToFileのラッパーメソッド。
'*          Stream のバイナリの内容をファイルに保存します｡
'*
'*          SaveToFile メソッドを使用すると、 Stream オブジェクトの内容をローカ
'*          ル ファイルにコピーできます。
'*          Stream オブジェクトの内容やプロパティは変更されません。
'*          SaveToFile メソッドを呼び出す前に、 Stream オブジェクトが開かれてい
'*          る必要があります。
'*          このメソッドでは、 Stream オブジェクトと基になるソースの関連付けは
'*          変更されません。
'*          Stream オブジェクトは、開かれたときのソースである元の URL または
'*          Record に関連付けられたままになります。
'*          SaveToFile メソッドの操作後、ストリーム内の現在の位置 (Position) は、
'*          ストリームの先頭 (0) に設定されます。
'*
'* @param FileName Stream の内容の保存先であるファイルの完全修飾名を含む文字列
'*                 型 ( String ) の値を指定します。
'*                 有効なローカルの場所、または UNC 値を介してアクセスできる場
'*                 所への保存が可能です。
'* @param Options  保存するファイルがまだ存在しない場合に、 SaveToFile メソッド
'*                 で新しいファイルを作成するかどうかを SaveOptionsEnum 値で指定
'*                 します。既定値は adSaveCreateNotExists です。
'*                 これらのオプションでは、指定したファイルが存在しない場合にエラ
'*                 ーが発生するように指定できます。
'*                 また、 SaveToFile メソッドで既存ファイルの現在の内容を上書き
'*                 するように指定することもできます。
'*                 [!メモ] adSaveCreateOverwrite を設定して既存ファイルを上書き
'*                 すると、 SaveToFile メソッドによって、元の既存ファイルから、
'*                 新しい EOS 以降のバイトがすべて削除されます。
'******************************************************************************
Sub SaveToFile(FileName As String, Optional Options As SaveOptionsEnum = adSaveCreateNotExist)
    
    '拡張処理:BOM除外対応
    With Me
        If UCase(.CharSet) = "UTF-8" And Not .BOM Then
            .ExcludeBom
        End If
    End With
        
    Call mOrigin.SaveToFile(FileName, Options)
End Sub

'******************************************************************************
'* [概  要] SetEOSメソッド
'* [詳  細] SetEOSのラッパーメソッド。
'*          ストリームの末尾の位置を設定します｡
'*
'*          SetEOS は、現在の Position がストリームの末尾になるように EOS
'*          プロパティの値を更新します。
'*          現在の位置より後ろにあったバイト値や文字はすべて切り捨てられます。
'*          Write、WriteText、および CopyTo の各メソッドでは、既存の Stream
'*          オブジェクト内の余分な値は切り捨てられないため、これらのバイト値
'*          や文字を切り捨てるには、 SetEOS でストリームの新しい末尾の位置を
'*          設定してください。
'*          EOS をストリームの実際の末尾より前の位置に設定すると、新しい EOS
'*          位置より後ろのデータはすべて失われます。
'*
'******************************************************************************
Sub SetEOS()
    Call mOrigin.SetEOS
End Sub

'******************************************************************************
'* [概  要] SkipLineメソッド
'* [詳  細] SkipLineのラッパーメソッド。
'*          テキスト ストリームを読み取っているときに、1 行全体をスキップします。
'*
'*          次の行区切り文字までのすべての文字 (行区切り文字を含む) がスキップ
'*          されます。
'*          既定では、LineSeparator は adCRLF です。EOS を越えてスキップしよう
'*          とした場合は、EOS が現在の位置になります。
'*          SkipLine メソッドは、テキスト ストリーム (Type が adTypeText のスト
'*          リーム) で使用します。
'*
'******************************************************************************
Sub SkipLine()
    Call mOrigin.SkipLine
End Sub

'******************************************************************************
'* [概  要] WriteBufメソッド
'* [詳  細] Writeのラッパーメソッド。
'*          バイナリデータを Stream オブジェクトに書き込みます。
'*
'*          指定したバイトが、各バイト間にスペースを一切挿入することなく Stream
'*          オブジェクトに書き込まれます。
'*          カレント Position は、書き込まれたデータの次のバイトに設定されます。
'*          Write メソッドがストリーム内の残りのデータを切り捨てることはありま
'*          せん。
'*          後ろのバイトを切り捨てるには、 SetEOS を呼び出してください。
'*          現在の EOS 位置を越えて書き込みを行うと、新しいバイトがすべて格納で
'*          きるように Stream の Size が大きくなり、 EOS が Stream 内の新しい
'*          末尾バイトへと移動します。
'*          Write メソッドは、バイナリ ストリーム (Type が adTypeBinary) で使用
'*          します。テキスト ストリーム (Type が adTypeText) の場合は、WriteText
'*          を使用してください。
'*
'* @param Buffer 書き込むバイト配列の入ったバリアント型 ( Variant ) の値を指定。
'******************************************************************************
Sub WriteBuf(Buffer)
    Call mOrigin.Write(Buffer)
End Sub

'******************************************************************************
'* [概  要] WriteTextメソッド
'* [詳  細] WriteTextのラッパーメソッド。
'*          指定されたテキスト文字列を Stream オブジェクトに書き込みます。
'*
'*          指定した文字列が、各文字列間にスペースや文字を挿入することなく
'*          Stream オブジェクトに書き込まれます。
'*          カレント Position は、書き込まれたデータの次の文字に設定されます。
'*          WriteText メソッドがストリーム内の残りのデータを切り捨てることは
'*          ありません。
'*          後ろの文字を切り捨てるには、 SetEOS を呼び出してください。
'*          現在の EOS 位置を越えて書き込みを行うと、新しい文字がすべて格納
'*          できるように Stream の Size が大きくなり、 EOS が Stream 内の新
'*          しい末尾バイトへと移動します。
'*          WriteText メソッドは、テキスト ストリーム (Type が adTypeText)
'*          で使用します。
'*          バイナリ ストリーム (Type が adTypeBinary) の場合は、Write を使用
'*          してください。
'*
'* @param Data    書き込むテキストが格納された文字列型 ( String ) の値を指定。
'* @param Options 省略可能。指定された文字列の末尾に行区切り文字を書き込むか
'*                どうかを指定する StreamWriteEnum 値を指定します。
'******************************************************************************
Sub WriteText(data As String, Optional Options As StreamWriteEnum = adWriteChar)
    Call mOrigin.WriteText(data, Options)
End Sub


'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] URLエンコード処理。
'* [詳  細] URLエンコードを行う。
'*          ScriptControlも関数も使用せず、自前でエンコードしているので
'*          どのExcelバージョン・32bit/64bitでも使用可能。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Public Function EncodeUrl(target As String, Optional cSet As String = "UTF-8") As String
    Dim buf() As Byte, s As String, i As Long
    
    buf = String2Binary(target, cSet)
    
    For i = 0 To UBound(buf)
        Dim flg As Boolean: flg = False
        Select Case buf(i)
            Case 45, 46, 95, 126 '-._~
                flg = True
            Case 48 To 57 '0-9
                flg = True
            Case 65 To 90 'A-Z
                flg = True
            Case 97 To 122 'a-z
                flg = True
        End Select
        If flg Then
            s = s & Chr(buf(i))
        Else
            s = s & "%" & Hex(buf(i))
        End If
    Next
    EncodeUrl = s
End Function

'******************************************************************************
'* [概  要] URLデコード処理。
'* [詳  細] URLデコードを行う。
'*          ScriptControlも関数も使用せず、自前でデコードしているので
'*          どのExcelバージョン・32bit/64bitでも使用可能。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Public Function DecodeUrl(target As String, Optional cSet As String = "UTF-8") As String
    Dim bin() As Byte, var, i As Long, j As Long, cnt As Long
    Dim buf() As Byte
    Dim tmp As String
    
    ' %で区切って処理
    var = Split(target, "%")
    
    ' Byte配列の要素数計算
    For i = LBound(var) To UBound(var)
        tmp = CStr(var(i))
        If i = LBound(var) And Left(target, 1) <> "%" Then
            For j = 1 To Len(tmp)
               cnt = cnt + 1
            Next j
        ElseIf Len(tmp) = 2 Then
            cnt = cnt + 1
        Else
            cnt = cnt + 1
            For j = 3 To Len(tmp)
               cnt = cnt + 1
            Next j
        End If
    Next i
    ReDim bin(0 To cnt - 1) As Byte
    cnt = 0
    
    ' Byteデータに変換
    For i = LBound(var) To UBound(var)
        tmp = CStr(var(i))
        
        If i = LBound(var) And Left(target, 1) <> "%" Then
            For j = 1 To Len(tmp)
               bin(cnt) = Asc(Mid(tmp, j, 1))
               cnt = cnt + 1
            Next j
        ElseIf Len(tmp) = 2 Then
            bin(cnt) = CByte(Val("&H" & tmp))
            cnt = cnt + 1
        Else
            bin(cnt) = CByte(Val("&H" & Mid(tmp, 1, 2)))
            cnt = cnt + 1
            For j = 3 To Len(tmp)
               bin(cnt) = Asc(Mid(tmp, j, 1))
               cnt = cnt + 1
            Next j
        End If
    Next i
    
    ' Byteから文字列へ変換
    DecodeUrl = Binary2String(bin, cSet)
End Function

'******************************************************************************
'* [概  要] Base64エンコード処理。
'* [詳  細] Base64エンコードを行う。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Public Function EncodeBase64(target As String, Optional cSet As String = "UTF-8") As String
    Dim text As String
    With CreateObject("Msxml2.DOMDocument.6.0")
        With .createElement("tmp")
            .DataType = "bin.base64"
            .NodeTypedValue = String2Binary(target, cSet)
            text = .text
        End With
    End With
    
    EncodeBase64 = text
End Function

'******************************************************************************
'* [概  要] Base64デコード処理。
'* [詳  細] Base64デコードを行う。
'* [参　考] <https://defun.exblog.jp/5309470/>
'*
'* @param target 対象となる文字列
'* @return デコード後文字列
'******************************************************************************
Public Function DecodeBase64(target As String, Optional cSet As String = "UTF-8") As String
    Dim bin() As Byte
    With CreateObject("Msxml2.DOMDocument.6.0")
        With .createElement("tmp")
            .DataType = "bin.base64"
            .text = target
            bin = .NodeTypedValue
        End With
    End With
    
    DecodeBase64 = Binary2String(bin, cSet)
End Function

'******************************************************************************
'* [概  要] 文字列→バイトデータ変換処理。
'* [詳  細] 文字列からバイトデータへ変換を行う。
'*
'* @param text 対象となる文字列
'* @param cSet キャラセット（文字コード指定）。省略時は「UTF-8」。
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @return 変換後バイトデータ
'******************************************************************************
Public Function String2Binary(text As String, Optional cSet As String = "UTF-8") As Byte()
    With Me
        .Mode = adModeReadWrite
        .CharSet = cSet
        .Type_ = adTypeText
        .BOM = False
        .OpenStream
        .WriteText text
        If UCase(.CharSet) = "UTF-8" And Not .BOM Then
            String2Binary = .ExcludeBom
        Else
            .Position = 0
            .Type_ = adTypeBinary
            String2Binary = .Read
        End If
        .CloseStream
    End With
End Function

'******************************************************************************
'* [概  要] バイトデータ→文字列変換処理。
'* [詳  細] バイトデータから文字列へ変換を行う。
'*
'* @param text 対象となるバイトデータ
'* @param cSet キャラセット（文字コード指定）。省略時は「UTF-8」。
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @return 変換後文字列
'******************************************************************************
Public Function Binary2String(bin() As Byte, Optional cSet As String = "UTF-8") As String
    With Me
        .Mode = adModeReadWrite
        .CharSet = cSet
        .Type_ = adTypeBinary
        .OpenStream
        .WriteBuf bin
        .Position = 0
        .Type_ = adTypeText
        Binary2String = .ReadText
        .CloseStream
    End With
End Function

'******************************************************************************
'* [概  要] BOM除外処理。
'* [詳  細] UTF-8テキストのBOMを除外する。
'* [備　考] 呼び出す前にStreamがOpenされていること。
'*          WriteText呼出し直後、またはファイル保存前に実行すること。
'*          （TODO:大容量データの考慮）
'*
'* @return BOM除外データ
'******************************************************************************
Public Function ExcludeBom() As Byte()
    With Me
        .Position = 0
        .Type_ = adTypeBinary
        .Position = 3 'BOM飛ばし
        ExcludeBom = .Read
        .Position = 0
        .WriteBuf ExcludeBom
        .SetEOS
    End With
End Function

'******************************************************************************
'* [概  要] テキストファイル読込処理。
'* [詳  細] テキストファイルの読込を行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param cSet キャラセット（文字コード指定）
'*             省略時は自動判別（"_autodetect_all"）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @return 読込テキストデータ
'******************************************************************************
Public Function ReadTextFile(filePath As String, Optional cSet As String = "_autodetect_all") As String
    Dim data As String
    With Me
        .Mode = adModeReadWrite
        .Type_ = adTypeText
        .CharSet = cSet
        .OpenStream

        .LoadFromFile filePath 'ファイル指定
        .Position = 0
        data = .ReadText()
        .CloseStream
    End With

    ReadTextFile = data
End Function

'******************************************************************************
'* [概  要] UTF-8テキストファイル読込処理。
'* [詳  細] UTF-8テキストファイルの読込を行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @return 読込テキストデータ
'******************************************************************************
Public Function ReadUTF8TextFile(filePath As String) As String
    ReadUTF8TextFile = ReadTextFile(filePath, "UTF-8")
End Function

'******************************************************************************
'* [概  要] テキストファイル読込処理。
'* [詳  細] テキストファイルを行単位で読込み、Variant配列に設定して返却する。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param cSet キャラセット（文字コード指定）
'*             省略時は自動判別（"_autodetect_all"）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @param separator 改行コード（省略可。規定はCRLF）
'* @return 読込テキストデータ（Variant配列）
'******************************************************************************
Public Function ReadTextFileToVArray(filePath As String, _
                                     Optional cSet As String = "_autodetect_all", _
                                     Optional separator As LineSeparatorEnum = adCRLF) As Variant
    Dim vArr(), Size As Long, Count As Long
    With Me
        .Mode = adModeReadWrite
        .Type_ = adTypeText
        .CharSet = cSet
        .LineSeparator = separator
        .OpenStream
        .LoadFromFile filePath 'ファイル指定
        Size = .CountLine
        ReDim Preserve vArr(Size - 1)
        .Position = 0
        Do While Not .EOS
            vArr(Count) = .ReadText(adReadLine)
            Count = Count + 1
        Loop
        .CloseStream
    End With

    ReadTextFileToVArray = vArr
End Function

'******************************************************************************
'* [概  要] テキストの行数取得処理。
'* [詳  細] テキストの行数を取得する。
'* [備　考] 呼び出す前にStreamがOpenされていること。
'*
'* @return 行数
'******************************************************************************
Public Function CountLine() As Long
    Dim lineCount As Long
    With Me
        .Position = 0
        Do While Not .EOS
            .SkipLine
            lineCount = lineCount + 1
        Loop
    End With

    CountLine = lineCount
End Function
  
'******************************************************************************
'* [概  要] テキストファイル行数取得処理。
'* [詳  細] テキストファイルを読込み、行数を返却する。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param cSet キャラセット（文字コード指定）
'*             省略時は自動判別（"_autodetect_all"）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @param separator 改行コード（省略可。規定はCRLF）
'* @return 行数
'******************************************************************************
Public Function ReadTextFileLineCount(filePath As String, _
                                      Optional cSet As String = "_autodetect_all", _
                                      Optional separator As LineSeparatorEnum = adCRLF) As Long
    Dim Size As Long
    With Me
        .Mode = adModeReadWrite
        .Type_ = adTypeText
        .CharSet = cSet
        .LineSeparator = separator
        .OpenStream
        .LoadFromFile filePath 'ファイル指定
        Size = .CountLine
        .CloseStream
    End With

    ReadTextFileLineCount = Size
End Function
  
'******************************************************************************
'* [概  要] テキストファイル書込処理。
'* [詳  細] テキストファイルの書込を行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param data 書込テキストデータ
'* @param cSet キャラセット（文字コード指定）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @param append 追記有無（省略可。規定はFalse:追記なし）
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub WriteTextFile(filePath As String, data As String, cSet As String, _
                         Optional Append As Boolean = False, _
                         Optional bomInclude As Boolean = False)
    With Me
        .Mode = adModeReadWrite
        .Type_ = adTypeText
        .CharSet = cSet
        .BOM = bomInclude
        .OpenStream
        If Append Then
            .AppendMode filePath
        End If
        .WriteText data, adWriteChar
        .SaveToFile filePath, adSaveCreateOverWrite 'ファイル上書指定
        .CloseStream
    End With
End Sub

'******************************************************************************
'* [概  要] UTF-8テキストファイル書込処理。
'* [詳  細] UTF-8テキストファイルの書込を行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param data 書込テキストデータ
'* @param append 追記有無（省略可。規定はFalse:追記なし）
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub WriteUTF8TextFile(filePath As String, data As String, _
                             Optional Append As Boolean = False, _
                             Optional bomInclude As Boolean = False)
     Call WriteTextFile(filePath, data, "UTF-8", Append, bomInclude)
End Sub

'******************************************************************************
'* [概  要] テキストファイル書込処理。
'* [詳  細] テキストファイルに配列データを行単位で書込を行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param vArray 書込配列データ（Variant）
'* @param cSet キャラセット（文字コード指定）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @param separator 改行コード（省略可。規定はCRLF）
'* @param append 追記有無（省略可。規定はFalse:追記なし）
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub WriteTextFileFromVArray(filePath As String, vArray As Variant, cSet As String, _
                                   Optional separator As LineSeparatorEnum = adCRLF, _
                                   Optional Append As Boolean = False, _
                                   Optional bomInclude As Boolean = False)
    Dim tmp
    With Me
        .Mode = adModeReadWrite
        .Type_ = adTypeText
        .CharSet = cSet
        .BOM = bomInclude
        .LineSeparator = separator
        .OpenStream
        If Append Then
            .AppendMode filePath
        End If
        For Each tmp In vArray
            .WriteText CStr(tmp), adWriteLine
        Next
        .SaveToFile filePath, adSaveCreateOverWrite 'ファイル上書指定
        .CloseStream
    End With
End Sub

'******************************************************************************
'* [概  要] ファイルアペンド処理。
'* [詳  細] ファイル追記するための処理を行う。
'*          ファイルの存在チェックを行い、存在する場合は対象ファイルの
'*          ポインタを終端へ移動する。
'* [備　考] 呼び出す前にStreamがOpenされていること。
'*
'* @param filePath 書込み対象となるテキストファイルのフルパス
'******************************************************************************
Public Sub AppendMode(filePath As String)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        With Me
            .LoadFromFile filePath
            .Position = .Size 'ポインタを終端へ
        End With
    End If
End Sub

'******************************************************************************
'* [概  要] BOM有無チェック処理。
'* [詳  細] UTF-8テキストファイルにBOMがあるかチェックを行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param チェック結果（true:BOM有、false:BOM無）
'******************************************************************************
Function IsBomIncluded(filePath As String) As Boolean
    Dim buf() As Byte
    Dim suspected As String
    Dim v As Variant

    With Me
        .OpenStream
        .Type_ = adTypeBinary
        .LoadFromFile filePath
        buf = .Read(3) 'BOMにあたる先頭3バイトを取得
        .CloseStream
    End With

    For Each v In buf
        suspected = suspected & Hex(v)
    Next
    
    If suspected = "EFBBBF" Then
        IsBomIncluded = True
    Else
        IsBomIncluded = False
    End If
End Function

'******************************************************************************
'* [概  要] ファイルダンプ処理。
'* [詳  細] ファイルを読込み、16進ダンプ文字データを返却する。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @return 16進ダンプ文字データ
'******************************************************************************
Public Function ReadFileToDump(filePath As String) As String
    Dim ret As String, buf As String, i As Long
    With Me
        .Mode = adModeReadWrite
        .Type_ = adTypeBinary
        .OpenStream
        .LoadFromFile filePath 'ファイル指定
        Do While Not .EOS
            buf = .Read(16)
 
            For i = 1 To LenB(buf)
                Dim tmp As String
                tmp = MidB(buf, i, 1)
                tmp = AscB(tmp)
                tmp = Hex(tmp)
                tmp = UCase(tmp)
                tmp = Right("0" & tmp, 2)
                ret = ret & tmp & " "
            Next
            ret = ret & vbCrLf
        Loop
        .CloseStream
    End With

    ReadFileToDump = ret
End Function

