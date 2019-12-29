VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODB.Errorクラスのラップ・拡張クラス
'* [詳  細] ADODB.ErrorクラスのWrapperとしての機能を提供する他、
'*          ADODB.Errorクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Errorクラスについて。
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Contains details about data access errors that pertain to a single operation
' involving the provider.
'
' プロバイダーが関与する単一の操作に関連するデータアクセスエラーに関する詳細が
' 含まれる。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Error
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
'* Description プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・エラーの内容を表す文字列の値を返却する。
'* ・Description プロパティは、エラーの簡単な説明を取得するために使う。
'*   プログラムで対処できないエラー､または処理することが望ましくないエラーは
'*   このプロパティの内容を表示してユーザーに警告する｡
'*   文字列は､ADO またはプロバイダーから渡される｡
'* ・プロバイダーは、特定のエラー テキストを ADO に渡す。
'*   ADO は、受け取ったプロバイダー エラーまたは警告ごとに Error オブジェク
'*   トを Errors コレクションに追加する。
'*   プロバイダーが渡すエラーをトレースするには、 Errors コレクションを列挙
'*   する。
'*-----------------------------------------------------------------------------
Property Get Description() As String
Attribute Description.VB_UserMemId = 0
    Description = mOrigin.Description
End Property

'*-----------------------------------------------------------------------------
'* HelpContext プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ヘルプファイルのトピックのコンテキストIDを返却する。
'* ・HelpFile プロパティでヘルプ ファイルが指定されている場合、HelpContext
'*   プロパティを使って特定のヘルプ トピックを自動的に表示できる。
'*   該当するヘルプ トピックがない場合、HelpContext プロパティは 0 を返し、
'*   HelpFile プロパティは長さ 0 の文字列 ("") を返す。
'*-----------------------------------------------------------------------------
Property Get HelpContext() As Long
    HelpContext = mOrigin.HelpContext
End Property

'*-----------------------------------------------------------------------------
'* HelpFile プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ヘルプファイルへの完全なパスを返却する。
'* ・HelpFile プロパティでヘルプ ファイルが指定されている場合、HelpContext
'*   プロパティを使って特定のヘルプ トピックを自動的に表示できる。
'*   該当するヘルプ トピックがない場合、HelpContext プロパティは 0 を返し、
'*   HelpFile プロパティは長さ 0 の文字列 ("") を返す。
'*-----------------------------------------------------------------------------
Property Get HelpFile() As String
    HelpFile = mOrigin.HelpFile
End Property

'*-----------------------------------------------------------------------------
'* NativeError プロパティ（読み取り専用）
'*
'* [補  足]
'* ・指定された Error オブジェクトでプロバイダー固有のエラーコードを示す。
'* ・NativeError プロパティは、特定の Error オブジェクトの、データベース固有の
'*   エラー情報を取得するために使用する。
'*   たとえば、Microsoft ODBC Provider for OLE DB と Microsoft SQL Server デー
'*   タベースを使う場合、SQL Server から送信されたネイティブエラーコードは、
'*   ODBC と ODBC Provider を経由して ADO の NativeError プロパティに渡される。
'*-----------------------------------------------------------------------------
Property Get NativeError() As Long
    NativeError = mOrigin.NativeError
End Property

'*-----------------------------------------------------------------------------
'* Number プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Error オブジェクトを一意に識別する数値を示す。
'* ・ErrorValueEnum 定数のいずれかに対応する値を返却する｡
'* ・Number プロパティは､発生したエラーを調べるために使用する｡
'*   プロパティの値は､エラー条件に対応した一意な数値である｡
'* ・Errorsコレクションは、16進形式（0x80004005など）またはlong値（2147467259
'*   など）でHRESULTを返す。
'*   これらのHRESULTは、OLE DBまたはOLE自体などの基礎となるコンポーネントによ
'*   って発生する可能性がある。
'*-----------------------------------------------------------------------------
Property Get Number() As Long
    Number = mOrigin.Number
End Property

'*-----------------------------------------------------------------------------
'* Source プロパティ（読み取り専用）
'*
'* [補  足]
'* ・エラーの発生源のオブジェクトまたはアプリケーションの名前を示す。
'* ・エラーが発生した元のオブジェクトまたはアプリケーションの名前を確認するに
'*   は、 errorオブジェクトのSourceプロパティを使用する。
'*   これは、オブジェクトのクラス名またはプログラム ID である可能性がある。
'*   ADO のエラーについては、プロパティの値は **ADODB. *** (objectname) 。
'*   objectnameは、エラーを発生させたオブジェクトの名前である。
'*   ADOX および ADO MD の場合、値はそれぞれ *adox. * * * objectnameおよび
'*   *adomd.nethttp * * objectnameとなる。
'* ・Error オブジェクトの Source プロパティ、 Number プロパティ、および
'*   Description プロパティのエラー情報に基づいて、エラーを適切に処理するため
'*   のコードを書くことができる。
'* ・Error オブジェクトの Source プロパティは読み取り専用。
'*-----------------------------------------------------------------------------
Property Get Source() As String
    Source = mOrigin.Source
End Property

'*-----------------------------------------------------------------------------
'* SQLState プロパティ（読み取り専用）
'*
'* [補  足]
'* ・特定の Error オブジェクトの SQL 状態を示す。
'* ・ANSI SQL 標準に準拠し、エラーコードを示す 5 文字の文字列型の値を返す。
'* ・SQL ステートメントの処理中にエラーが発生した場合に、プロバイダーが返す
'*   5 文字のエラーコードを取得するには SQLState プロパティを使用する。
'*   たとえば、Microsoft OLE DB Provider for ODBC を Microsoft SQL Server デー
'*   タベースと共に使用する場合、ODBC に固有のエラー、または Microsoft SQL
'*   Server に起因するエラーに基づいて、ODBC から SQL 状態のエラーコードが生成
'*   された後、ODBC エラーにマップされる。
'*   これらのエラーコードは ANSI SQL 標準で規定されているが、実装方法はデータ
'*   ソースによって異なる場合がある。
'*-----------------------------------------------------------------------------
Property Get SQLState() As String
    SQLState = mOrigin.SQLState
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("ADODB.Error")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
