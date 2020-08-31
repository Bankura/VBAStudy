VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Databaseユーティリティクラス
'* [詳  細] Databaseに関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考] http://kojama.la.coocan.jp/works/rdbms/conn/connstr.html
'*          https://excel-ubara.com/excelvba4/EXCEL273.html
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************
'Oracle ODBC接続(Microsoft製Driver Oracle製は{Oracle in OraHomexx} xxはバージョンにより変わる)
Private Const CONN_STR_ODBC_ORACLE As String = "Driver={Microsoft ODBC for Oracle};Server={0}"
'Microsoft SQL Server ODBC接続
Private Const CONN_STR_ODBC_SQLSERVER As String = "Driver={SQL Server};Server={0};Database={1}"
'Microsoft SQL Server LocalDB ODBC接続
Private Const CONN_STR_ODBC_SQLSERVER_LOCALDB As String = "Driver={SQL Server Native Client 11.0};Server=(LocalDB)\v11.0;AttachDBFileName={0};Database={1};Trusted_Connection=Yes"
'PostgreSQL ODBC接続
Private Const CONN_STR_ODBC_POSTGRESQL As String = "Driver={PostgreSQL};Server={0};Port={1};Database={2}"
'MySQL ODBC接続
Private Const CONN_STR_ODBC_MYSQL As String = "Driver={mySQL ODBC {0} Driver};Server={1};Port={2};Database={3}"
'Access ODBC接続
Private Const CONN_STR_ODBC_ACCESS As String = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ={0}"
'SQLite3 ODBC接続
Private Const CONN_STR_ODBC_SQLITE3 As String = "Driver=SQLite3 ODBC Driver;Database={0}"

'Oracle OLE接続(Microsoft製Provider Oracle製はOraOLEDB.Oracle)
Private Const CONN_STR_OLE_ORACLE As String = "Provider= MSDAORA;Data Source={0}"
'Microsoft SQL Server OLE接続
Private Const CONN_STR_OLE_SQLSERVER As String = "Provider=SQLOLEDB;Data Source={0};Initial Catalog={1}"
'PostgreSQL OLE接続
Private Const CONN_STR_OLE_POSTGRESQL As String = "Provider=PostgreSQL OLE DB Provider;Data Source={0};Location={1}"
'Access OLE接続
Private Const CONN_STR_OLE_ACCESS As String = "Provider={0};Data Source={1}"
'Excel2007 OLE接続
Private Const CONN_STR_OLE_EXCEL As String = "Provider={0};Data Source={1};Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
'Excel2003 OLE接続
Private Const CONN_STR_OLE_XL2003 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
'CSV OLE接続
Private Const CONN_STR_OLE_CSV As String = "Provider={0};Data Source={1};Extended Properties=""text;Text;HDR=Yes;FMT=Delimited"""
'HTML OLE接続
Private Const CONN_STR_OLE_HTML As String = "Provider={0};Data Source={0};Extended Properties=""HTML Import;HDR=Yes;"""


'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mConn As Object

'******************************************************************************
'* プロパティ定義
'******************************************************************************


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mConn = CreateObject("ADODB.Connection")
End Sub
Private Sub Class_Terminate()
    If Not mConn Is Nothing Then
        If mConn.State <> adStateClosed Then
            Call mConn.Close
        End If
    End If
    Set mConn = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] OpenPostgreSql メソッド
'* [詳  細] PostgreSqlデータベースへの接続を開く｡
'*          PostgreSql接続文字列を作成し、Connection.Open メソッドを実行する。
'*
'* @param serverName   接続先サーバ名
'* @param portNo       接続先ポート番号
'* @param databaseName 接続先データベース名
'* @param userId       省略可能。接続を確立するときに使用するユーザー名。
'* @param password     省略可能。接続を確立するときに使用するパスワード。
'* @param options      省略可能。接続が確立された後 (同期) または接続が
'*                     確立される前 (非同期) のどちらでこのメソッドが返る
'*                     かを指定する ConnectOptionEnum 値。
'******************************************************************************
Public Sub OpenPostgreSql(serverName As String, _
                          portNo As Integer, _
                          databaseName As String, _
                          Optional userId As String, _
                          Optional password As String, _
                          Optional options As Long = -1)
    Dim cnnStr As String
    cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_POSTGRESQL, serverName, portNo, databaseName)
    Call mConn.Open(cnnStr, userId, password, options)
End Sub

Public Sub OpenAccess(databaseName As String, _
                      Optional userId As String, _
                      Optional password As String, _
                      Optional options As Long = -1)
    Dim cnnStr As String
    cnnStr = StringUtils.PlaceHolder(CONN_STR_OLE_ACCESS, GetAceOleDbProviderName(), databaseName)
    Call mConn.Open(cnnStr, userId, password, options)
End Sub

Private Function GetAceOleDbProviderName() As String
    Dim vArr, vMsAceArr
    Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
    stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Classes", vArr
    vMsAceArr = ArrayUtils.Search(vArr, "Microsoft.ACE.OLEDB")
    If UBound(vMsAceArr) > 0 Then
        GetAceOleDbProviderName = vMsAceArr(UBound(vMsAceArr))
    Else
        err.Raise 9999, "GetAceOleDbProviderName", "Microsoft.ACE.OLEDB Not Found"
    End If
End Function

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
'Public Function PrintSql() As String
'    Dim p As Variant
'    Dim v As Variant
'    Dim i As Integer
'    Dim s As String
'    i = 1
'    ReDim v(Me.Parameters.Count)
'    For Each p In mOrigin.Parameters '★
'        Select Case p.Type
'            'ここはニーズに応じてタイプを追加
'            Case adInteger
'                v(i) = p.Value
'            Case Else
'                v(i) = "'" & p.Value & "'"
'        End Select
'        i = i + 1
'    Next
'    s = mOrigin.CommandText '★
'    For i = 1 To UBound(v)
'        s = Replace(s, "?", v(i), 1, 1)
'    Next
'    PrintSql = s
'End Function


