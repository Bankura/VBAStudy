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
'* [依  存] ODBCドライバ
'*          <Oracle ODBC Driver (Oracle Instant Client同梱)>
'*            https://www.oracle.com/database/technologies/instant-client/downloads.html
'*          <Microsoft ODBC for Oracle (32bitのみ:Oracle 8 まで。Oracle版を推奨)>
'*            Windows同梱。以下によると、今後のバージョンの Windows では削除される予定のため使用しない。
'*            https://docs.microsoft.com/ja-jp/sql/odbc/microsoft/odbc-driver-for-oracle?view=sql-server-ver15
'*          <ODBC Driver for SQL Server>
'*            https://docs.microsoft.com/ja-jp/sql/connect/odbc/download-odbc-driver-for-sql-server
'*          <psqlODBC (PostgreSQL ODBC Driver)>
'*            https://www.postgresql.org/ftp/odbc/versions/msi/
'*          <MySQL Connector/ODBC>
'*            https://dev.mysql.com/downloads/connector/odbc/
'*          <SQLite ODBC Driver>
'*            http://www.ch-werner.de/sqliteodbc/
'*          <Microsoft Access データベース エンジン 2016 再頒布可能コンポーネント(ODBC/OLEDB両方)>
'*            https://www.microsoft.com/en-us/download/details.aspx?id=54920
'*
'*          OLEDBプロバイダ
'*          <Oracle OLEDB (Oracle Instant Client同梱)>
'*            https://www.oracle.com/database/technologies/instant-client/downloads.html
'*          <Microsoft OLE DB Driver for SQL Server>
'*            https://docs.microsoft.com/ja-jp/sql/connect/oledb/download-oledb-driver-for-sql-server
'*          <PgOleDB (PostgreSQL OLEDB Provider) 32bitのみ: ODBC使用を推奨>
'*            https://www.postgresql.org/ftp/projects/pgFoundry/stackbuilder/PgOleDB/PgOleDB1.0.0.20/
'*          <Microsoft Access データベース エンジン 2016 再頒布可能コンポーネント(ODBC/OLEDB両方)>
'*            https://www.microsoft.com/en-us/download/details.aspx?id=54920
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
'*-----------------------------------------------------------------------------
'* カーソル サービスの場所を表す。
'*
'*-----------------------------------------------------------------------------
Public Enum CursorLocationEnum
    adUseClient = 3
    adUseServer = 2
    adUseNone = 1
End Enum

'*-----------------------------------------------------------------------------
'* Recordset オブジェクトが使用するカーソルの種類。
'*
'*-----------------------------------------------------------------------------
Public Enum CursorTypeEnum
    adOpenDynamic = 2
    adOpenForwardOnly = 0
    adOpenKeyset = 1
    adOpenStatic = 3
    adOpenUnspecified = -1
End Enum
'*-----------------------------------------------------------------------------
'* Connection オブジェクトのトランザクション分離レベルを表す。
'*
'*-----------------------------------------------------------------------------
Public Enum IsolationLevelEnum
    adXactUnspecified = -1
    adXactChaos = 16
    adXactBrowse = 256
    adXactReadUncommitted = 256
    adXactCursorStability = 4096
    adXactReadCommitted = 4096
    adXactRepeatableRead = 65536
    adXactIsolated = 1048576
    adXactSerializable = 1048576
End Enum

'*-----------------------------------------------------------------------------
'* 編集時にレコードに適用されるロックの種類を表す。
'*
'*-----------------------------------------------------------------------------
Public Enum LockTypeEnum
    adLockBatchOptimistic = 4
    adLockOptimistic = 3
    adLockPessimistic = 2
    adLockReadOnly = 1
    adLockUnspecified = -1
End Enum

'******************************************************************************
'* 定数定義
'******************************************************************************
' Oracle ODBC接続(Oracle製Driver)
Private Const CONN_STR_ODBC_ORACLE_ORA As String = "Driver={{0}};DBQ=" & _
    "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={1})(PORT={2}))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME={3})))"
' Microsoft SQL Server ODBC接続
Private Const CONN_STR_ODBC_SQLSERVER As String = "Driver={SQL Server};Server={0};Database={1}"
' Microsoft SQL Server LocalDB ODBC接続
Private Const CONN_STR_ODBC_SQLSERVER_LOCALDB As String = "Driver={SQL Server Native Client 11.0};Server=(LocalDB)\{0};AttachDBFileName={1};Database={2};Trusted_Connection=Yes"
' PostgreSQL ODBC接続
Private Const CONN_STR_ODBC_POSTGRESQL As String = "Driver={{0}};Server={1};Port={2};Database={3}"
' MySQL ODBC接続
Private Const CONN_STR_ODBC_MYSQL As String = "Driver={{0}};Server={1};Port={2};Database={3}"
' Access/Excel/CSV/TXT ODBC接続
Private Const CONN_STR_ODBC_ACCESS As String = "Driver={{0}};DBQ={1}"
Private Const CONN_STR_ODBC_ACCESS_PASS As String = "Driver={{0}};DBQ={1};Uid=Admin;Pwd={2}"
' SQLite ODBC接続
Private Const CONN_STR_ODBC_SQLITE As String = "Driver={{0}};Database={1}"

' Oracle OLE接続(Oracle製Provider)
Private Const CONN_STR_OLE_ORACLE_ORA As String = "Provider= OraOLEDB.Oracle;Data Source=" & _
    "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={0})(PORT={1}))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME={2})))"
' Microsoft SQL Server OLE接続
Private Const CONN_STR_OLE_SQLSERVER As String = "Provider=SQLOLEDB;Data Source={0};Initial Catalog={1}"
' PostgreSQL OLE接続
Private Const CONN_STR_OLE_POSTGRESQL As String = "Provider=PostgreSQL OLE DB Provider;Data Source={1};Location={2}"
' Access OLE接続
Private Const CONN_STR_OLE_ACCESS As String = "Provider={0};Data Source={1}"
Private Const CONN_STR_OLE_ACCESS_PASS As String = "Provider={0};Data Source={1};User ID=Admin;Jet OLEDB:Database Password={2}"
' Excel/CSV/TXT/HTML OLE接続
Private Const CONN_STR_OLE_FILEDATA As String = "Provider={0};Data Source={1};Extended Properties=""{2}"""

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mConn As Object

'******************************************************************************
'* プロパティ定義
'******************************************************************************
Public Property Get IsReady() As Boolean
    IsReady = False
    If Not mConn Is Nothing Then
        If mConn.State = adStateOpen Then
            IsReady = True
        End If
    End If
End Property

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
'* [概  要] OpenOracle メソッド
'* [詳  細] Oracleデータベースへの接続を開く｡
'*          Oracle接続文字列を作成し、Connection.Open メソッドを実行する。
'*
'* @param serverName   接続先サーバ名
'* @param portNo       接続先ポート番号
'* @param serviceName  接続先サービス名（グローバルデータベース名）
'* @param userId       省略可能。接続を確立するときに使用するユーザー名。
'* @param password     省略可能。接続を確立するときに使用するパスワード。
'* @param connectType  省略可能。接続タイプ（方法）。「odbc」か「oledb」を指定。
'******************************************************************************
Public Sub OpenOracle(ByVal serverName As String, _
                          ByVal portNo As Integer, _
                          ByVal ServiceName As String, _
                          Optional ByVal userId As String, _
                          Optional ByVal Password As String, _
                          Optional ByVal connectType As String = "oledb")
    Dim cnnStr As String
    If LCase(connectType) = "odbc" Then
        cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_ORACLE_ORA, GetOracleODBCDriverName(), serverName, portNo, ServiceName)
    Else
        cnnStr = StringUtils.PlaceHolder(CONN_STR_OLE_ORACLE_ORA, serverName, portNo, ServiceName)
    End If
    Call mConn.Open(cnnStr, userId, Password)
End Sub

'******************************************************************************
'* [概  要] GetOracleODBCDriverName メソッド
'* [詳  細] Oracle用ODBCドライバー名を取得する｡
'*
'* @param use32bit 32ビット用ドライバー名を取得するかどうか指定
'*
'******************************************************************************
Private Function GetOracleODBCDriverName(Optional use32bit As Boolean = False) As String
    GetOracleODBCDriverName = GetODBCDriverName("Oracle in Ora", , use32bit)
End Function

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
'* @param connectType  省略可能。接続タイプ（方法）。「odbc」か「oledb」を指定。
'******************************************************************************
Public Sub OpenPostgreSql(ByVal serverName As String, _
                          ByVal portNo As Integer, _
                          ByVal databaseName As String, _
                          Optional ByVal userId As String, _
                          Optional ByVal Password As String, _
                          Optional ByVal connectType As String = "odbc")
    Dim cnnStr As String
    If LCase(connectType) = "odbc" Then
        cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_POSTGRESQL, GetPostgreSqlODBCDriverName(), serverName, portNo, databaseName)
    Else
        ' OLEDB接続（ポート指定不可）
        cnnStr = StringUtils.PlaceHolder(CONN_STR_OLE_POSTGRESQL, serverName, databaseName)
    End If
    Call mConn.Open(cnnStr, userId, Password)
End Sub

'******************************************************************************
'* [概  要] GetPostgreSqlODBCDriverName メソッド
'* [詳  細] PostgreSQL用ODBCドライバー名を取得する｡
'*
'* @param tryDefaultUse 取得できない場合にデフォルト値を返却するかどうか指定
'* @param use32bit 32ビット用ドライバー名を取得するかどうか指定
'*
'******************************************************************************
Private Function GetPostgreSqlODBCDriverName(Optional ByVal tryDefaultUse As Boolean = True, Optional use32bit As Boolean = False) As String
    If tryDefaultUse Then
        GetPostgreSqlODBCDriverName = GetODBCDriverName("PostgreSQL", "PostgreSQL Unicode", use32bit)
    Else
        GetPostgreSqlODBCDriverName = GetODBCDriverName("PostgreSQL", , use32bit)
    End If
End Function

'******************************************************************************
'* [概  要] OpenSqlServer メソッド
'* [詳  細] SqlServerデータベースへの接続を開く｡
'*          SqlServer接続文字列を作成し、Connection.Open メソッドを実行する。
'*
'* @param serverName   接続先サーバ名
'* @param databaseName 接続先データベース名
'* @param portNo       省略可能。接続先ポート番号
'* @param userId       省略可能。接続を確立するときに使用するユーザー名。
'* @param password     省略可能。接続を確立するときに使用するパスワード。
'* @param connectType  省略可能。接続タイプ（方法）。「odbc」か「oledb」を指定。
'******************************************************************************
Public Sub OpenSqlServer(ByVal serverName As String, _
                         ByVal databaseName As String, _
                         Optional ByVal portNo, _
                         Optional ByVal userId As String, _
                         Optional ByVal Password As String, _
                         Optional ByVal connectType As String = "odbc")
    Dim cnnStr As String
    If Not IsMissing(portNo) Then
        serverName = serverName & "," & portNo
    End If
    If LCase(connectType) = "odbc" Then
        cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_SQLSERVER, serverName, databaseName)
    Else
        cnnStr = StringUtils.PlaceHolder(CONN_STR_OLE_SQLSERVER, serverName, databaseName)
    End If
    Call mConn.Open(cnnStr, userId, Password)
End Sub

'******************************************************************************
'* [概  要] OpenSqlServerLocalDB メソッド
'* [詳  細] SqlServer（LocalDB）データベースへの接続を開く｡
'*          SqlServer（LocalDB）接続文字列を作成し、Connection.Open メソッド
'*          を実行する。
'*
'* @param dbFileName   データベースファイル名
'* @param databaseName 接続先データベース名
'* @param serverName  省略可能。サーバ名。省略時は「sqllocaldb info」コマンド
'*                    の実行結果からサーバ名を取得する。
'******************************************************************************
Public Sub OpenSqlServerLocalDB(ByVal dbFileName As String, _
                                ByVal databaseName As String, _
                                Optional ByVal serverName As String)
    Dim cnnStr As String
    If serverName = "" Then
        Dim cmdr As DosCommander: Set cmdr = New DosCommander
        serverName = StringUtils.RemoveEach(cmdr.GetCommandResult("sqllocaldb info"), Array(vbCrLf, vbCr, vbLf))
    End If
    cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_SQLSERVER_LOCALDB, serverName, dbFileName, databaseName)
    Call mConn.Open(cnnStr)
End Sub

'******************************************************************************
'* [概  要] OpenMySql メソッド
'* [詳  細] MySqlデータベースへの接続を開く｡
'*          MySql接続文字列を作成し、Connection.Open メソッドを実行する。
'*          ODBC接続のみ。
'*
'* @param serverName   接続先サーバ名
'* @param portNo       接続先ポート番号
'* @param databaseName 接続先データベース名
'* @param userId       省略可能。接続を確立するときに使用するユーザー名。
'* @param password     省略可能。接続を確立するときに使用するパスワード。
'******************************************************************************
Public Sub OpenMySql(ByVal serverName As String, _
                     ByVal portNo As Integer, _
                     ByVal databaseName As String, _
                     Optional ByVal userId As String, _
                     Optional ByVal Password As String)
    Dim cnnStr As String
    cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_MYSQL, GetODBCDriverName(Array("MySQL ODBC", "Unicode")), serverName, portNo, databaseName)
    Call mConn.Open(cnnStr, userId, Password)
End Sub

'******************************************************************************
'* [概  要] OpenSQLite メソッド
'* [詳  細] SQLiteデータベースへの接続を開く｡
'*          SQLite接続文字列を作成し、Connection.Open メソッド
'*          を実行する。
'*
'* @param databaseFileName   データベースファイル名
'*
'******************************************************************************
Public Sub OpenSQLite(ByVal databaseFileName As String)
    Dim cnnStr As String
    cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_SQLITE, GetODBCDriverName(Array("SQLite", "ODBC")), databaseFileName)
    Call mConn.Open(cnnStr)
End Sub

'******************************************************************************
'* [概  要] OpenAccess メソッド
'* [詳  細] Accessデータベースへの接続を開く｡
'*          Access接続文字列を作成し、Connection.Open メソッド
'*          を実行する。
'*
'* @param databaseFileName データベースファイル名
'* @param password     省略可能。接続を確立するときに使用するパスワード。
'* @param connectType  省略可能。接続タイプ（方法）。「odbc」か「oledb」を指定。
'*
'******************************************************************************
Public Sub OpenAccess(ByVal databaseFileName As String, _
                      Optional ByVal Password As String, _
                      Optional ByVal connectType As String = "oledb")
    Dim cnnStr As String
    If LCase(connectType) = "odbc" Then
        If Password = "" Then
            cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_ACCESS, GetAccessODBCDriverName(), databaseFileName)
        Else
            cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_ACCESS_PASS, GetAccessODBCDriverName(), databaseFileName, Password)
        End If
    Else
        If Password = "" Then
            cnnStr = StringUtils.PlaceHolder(CONN_STR_OLE_ACCESS, GetAccessOLEDBProviderName(), databaseFileName)
        Else
            cnnStr = StringUtils.PlaceHolder(CONN_STR_OLE_ACCESS_PASS, GetAccessOLEDBProviderName(), databaseFileName, Password)
        End If
    End If
    Call mConn.Open(cnnStr)
End Sub

'******************************************************************************
'* [概  要] OpenExcel メソッド
'* [詳  細] Excelファイルをデータベースとみなして接続を行う｡
'*          Excel接続文字列を作成し、Connection.Open メソッド
'*          を実行する。
'*
'* @param filePath      Excelファイルパス
'* @param headerExists  省略可能。ヘッダ。
'* @param readwriteMode 省略可能。0:書込モード、1:読込モード、2:双方向モード。
'* @param connectType   省略可能。接続タイプ（方法）。「odbc」か「oledb」を指定。
'*
'******************************************************************************
Public Sub OpenExcel(ByVal filePath As String, _
                     Optional ByVal HeaderExists As Boolean = True, _
                     Optional ByVal readwriteMode As Integer = 1, _
                     Optional ByVal connectType As String = "oledb")
    Dim cnnStr As String
    If LCase(connectType) = "odbc" Then
        cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_ACCESS, GetExcelODBCDriverName(), filePath)
    Else
        Dim exProp As New ArrayEx
        exProp.Add GetXlIsamType(filePath)
        If HeaderExists Then
            exProp.Add "HDR=Yes"
        Else
            exProp.Add "HDR=No"
        End If
        Select Case readwriteMode
            Case 0
                exProp.Add "IMEX=0"
            Case 1
                exProp.Add "IMEX=1"
            Case 2
                exProp.Add "IMEX=2"
        End Select
        cnnStr = StringUtils.PlaceHolder(CONN_STR_OLE_FILEDATA, GetAccessOLEDBProviderName(), filePath, exProp.Join(";"))
    End If
    Call mConn.Open(cnnStr)
End Sub

'******************************************************************************
'* [概  要] OpenCsv メソッド
'* [詳  細] Csvファイルをデータベースとみなして接続を行う｡
'*          接続文字列を作成し、Connection.Open メソッド
'*          を実行する。
'*
'* @param dirPath      CSVファイルの格納されているパス
'* @param headerExists 省略可能。ヘッダ。
'* @param connectType  省略可能。接続タイプ（方法）。「odbc」か「oledb」を指定。
'*
'******************************************************************************
Public Sub OpenCsv(ByVal dirPath As String, _
                   Optional ByVal HeaderExists As Boolean = True, _
                   Optional ByVal connectType As String = "oledb")
    Dim cnnStr As String
    If LCase(connectType) = "odbc" Then
        cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_ACCESS, GetTextODBCDriverName(), dirPath)
    Else
        Dim exProp As New ArrayEx
        exProp.Add "TEXT"
        If HeaderExists Then
            exProp.Add "HDR=Yes"
        Else
            exProp.Add "HDR=No"
        End If
        exProp.Add "FMT=Delimited"
        
        cnnStr = StringUtils.PlaceHolder(CONN_STR_OLE_FILEDATA, GetAccessOLEDBProviderName(), dirPath, exProp.Join(";"))
    End If
    Call mConn.Open(cnnStr)
End Sub

'******************************************************************************
'* [概  要] OpenText メソッド
'* [詳  細] Textファイルをデータベースとみなして接続を行う｡
'*          接続文字列を作成し、Connection.Open メソッド
'*          を実行する。
'*
'* @param dirPath      CSVファイルの格納されているパス
'* @param headerExists 省略可能。ヘッダ。
'* @param connectType  省略可能。接続タイプ（方法）。「odbc」か「oledb」を指定。
'*
'******************************************************************************
Public Sub OpenText(ByVal dirPath As String, _
                    Optional ByVal HeaderExists As Boolean = True, _
                    Optional ByVal connectType As String = "oledb")
    Dim cnnStr As String
    If LCase(connectType) = "odbc" Then
        cnnStr = StringUtils.PlaceHolder(CONN_STR_ODBC_ACCESS, GetTextODBCDriverName(), dirPath)
    Else
        Dim exProp As New ArrayEx
        exProp.Add "TEXT"
        If HeaderExists Then
            exProp.Add "HDR=Yes"
        Else
            exProp.Add "HDR=No"
        End If
        cnnStr = StringUtils.PlaceHolder(CONN_STR_OLE_FILEDATA, GetAccessOLEDBProviderName(), dirPath, exProp.Join(";"))
    End If
    Call mConn.Open(cnnStr)
End Sub

'******************************************************************************
'* [概  要] OpenHtml メソッド
'* [詳  細] Htmlファイルをデータベースとみなして接続を行う｡
'*          接続文字列を作成し、Connection.Open メソッドを実行する。
'*
'* @param filePath Htmlファイルパス
'*
'******************************************************************************
Public Sub OpenHtml(ByVal filePath As String)
    Dim cnnStr As String
    cnnStr = StringUtils.PlaceHolder(CONN_STR_OLE_FILEDATA, GetAccessOLEDBProviderName(), filePath, "HTML Import;HDR=Yes")
    Call mConn.Open(cnnStr)
End Sub

'******************************************************************************
'* [概  要] GetXlIsamType メソッド
'* [詳  細] Excelファイルの拡張子からExcelファイルのISAMタイプ名を取得する｡
'*
'* @param filePath Excelファイルパス
'* @return String ExcelファイルのISAMタイプ名
'******************************************************************************
Private Function GetXlIsamType(ByVal filePath As String) As String
    Select Case IO.GetExtensionName(filePath)
        Case "xls"
            GetXlIsamType = "Excel 8.0"
        Case "xlsb"
            GetXlIsamType = "Excel 12.0"
        Case "xlsm"
            GetXlIsamType = "Excel 12.0 Macro"
        Case "xlsx"
            GetXlIsamType = "Excel 12.0 Xml"
    End Select
End Function

'******************************************************************************
'* [概  要] GetAccessODBCDriverName メソッド
'* [詳  細] Access用ODBCドライバー名を取得する｡
'*
'* @param tryDefaultUse 取得できない場合にデフォルト値を返却するかどうか指定
'* @param use32bit 32ビット用ドライバー名を取得するかどうか指定
'* @return String Access用ODBCドライバー名
'******************************************************************************
Private Function GetAccessODBCDriverName(Optional ByVal tryDefaultUse As Boolean = True, Optional use32bit As Boolean = False) As String
    If tryDefaultUse Then
        GetAccessODBCDriverName = GetODBCDriverName(".accdb)", "Microsoft Access Driver (*.mdb, *.accdb)", use32bit)
    Else
        GetAccessODBCDriverName = GetODBCDriverName(".accdb)", , use32bit)
    End If
End Function

'******************************************************************************
'* [概  要] GetExcelODBCDriverName メソッド
'* [詳  細] Excelファイル接続用ODBCドライバー名を取得する｡
'*
'* @param tryDefaultUse 取得できない場合にデフォルト値を返却するかどうか指定
'* @param use32bit      32ビット用ドライバー名を取得するかどうか指定
'* @return String       Excelファイル接続用ODBCドライバー名
'******************************************************************************
Private Function GetExcelODBCDriverName(Optional ByVal tryDefaultUse As Boolean = True, Optional use32bit As Boolean = False) As String
    If tryDefaultUse Then
        GetExcelODBCDriverName = GetODBCDriverName(".xlsb)", "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)", use32bit)
    Else
        GetExcelODBCDriverName = GetODBCDriverName(".xlsb)", , use32bit)
    End If
End Function

'******************************************************************************
'* [概  要] GetTextODBCDriverName メソッド
'* [詳  細] Textファイル接続用ODBCドライバー名を取得する｡
'*
'* @param tryDefaultUse 取得できない場合にデフォルト値を返却するかどうか指定
'* @param use32bit      32ビット用ドライバー名を取得するかどうか指定
'* @return String       Textファイル接続用ODBCドライバー名
'******************************************************************************
Private Function GetTextODBCDriverName(Optional ByVal tryDefaultUse As Boolean = True, Optional use32bit As Boolean = False) As String
    If tryDefaultUse Then
        GetTextODBCDriverName = GetODBCDriverName("Text Driver (*.txt", "Microsoft Text Driver (*.txt; *.csv)", use32bit)
    Else
        GetTextODBCDriverName = GetODBCDriverName("Text Driver (*.txt", , use32bit)
    End If
End Function

'******************************************************************************
'* [概  要] GetAccessOLEDBProviderName メソッド
'* [詳  細] Access用OLEDBプロバイダー名を取得する｡
'*
'* @param tryDefaultUse 取得できない場合にデフォルト値を返却するかどうか指定
'* @param use32bit      32ビット用プロバイダー名を取得するかどうか指定
'* @return String       Access用OLEDBプロバイダー名
'******************************************************************************
Private Function GetAccessOLEDBProviderName(Optional ByVal tryDefaultUse As Boolean = True, Optional use32bit As Boolean = False) As String
    Dim v, vArr, vArr2, vProviderArr
    Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()

    stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\", vArr
    For Each v In ArrayUtils.RegexSearch(vArr, "\d\d\.\d")
        stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\" & v & "\ClickToRun\REGISTRY\MACHINE\Software\Classes", vArr2
        vProviderArr = ArrayUtils.Search(vArr2, "Microsoft.ACE.OLEDB.")
        If Not ArrayUtils.IsEmptyArray(vProviderArr) Then
            GetAccessOLEDBProviderName = vProviderArr(UBound(vProviderArr))
            Exit Function
        End If
    Next
    
    stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\", vArr2
    vProviderArr = ArrayUtils.Search(vArr2, "Microsoft.ACE.OLEDB.")
    If Not ArrayUtils.IsEmptyArray(vProviderArr) Then
        GetAccessOLEDBProviderName = vProviderArr(UBound(vProviderArr))
        Exit Function
    End If
    
    If tryDefaultUse Then
        GetAccessOLEDBProviderName = GetOleDbProviderName("Microsoft.ACE.OLEDB.", "Microsoft.ACE.OLEDB.12.0", use32bit)
    Else
        GetAccessOLEDBProviderName = GetOleDbProviderName("Microsoft.ACE.OLEDB.", , use32bit)
    End If
End Function

'******************************************************************************
'* [概  要] GetODBCDriverName メソッド
'* [詳  細] ODBCドライバー名をレジストリから取得する｡
'*
'* @param driverKeyword     ドライバー名を検索するキーワード（配列で複数指定可）
'* @param defaultDriverName 検索がHitしなかった場合のデフォルトドライバー名
'* @param use32bit          32ビット用ドライバー名を取得するかどうか指定
'* @return String           ODBCドライバー名
'******************************************************************************
Private Function GetODBCDriverName(ByVal driverKeyword, Optional ByVal defaultDriverName As String, Optional use32bit As Boolean = False) As String
    Dim vArr, vDriverArr
    Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
    If use32bit Then
        stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI", vArr
    Else
        stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBCINST.INI", vArr
    End If
    
    vDriverArr = ArrayUtils.Search(vArr, driverKeyword)
    If Not ArrayUtils.IsEmptyArray(vDriverArr) Then
        GetODBCDriverName = vDriverArr(UBound(vDriverArr))
    Else
        If defaultDriverName <> "" Then
            GetODBCDriverName = defaultDriverName
        Else
            Err.Raise 9999, "GetODBCDriverName", "ODBC Driver Not Found"
        End If
    End If
End Function

'******************************************************************************
'* [概  要] GetOleDbProviderName メソッド
'* [詳  細] OLEDBプロバイダー名をレジストリから取得する｡
'*
'* @param provideKeyword      プロバイダー名を検索するキーワード（配列で複数指定可）
'* @param defaultProviderName 検索がHitしなかった場合のデフォルトプロバイダー名
'* @param use32bit            32ビット用ドライバー名を取得するかどうか指定
'* @return String             OLEDBプロバイダー名
'******************************************************************************
Private Function GetOleDbProviderName(ByVal provideKeyword As String, Optional ByVal defaultProviderName As String, Optional use32bit As Boolean = False) As String
    Dim vArr, vProviderArr
    Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
    If use32bit Then
        stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\Classes", vArr
    Else
        stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Classes", vArr
    End If
    
    vProviderArr = ArrayUtils.Search(vArr, provideKeyword)
    If Not ArrayUtils.IsEmptyArray(vProviderArr) Then
        GetOleDbProviderName = vProviderArr(UBound(vProviderArr))
    Else
        If defaultProviderName <> "" Then
            GetOleDbProviderName = defaultProviderName
        Else
            Err.Raise 9999, "GetOleDbProviderName", "OLEDB Provider Not Found"
        End If
    End If
End Function

'******************************************************************************
'* [概  要] SelectList メソッド
'* [詳  細] SQL（SELECT）を実行し、実行結果をVariant配列で返却する｡
'*
'* @param sql             SQL（SELECT）
'* @param sqlParams       SQLパラメータ（Variant1次元配列）
'* @param pCursorLocation カーソルサービスの場所
'* @param pCursorType     カーソルタイプ
'* @param pLockType       ロックの種類
'* @return Variant        SQL実行結果
'*
'******************************************************************************
Public Function SelectList(ByVal sql As String, _
                           Optional ByVal sqlParams As Variant, _
                           Optional ByVal pCursorLocation As CursorLocationEnum = adUseClient, _
                           Optional ByVal pCursorType As CursorTypeEnum = adOpenForwardOnly, _
                           Optional ByVal pLockType As LockTypeEnum = adLockReadOnly) As Variant

    SelectList = SelectRecordset(sql, sqlParams, pCursorLocation, pCursorType, pLockType).GetRows
End Function

'******************************************************************************
'* [概  要] SelectRecordset メソッド
'* [詳  細] SQL（SELECT）を実行し、実行結果のRecordsetを返却する｡
'*
'* @param sql                   SQL（SELECT）
'* @param sqlParams             SQLパラメータ（Variant1次元配列）
'* @param pCursorLocation       カーソルサービスの場所
'* @param pCursorType           カーソルタイプ
'* @param pLockType             ロックの種類
'* @return Object(Of Recordset) SQL実行結果
'*
'******************************************************************************
Public Function SelectRecordset(ByVal sql As String, _
                                Optional ByVal sqlParams As Variant, _
                                Optional ByVal pCursorLocation As CursorLocationEnum = adUseClient, _
                                Optional ByVal pCursorType As CursorTypeEnum = adOpenForwardOnly, _
                                Optional ByVal pLockType As LockTypeEnum = adLockReadOnly) As Object
                                
    Dim adoRs As Object: Set adoRs = CreateObject("ADODB.Recordset")
    adoRs.CursorLocation = pCursorLocation
  
    ' TODO:SQLパラメータの設定
  
    adoRs.Open sql, mConn, pCursorType, pLockType
    Set SelectRecordset = adoRs
End Function

'******************************************************************************
'* [概  要] Disconnect メソッド
'* [詳  細] DBへのコネクションが開いている場合、クローズする｡
'*
'******************************************************************************
Public Sub Disconnect()
    If Not mConn Is Nothing Then
        If mConn.State <> adStateClosed Then
            Call mConn.Close
        End If
    End If
End Sub

'******************************************************************************
'* [概  要] Begin メソッド
'* [詳  細] トランザクションを開始する｡
'*
'******************************************************************************
Public Sub Begin()
     mConn.BeginTrans
End Sub

'******************************************************************************
'* [概  要] Commit メソッド
'* [詳  細] トランザクションをコミットする｡
'*
'******************************************************************************
Public Sub Commit()
    mConn.CommitTrans
End Sub

'******************************************************************************
'* [概  要] Rollback メソッド
'* [詳  細] トランザクションをロールバックする｡
'*
'******************************************************************************
Public Sub Rollback()
    mConn.RollbackTrans
End Sub

'******************************************************************************
'* [概  要] AccessCompactDB
'* [詳  細] Accessデータベースを最適化する｡
'*
'* @param accDbPath Accessデータベースファイルパス
'******************************************************************************
Public Sub AccessCompactDB(accDbPath As String)
    Dim newAccDbPath As String: newAccDbPath = IO.GetTempFilePath()
    Dim dbe As Object: Set dbe = CreateObject("DAO.DBEngine.120")
    dbe.CompactDatabase accDbPath, newAccDbPath
    IO.DeleteFile accDbPath
    IO.MoveFile newAccDbPath, accDbPath
End Sub


'******************************************************************************
'* [概  要] SQL表示処理。
'* [詳  細] SQLの埋込パラメータを「?」から置き換える。
'* [参  考] http://kagen88.blogspot.com/2012/12/adodbcommand.html
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


