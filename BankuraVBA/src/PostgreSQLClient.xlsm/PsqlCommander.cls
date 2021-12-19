VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PsqlCommander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] PSQLコマンド実行クラス
'* [詳  細] PostgeSQLのコマンドラインツールPSQLのコマンド実行に関する処理
'*          を提供する。
'*
'* [参  考]　https://kenpg.bitbucket.io/blog/201507/24.html
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

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mCommander As Commander
Private mPsqlPath As String
Private mDbHost As String
Private mDbPort As String
Private mDbName As String
Private mDbUserName As String
Private mDbPassword As String
Private mClEncode As String
Private mTuplesOnly As Boolean

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* PsqlPath プロパティ
'*
'* psqlコマンドのフルパス
'*-----------------------------------------------------------------------------
Public Property Get PsqlPath() As String
    PsqlPath = mPsqlPath
End Property
Public Property Let PsqlPath(arg As String)
    mPsqlPath = StringUtils.WrapIfMissing(arg, """")
End Property

'*-----------------------------------------------------------------------------
'* DbHost プロパティ
'*
'* アクセスするDBのホスト（例：localhost）
'*-----------------------------------------------------------------------------
Public Property Get DbHost() As String
    DbHost = mDbHost
End Property
Public Property Let DbHost(arg As String)
    mDbHost = arg
End Property

'*-----------------------------------------------------------------------------
'* DbPort プロパティ
'*
'* アクセスするDBのポート番号
'*-----------------------------------------------------------------------------
Public Property Get DbPort() As String
    DbPort = mDbPort
End Property
Public Property Let DbPort(arg As String)
    mDbPort = arg
End Property

'*-----------------------------------------------------------------------------
'* DbName プロパティ
'*
'* アクセスするDB名
'*-----------------------------------------------------------------------------
Public Property Get dbName() As String
    dbName = mDbName
End Property
Public Property Let dbName(arg As String)
    mDbName = arg
End Property

'*-----------------------------------------------------------------------------
'* DbUserName プロパティ
'*
'* DBユーザ名
'*-----------------------------------------------------------------------------
Public Property Get DbUserName() As String
    DbUserName = mDbUserName
End Property
Public Property Let DbUserName(arg As String)
    mDbUserName = arg
End Property

'*-----------------------------------------------------------------------------
'* DbPassword プロパティ
'*
'* DBパスワード
'*-----------------------------------------------------------------------------
Public Property Get DbPassword() As String
    DbPassword = mDbPassword
End Property
Public Property Let DbPassword(arg As String)
    mDbPassword = arg
End Property

'*-----------------------------------------------------------------------------
'* ClEncode プロパティ
'*
'* Clientエンコード
'*-----------------------------------------------------------------------------
Public Property Get ClEncode() As String
    ClEncode = mClEncode
End Property
Public Property Let ClEncode(arg As String)
    mClEncode = arg
End Property

'*-----------------------------------------------------------------------------
'* TuplesOnly プロパティ
'*
'* 出力結果をデータのみにするかどうか。Falseにすると、ヘッダと行数が出力される。
'*-----------------------------------------------------------------------------
Public Property Get TuplesOnly() As Boolean
    TuplesOnly = mTuplesOnly
End Property
Public Property Let TuplesOnly(arg As Boolean)
    mTuplesOnly = arg
End Property
'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mCommander = New Commander
    mClEncode = "SJIS"
    mTuplesOnly = False
    Call SetPsqlInstallPathFromRegistry
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] SetPsqlBaseCommand
'* [詳  細] Psqlのコマンド・環境変数を設定する。
'*          事前に各プロパティを設定しておく必要がある。
'*
'******************************************************************************
Private Sub SetPsqlBaseCommand()
    mCommander.BaseCommand = mPsqlPath & " -h " & mDbHost & " -p " & mDbPort & _
                                         " -d " & mDbName & " -U " & mDbUserName & _
                                         " -c ""{0}"" -A " & "-o {1}"
    If mTuplesOnly Then
        mCommander.BaseCommand = mCommander.BaseCommand & " -t "
    End If
   
    ' パスワードを環境変数に設定
    Call Base.AppendEnvItem("PGPASSWORD", mDbPassword)
    ' Client(psql)のエンコード設定：デフォルト「SJIS」
    Call Base.AppendEnvItem("PGCLIENTENCODING", mClEncode)
End Sub

'******************************************************************************
'* [概  要] psqlコマンドを介してSQLを実行する。
'* [詳  細] psqlコマンドを介して、SQLを実行し、実行結果を返却する。
'*          ローカル環境にpsqlがインストールされている必要がある。
'*          psqlはpostgreSQLインストール時に合わせてインストールされる。
'*          実行結果は一時ファイルに一度出力後、読み出しを行うため、
'*          大量データにも対応可能（現状一括読出しのため、GiBレベルだと
'*          対応は難しいかもしれない）。
'*
'* @param sql SQL文
'* @return 実行結果データ（2次元配列）
'*
'******************************************************************************
Public Function Exec(sql As String) As Variant
    Call SetPsqlBaseCommand
    Dim resultData: resultData = mCommander.GetCommandResultAsArray(sql)
    Exec = EditQueryResult(resultData)
End Function

'******************************************************************************
'* [概  要] psqlコマンドを介してSQLを実行する。
'* [詳  細] psqlコマンドを介して、SQLを実行し、実行結果を返却する。
'*          ローカル環境にpsqlがインストールされている必要がある。
'*          psqlはpostgreSQLインストール時に合わせてインストールされる。
'*
'* [注  意] WScript.ShellのExecメソッドの標準出力及びエラー出力の
'*          バッファサイズが4KBのため、大量にデータが出力される場合バッファが
'*          一杯になり処理がフリーズする。
'*          そのため、Exec終了を待たずに、標準出力から行毎の読み出しを行うことで
'*          バッファがいっぱいにならないようにしているが、
'*          標準出力の1行分のデータ量が4KBを超える場合は対応できずフリーズする。
'*
'* @param sql SQL文
'* @return CommandResult情報
'*
'******************************************************************************
Public Function ExecToCommandResult(sql As String) As CommandResult
    Dim cmd As String, oExec
    
    ' コマンド組立て
    cmd = mPsqlPath & " -h " & mDbHost & " -p " & mDbPort & " -d " & mDbName & " -U " & mDbUserName & " -c """ & sql & """ -A"
    If mTuplesOnly Then
        cmd = cmd & " -t"
    End If

    ' パスワードを環境変数に設定
    Call Base.AppendEnvItem("PGPASSWORD", mDbPassword)
    ' Client(psql)のエンコード設定：デフォルト「SJIS」
    Call Base.AppendEnvItem("PGCLIENTENCODING", mClEncode)

    ' コマンド実行・結果設定
    Dim res As CommandResult: Set res = mCommander.Exec(cmd)
    
    ' コマンド失敗時処理
    If res.ExitCd <> 0 Then
        Set ExecToCommandResult = res
        Exit Function
    End If

    If ArrayUtils.IsEmptyArray(res.StdOutList) Then
        Set ExecToCommandResult = res
        Exit Function
    End If
    
    ' 実行結果編集
    Dim resultEditData: resultEditData = EditQueryResult(res.StdOutList)

    ' 返却値設定
    res.StdOutList = resultEditData
    res.RowMax = ArrayUtils.GetLength(resultEditData)
    res.ColMax = ArrayUtils.GetLength(resultEditData, 2)
    Set ExecToCommandResult = res
End Function

'******************************************************************************
'* [概  要] SQL実行結果編集処理。
'* [詳  細] psqlコマンドを介したSQL実行結果を2次元配列に編集する。
'*
'* @param resultData 実行結果データ（1次元配列）
'* @return 編集済み実行結果データ（2次元配列）
'*
'******************************************************************************
Private Function EditQueryResult(resultData)
    ' 正常時処理
    Dim lRowMax As Long: lRowMax = 0
    Dim lColMax As Long: lColMax = 0
    
    If ArrayUtils.IsEmptyArray(resultData) Then
        EditQueryResult = ArrayUtils.EmptyVariant2DArray
        Exit Function
    End If
    lRowMax = ArrayUtils.GetLength(resultData)
    
    ' 列数計測
    Dim vTmpCols As Variant: vTmpCols = Split(resultData(0), "|")
    lColMax = UBound(vTmpCols) - LBound(vTmpCols) + 1

    'データ設定
    Dim vArray()
    ReDim vArray(0 To lRowMax - 1, 0 To lColMax - 1)
    
    Dim i As Long, cols As Variant
    Dim row
    For Each row In resultData
        'Debug.Print "Line: " & row
        
        ' SQL実行結果に列区切り(|)なしのパターン
        If StringUtils.ContainsNone(row, "|") Or row = "" Then
            ' 1列目以外は空埋め
            Call ArrayUtils.SetDefaultTo2DArrayRow(vArray, i, "")
            vArray(i, 0) = row
        ' 列区切り(|)あり
        Else
            cols = Split(row, "|")
            Call ArrayUtils.Copy1DArrayTo2DArrayRow(vArray, i, cols)
        End If
        i = i + 1
    Next
    EditQueryResult = ArrayUtils.RemoveEmpty(vArray)
End Function

'******************************************************************************
'* [概  要] SetPsqlInstallPathFromRegistry
'* [詳  細] レジストリからPsqlのインストールパスを取得し、設定する。
'*
'******************************************************************************
Private Sub SetPsqlInstallPathFromRegistry()
    Const PSQL_REG_PATH As String = "SOFTWARE\PostgreSQL\Installations"
    Const PSQL_REG_WOW_PATH As String = "SOFTWARE\Wow6432Node\PostgreSQL\Installations"
    
    Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
    Dim regPsqlInstallPathString As String
    Dim regRoot As Long

    Dim vArr
    stdRegProv.EnumKey HKEY_CURRENT_USER, PSQL_REG_PATH, vArr
    If IsNull(vArr) Then
        stdRegProv.EnumKey HKEY_LOCAL_MACHINE, PSQL_REG_PATH, vArr
        If IsNull(vArr) Then
            stdRegProv.EnumKey HKEY_LOCAL_MACHINE, PSQL_REG_WOW_PATH, vArr
            regPsqlInstallPathString = PSQL_REG_WOW_PATH
            regRoot = HKEY_LOCAL_MACHINE
        Else
            regPsqlInstallPathString = PSQL_REG_PATH
            regRoot = HKEY_LOCAL_MACHINE
        End If
    Else
        regPsqlInstallPathString = PSQL_REG_PATH
        regRoot = HKEY_CURRENT_USER
    End If
    
    If Not IsNull(vArr) Then
        ArrayUtils.Sort vArr, False
        stdRegProv.GetStringValue regRoot, regPsqlInstallPathString & "\" & CStr(vArr(0)), "Base Directory", mPsqlPath
        mPsqlPath = StringUtils.wrap(mPsqlPath & "\bin\psql.exe", """")
    End If
End Sub
