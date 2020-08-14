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
    mPsqlPath = arg
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
Public Property Get DbName() As String
    DbName = mDbName
End Property
Public Property Let DbName(arg As String)
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


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mCommander = New Commander
    mClEncode = "SJIS"
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] psqlコマンドを介してSQLを実行する。
'* [詳  細] psqlコマンドを介して、SQLを実行し、実行結果を返却する。
'*          ローカル環境にpsqlがインストールされている必要がある。
'*          psqlはpostgreSQLインストール時に合わせてインストールされる。
'*
'* @param sql SQL文
'* @return CommandResult情報
'*
'******************************************************************************
Public Function ExecPsql(sql As String) As CommandResult
    Dim cmd As String, oExec
    
    ' コマンド組立て
    cmd = mPsqlPath & " -h " & mDbHost & " -p " & mDbPort & " -d " & mDbName & " -U " & mDbUserName & " -c """ & sql & """ -A"
    Debug.Print cmd
    
    ' Wscript.Shellオブジェクト生成
    With Core.Wsh
        ' パスワードを環境変数に設定
        .Environment("Process").Item("PGPASSWORD") = mDbPassword
    
        ' Client(psql)のエンコード設定：デフォルト「SJIS」
        .Environment("Process").Item("PGCLIENTENCODING") = mClEncode
    End With

    ' コマンド実行・結果設定
    Dim res As CommandResult: Set res = mCommander.Exec(cmd)
    
    ' コマンド失敗時処理
    If res.ExitCd <> 0 Then
        Set ExecPsql = res
        Exit Function
    End If

    ' 正常時処理
    Dim lRowMax As Long: lRowMax = 0
    Dim lColMax As Long: lColMax = 0
    
    If ArrayUtils.IsEmptyArray(res.StdOutList) Then
        Set ExecPsql = res
        Exit Function
    End If
    ' 列数計測
    Dim vTmpCols As Variant: vTmpCols = Split(res.StdOutList(0), "|")
    lColMax = UBound(vTmpCols) - LBound(vTmpCols) + 1

    'データ設定
    Dim vArray()
    ReDim vArray(0 To lRowMax - 1, 0 To lColMax - 1)
    
    Dim i As Long, cols As Variant
    Dim row
    For Each row In res.StdOutList
        Debug.Print "Line: " & row
        
        ' SQL実行結果に列区切り(|)なしのパターン
        If StringUtils.ContainsNone(row, "|") Then
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

    res.StdOutList = vArray
    res.RowMax = lRowMax
    res.ColMax = lColMax
    Set ExecPsql = res
End Function
