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
'* [参  考]
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
    Call mConn.Open(cnnStr, UserID, Password, Options)
End Sub


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


