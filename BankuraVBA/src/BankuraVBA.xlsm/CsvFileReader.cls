VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IRowEditor

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CSVFile読込クラス
'* [詳  細] CSVFile読込用のユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mAdoStream As Object
Private mCharSet As String
Private mLineSeparator As LineSeparatorsEnum
Private mSize As Long
Private mHeaderExists As Boolean
Private mQuotExists As Boolean
Private mFilePath As String
Private mValidFormat As Boolean
Private mRowCount As Long
Private mColumnCount As Long
Private mProgressReporter As IProgressReporter
Private mReportUnit As Long

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* FilePath プロパティ
'*-----------------------------------------------------------------------------
Property Get filePath() As String
    filePath = mFilePath
End Property
Property Let filePath(arg As String)
    mFilePath = arg
    If FileExists(arg) Then
        mSize = IO.GetFile(mFilePath).size
    End If
End Property

'*-----------------------------------------------------------------------------
'* Charset プロパティ
'*-----------------------------------------------------------------------------
Property Get CharSet() As String
    CharSet = mCharSet
End Property
Property Let CharSet(arg As String)
    mCharSet = arg
End Property

'*-----------------------------------------------------------------------------
'* LineSeparator プロパティ
'*-----------------------------------------------------------------------------
Property Get LineSeparator() As LineSeparatorsEnum
    LineSeparator = mLineSeparator
End Property
Property Let LineSeparator(arg As LineSeparatorsEnum)
    mLineSeparator = arg
End Property

'*-----------------------------------------------------------------------------
'* Size プロパティ（読み取り専用）
'*-----------------------------------------------------------------------------
Property Get size() As Long
    size = mSize
End Property

'*-----------------------------------------------------------------------------
'* HeaderExists プロパティ
'*-----------------------------------------------------------------------------
Property Get HeaderExists() As Boolean
    HeaderExists = mHeaderExists
End Property
Property Let HeaderExists(arg As Boolean)
    mHeaderExists = arg
End Property

'*-----------------------------------------------------------------------------
'* QuotExists プロパティ
'*-----------------------------------------------------------------------------
Property Get QuotExists() As Boolean
    QuotExists = mQuotExists
End Property
Property Let QuotExists(arg As Boolean)
    mQuotExists = arg
End Property

'*-----------------------------------------------------------------------------
'* ValidFormat プロパティ（読み取り専用）
'*-----------------------------------------------------------------------------
Property Get ValidFormat() As Boolean
    ValidFormat = mValidFormat
End Property

'*-----------------------------------------------------------------------------
'* ProgressReporter プロパティ
'*-----------------------------------------------------------------------------
Property Get ProgressReporter() As IProgressReporter
    Set ProgressReporter = mProgressReporter
End Property
Property Set ProgressReporter(arg As IProgressReporter)
    Set mProgressReporter = arg
End Property

'*-----------------------------------------------------------------------------
'* ReportUnit  プロパティ
'*-----------------------------------------------------------------------------
Property Get ReportUnit() As Long
    ReportUnit = mReportUnit
End Property
Property Let ReportUnit(arg As Long)
    mReportUnit = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mAdoStream = CreateObject("ADODB.Stream")
    mCharSet = "UTF-8"
    mLineSeparator = adLf
    mHeaderExists = False
    mQuotExists = False
    mValidFormat = True
    mRowCount = 0
    mColumnCount = 0
    mReportUnit = 500
End Sub
Private Sub Class_Terminate()
    If Not mAdoStream Is Nothing Then
        If mAdoStream.State <> adStateClosed Then
            Call mAdoStream.Close
        End If
    End If
    Set mAdoStream = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] FileExists メソッド
'* [詳  細] ファイルが存在するかどうか判定します｡
'*
'* @param targetFilePath ファイルパス。省略可能。
'* @return 判定結果
'*
'******************************************************************************
Public Function FileExists(Optional ByVal targetFilePath As String) As Boolean
    Dim tempPath As String: tempPath = mFilePath
    If targetFilePath <> "" Then
        tempPath = targetFilePath
    End If
    FileExists = FileUtils.FileExists(tempPath)
End Function

'******************************************************************************
'* [概  要] IsEmpty メソッド
'* [詳  細] ファイルが空ファイルかどうか判定します｡
'*
'* @return 判定結果
'*
'******************************************************************************
Public Function IsEmpty() As Boolean
    IsEmpty = (mSize = 0)
End Function

'******************************************************************************
'* [概  要] ShowCsvFileDialog メソッド
'* [詳  細] CSVファイルを選択するダイアログを表示します｡
'*
'* @param initialFile 初期表示用ファイルパス。省略可能。
'* @return 選択したファイルのパス
'*
'******************************************************************************
Public Function ShowCsvFileDialog(Optional ByVal initialFile As String) As String
    mFilePath = UXUtils.ShowCsvFileDialog(initialFile)
    If mFilePath <> "" Then
        mSize = IO.GetFile(mFilePath).size
    Else
        mSize = 0
    End If
    ShowCsvFileDialog = mFilePath
End Function

'******************************************************************************
'* [概  要] CSVファイル読込処理。
'* [詳  細] CSVファイルを行単位で読込み、Array2DExオブジェクトに設定して返却する。
'*
'* @return Array2DEx 読込テキストデータ
'******************************************************************************
Public Function Read() As Array2DEx
    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    Dim lSize As Long
    With mAdoStream
        .mode = adModeReadWrite
        .Type = adTypeText
        .CharSet = mCharSet
        .LineSeparator = mLineSeparator
        .Open
        .LoadFromFile mFilePath
        lSize = Me.CountLine
        
        ' 空ファイルは空Variant配列で返却
        If (mHeaderExists And lSize <= 1) Or (Not mHeaderExists And lSize = 0) Then
            .Close
            Set Read = arr2dex
            Exit Function
        End If
        .Position = 0
    End With

    ' RowEnumeratorの設定
    Dim rEnumerator As RowEnumerator: Set rEnumerator = New RowEnumerator
    Set rEnumerator.RowEditor = Me
    rEnumerator.SourceData = mAdoStream
    rEnumerator.DestTarget = arr2dex
    rEnumerator.IsActionable = False
    rEnumerator.MaxReportCount = lSize
    rEnumerator.endRow = lSize
    mRowCount = lSize
    
    ' ProgressControllerの設定
    Dim pController As ProgressController: Set pController = New ProgressController
    If mProgressReporter Is Nothing Then
        pController.IsReportable = False
    Else
        Set pController.ProgressReporter = mProgressReporter
        pController.IsReportable = True
        rEnumerator.ReportUnit = mReportUnit
    End If

    ' 処理開始
    On Error GoTo ErrorHandler
    pController.Start rEnumerator

    Set Read = arr2dex
    Exit Function
ErrorHandler:
    If Err.Number = 9998 Then
        mValidFormat = False
        Set Read = arr2dex
    ElseIf Err.Number = 9997 Then
        Set Read = arr2dex
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

'******************************************************************************
'* [概  要] CSVファイル読込処理。
'* [詳  細] CSVファイルを行単位で読込み、Variant配列に設定して返却する。
'*
'* @return 読込テキストデータ（Variant配列）
'******************************************************************************
Public Function ReadToVArray() As Variant
    Dim arr2dex As Array2DEx
    Set arr2dex = Read()
    ReadToVArray = arr2dex.To2DArray
End Function

'******************************************************************************
'* [概  要] CSV行分割処理。
'* [詳  細] CSV文字列をカンマごとに分割してVariant配列に格納し、返却する。
'*
'* @param line CSV行
'* @param quotExist ダブルクォート存在有無（True：あり）。省略可能。
'* @param delimiter デリミタ。省略可能。
'* @return 分割後Variant配列
'******************************************************************************
Public Function SplitCsvRow(line As String, Optional ByVal QuotExists As Boolean = False, Optional Delimiter As String = ",") As Variant
    SplitCsvRow = StringUtils.SplitCsvRow(line, QuotExists, Delimiter)
End Function

'******************************************************************************
'* [概  要] テキストの行数取得処理。
'* [詳  細] テキストの行数を取得する。
'* [備　考] 呼び出す前にStreamがOpenされていること。
'*
'* @return 行数
'******************************************************************************
Public Function CountLine() As Long
    CountLine = FileUtils.CountLine(mAdoStream)
End Function

'******************************************************************************
'* [概  要] StartProcess
'* [詳  細] 開始時の処理を行う。
'*
'******************************************************************************
Private Sub StartProcess(enumerator As Object)
    Dim rEnumerator As RowEnumerator, arr2dex As Array2DEx
    Set rEnumerator = enumerator

    Set arr2dex = rEnumerator.DestTarget
    Dim vRow
    With rEnumerator.SourceData
        ' 1行目処理
        vRow = SplitCsvRow(.ReadText(adReadLine), True)
        If Not mHeaderExists Then
            arr2dex.AddRow vRow
        End If
        mColumnCount = UBound(vRow)
    End With
    rEnumerator.startRow = 2
End Sub

'******************************************************************************
'* [概  要] Edit
'* [詳  細] 行編集処理を行う。
'*
'******************************************************************************
Private Sub Edit(enumerator As Object)
    Dim rEnumerator As RowEnumerator, arr2dex As Array2DEx
    Set rEnumerator = enumerator
    Set arr2dex = rEnumerator.DestTarget

    Dim vRow
    With rEnumerator.SourceData
        vRow = SplitCsvRow(.ReadText(adReadLine), mQuotExists)
        If UBound(vRow) <> mColumnCount Then
            .Close
            Err.Raise 9998, "CsvFileReader#Edit", "行毎の項目数が不整合です。"
        End If
        arr2dex.AddRow vRow
    End With
End Sub

'******************************************************************************
'* [概  要] EndProcess
'* [詳  細] 終了時の処理を行う。
'*
'******************************************************************************
Private Sub EndProcess(enumerator As Object)
    enumerator.SourceData.Close
End Sub

'******************************************************************************
'* [概  要] CancelProcess
'* [詳  細] キャンセル時の処理を行う。
'*
'******************************************************************************
Private Sub CancelProcess(enumerator As Object)
    enumerator.SourceData.Close
End Sub

'******************************************************************************
'* インタフェース実装
'******************************************************************************
'*-----------------------------------------------------------------------------
'* IRowEditor#StartProcess
'*
'*-----------------------------------------------------------------------------
Private Sub IRowEditor_StartProcess(enumerator As Object)
    StartProcess enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IRowEditor#Edit
'*
'*-----------------------------------------------------------------------------
Private Sub IRowEditor_Edit(enumerator As Object)
    Edit enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IRowEditor#EndProcess
'*
'*-----------------------------------------------------------------------------
Private Sub IRowEditor_EndProcess(enumerator As Object)
    EndProcess enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IRowEditor#CancelProcess
'*
'*-----------------------------------------------------------------------------
Public Sub IRowEditor_CancelProcess(enumerator As Object)
    CancelProcess enumerator
End Sub

