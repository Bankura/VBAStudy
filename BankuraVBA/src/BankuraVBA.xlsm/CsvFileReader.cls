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

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CSVFile読込クラス
'* [詳  細] CSVFile読込用のユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* Enum定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オブジェクトが開いているか閉じているか、データ ソースに接続中か、コマンドを
'* 実行中か、またはデータを取得中かを表します。
'*-----------------------------------------------------------------------------
Private Enum ObjectStateEnum
    adStateClosed = 0     'オブジェクトが閉じていることを示します。
    adStateOpen = 1       'オブジェクトが開いていることを示します。
    adStateConnecting = 2 'オブジェクトが接続中であることを示します。
    adStateExecuting = 4  'オブジェクトがコマンドを実行中であることを示します。
    adStateFetching = 8   'オブジェクトの行を取得中であることを示します。
End Enum

'*-----------------------------------------------------------------------------
'* Stream オブジェクトに保存するデータの種類を表します。
'*
'*-----------------------------------------------------------------------------
Private Enum StreamTypeEnum
    adTypeBinary = 1 'バイナリ データを示します。
    adTypeText = 2   '既定値。Charset で指定された文字セットのテキスト データを示します。
End Enum

'*-----------------------------------------------------------------------------
'* Stream オブジェクトから、ストリーム全体を読み取るか、または次の行を読み取るかを表します。
'*
'*-----------------------------------------------------------------------------
Private Enum StreamReadEnum
    adReadAll = -1  '既定値。現在の位置から EOS マーカー方向に、すべてのバイトをストリームから読み取ります。これは、バイナリ ストリーム (Type は adTypeBinary) に唯一有効な StreamReadEnum 値です。
    adReadLine = -2 'ストリームから次の行を読み取ります (LineSeparator プロパティで指定)。
End Enum

'*-----------------------------------------------------------------------------
'* Connection 内のデータの編集、Record のオープン、または Record および Stream
'*  オブジェクトの Mode プロパティの値の指定に対する権限を表します。
'*-----------------------------------------------------------------------------
Private Enum ConnectModeEnum
    adModeRead = 1
    adModeReadWrite = 3
    adModeRecursive = 4194304
    adModeShareDenyNone = 16
    adModeShareDenyRead = 4
    adModeShareDenyWrite = 8
    adModeShareExclusive = 12
    adModeUnknown = 0
    adModeWrite = 2
End Enum

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mFso As Object
Private mFile As Object
Private mAdoStream As Object
Private mCharSet As String
Private mLineSeparator As LineSeparatorsEnum
Private mSize As Long
Private mHeaderExists As Boolean
Private mFilePath As String
Private mValidFormat As Boolean

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
        Set mFile = mFso.GetFile(mFilePath)
        mSize = mFile.size
    End If
End Property

'*-----------------------------------------------------------------------------
'* Charset プロパティ
'*-----------------------------------------------------------------------------
Property Get charSet() As String
    charSet = mCharSet
End Property
Property Let charSet(arg As String)
    mCharSet = arg
End Property

'*-----------------------------------------------------------------------------
'* LineSeparator プロパティ
'*-----------------------------------------------------------------------------
Property Get lineSeparator() As LineSeparatorsEnum
    lineSeparator = mLineSeparator
End Property
Property Let lineSeparator(arg As LineSeparatorsEnum)
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
Property Get headerExists() As Boolean
    headerExists = mHeaderExists
End Property
Property Let headerExists(arg As Boolean)
    mHeaderExists = arg
End Property

'*-----------------------------------------------------------------------------
'* ValidFormat プロパティ（読み取り専用）
'*-----------------------------------------------------------------------------
Property Get ValidFormat() As Boolean
    ValidFormat = mValidFormat
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mFso = CreateObject("Scripting.FileSystemObject")
    Set mAdoStream = CreateObject("ADODB.Stream")
    mCharSet = "UTF-8"
    mLineSeparator = adLF
    mHeaderExists = False
    mValidFormat = True
End Sub
Private Sub Class_Terminate()
    If Not mAdoStream Is Nothing Then
        If mAdoStream.State <> adStateClosed Then
            Call mAdoStream.Close
        End If
    End If
    Set mAdoStream = Nothing
    Set mFso = Nothing
    Set mFile = Nothing
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
    FileExists = mFso.FileExists(tempPath)
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
    ShowCsvFileDialog = ""
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .InitialFileName = initialFile
        .AllowMultiSelect = False
        If .Show = True Then
            mFilePath = .SelectedItems(1)
            Set mFile = mFso.GetFile(mFilePath)
            mSize = mFile.size
            ShowCsvFileDialog = mFilePath
        End If
    End With
End Function


'******************************************************************************
'* [概  要] テキストファイル読込処理。
'* [詳  細] テキストファイルを行単位で読込み、Variant配列に設定して返却する。
'*
'* @return 読込テキストデータ（Variant配列）
'******************************************************************************
Public Function ReadTextFileToVArray() As Variant

    Dim lSize As Long, lCount As Long, i As Long, vRow, vArr()
    With mAdoStream
        .mode = adModeReadWrite
        .Type = adTypeText
        .charSet = mCharSet
        .lineSeparator = mLineSeparator
        .Open
        .LoadFromFile mFilePath
        lSize = Me.CountLine
        
        '空ファイルは空Variant配列で返却
        If (mHeaderExists And lSize <= 1) Or (Not mHeaderExists And lSize = 0) Then
            .Close
            Exit Function
        End If
        '1行目処理
        .Position = 0
        vRow = SplitCsvRow(.ReadText(adReadLine), True)
         If Not mHeaderExists Then
            ReDim vArr(0 To lSize - 1, 0 To UBound(vRow))
            For i = LBound(vRow) To UBound(vRow)
                vArr(lCount, i) = vRow(i)
            Next
            lCount = lCount + 1
        Else
            ReDim vArr(0 To lSize - 2, 0 To UBound(vRow))
        End If
        Dim ColumnCount As Long: ColumnCount = UBound(vRow)
        
        Dim statusBarTxt As String: statusBarTxt = "CSV読込処理中："
        Application.StatusBar = statusBarTxt
         
        '2行目以降処理
        Do While Not .EOS
            vRow = SplitCsvRow(.ReadText(adReadLine), True)
            If UBound(vRow) <> ColumnCount Then
                .Close
                ReadTextFileToVArray = Empty
                mValidFormat = False
                Exit Function
            End If

            If lCount Mod 5000 = 0 Then
                statusBarTxt = statusBarTxt & "■"
                Application.StatusBar = statusBarTxt
                Call UXUtils.CheckEvents
            End If
            
            For i = LBound(vRow) To UBound(vRow)
                vArr(lCount, i) = vRow(i)
            Next
            lCount = lCount + 1
        Loop
        .Close
    End With

    ReadTextFileToVArray = vArr
End Function

'******************************************************************************
'* [概  要] CSV行分割処理。
'* [詳  細] CSV文字列をカンマごとに分割してVariant配列に格納し、返却する。
'*          項目両端の「"」を削除する。
'*
'* @param line CSV行
'* @param quotExist ダブルクォート存在有無（True：あり）。省略可能。
'* @param delimiter デリミタ。省略可能。
'* @return 分割後Variant配列
'******************************************************************************
Public Function SplitCsvRow(line As String, Optional ByVal quotExist As Boolean = False, Optional delimiter As String = ",") As Variant
    If quotExist Then
        Dim lQuoteCount As Long, lItemCount As Long, i As Long, tmpTxt As String, v()
        For i = 1 To Len(line)
            Select Case Mid(line, i, 1)
                Case delimiter
                    If lQuoteCount Mod 2 = 0 Then
                        lItemCount = lItemCount + 1
                    End If
                Case """"
                    lQuoteCount = lQuoteCount + 1
            End Select
        Next
        ReDim v(lItemCount) As Variant
        lQuoteCount = 0: lItemCount = 0
        For i = 1 To Len(line)
            Select Case Mid(line, i, 1)
                Case delimiter
                    If lQuoteCount Mod 2 = 0 Then
                        v(lItemCount) = EditIncludeQuotItem(tmpTxt)
                        tmpTxt = ""
                        lItemCount = lItemCount + 1
                    Else
                        tmpTxt = tmpTxt & Mid(line, i, 1)
                    End If
                Case """"
                    lQuoteCount = lQuoteCount + 1
                    tmpTxt = tmpTxt & Mid(line, i, 1)
                Case Else
                    tmpTxt = tmpTxt & Mid(line, i, 1)
            End Select
        Next
        If tmpTxt <> "" Then
            v(lItemCount) = EditIncludeQuotItem(tmpTxt)
        End If
        SplitCsvRow = v
    Else
        SplitCsvRow = Split(line, delimiter)
    End If
End Function

'******************************************************************************
'* [概  要] ダブルクォート入り項目編集処理。
'* [詳  細] エスケープされているダブルクォート「""」を「"」に変換し、
'*          項目両端の「"」を削除する。
'*
'* @param text ダブルクォート入り項目
'* @return 編集後項目
'******************************************************************************
Private Function EditIncludeQuotItem(ByVal text As String) As String
    text = Replace(text, """""", """")
    If Left(text, 1) = """" And Right(text, 1) = """" Then
        If Len(text) <= 2 Then
            EditIncludeQuotItem = ""
        Else
            EditIncludeQuotItem = Mid(text, 2, Len(text) - 2)
            
        End If
    End If
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
    With mAdoStream
        .Position = 0
        Do While Not .EOS
            .SkipLine
            lineCount = lineCount + 1
        Loop
    End With

    CountLine = lineCount
End Function

