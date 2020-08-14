VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] File書込クラス
'* [詳  細] File書込用のユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
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
'* Stream オブジェクトからファイルに保存するときに、ファイルを作成するか、
'* 上書きするかを表します。これらの値は AND 演算子で結合できます。
'*-----------------------------------------------------------------------------
Private Enum SaveOptionsEnum
    adSaveCreateNotExist = 1  '既定値。FileName パラメーターで指定したファイルがない場合は新しいファイルが作成されます。
    adSaveCreateOverWrite = 2 'Filename パラメーターで指定したファイルがある場合は、現在開かれている Stream オブジェクトのデータでファイルが上書きされます。
End Enum

'*-----------------------------------------------------------------------------
'* Stream オブジェクトに書き込む文字列に、行区切り記号を追加するかどうかを表します。
'*
'*-----------------------------------------------------------------------------
Private Enum StreamWriteEnum
    adWriteChar = 0 '既定値。Stream オブジェクトに、Data パラメーターで指定したテキスト文字列を書き込みます。
    adWriteLine = 1 'Stream オブジェクトに、テキスト文字列と行区切り記号を書き込みます。LineSeparator プロパティが定義されていない場合は、実行時エラーを返します。
End Enum

'*-----------------------------------------------------------------------------
'* Connection 内のデータの編集、Record のオープン、または Record および Stream
'*  オブジェクトの Mode プロパティの値の指定に対する権限を表します。
'*-----------------------------------------------------------------------------
Private Enum ConnectModeEnum
    adModeRead = 1            '読み取り専用の権限を表します。
    adModeReadWrite = 3       '読み取り/書き込み両方の権限を表します。
    adModeRecursive = 4194304 '他の共有拒否値 (adModeShareDenyNone、 adModeShareDenyWrite、またはadModeShareDenyRead) と共に使用して、現在のレコードのすべてのサブレコードに共有制限を伝達します。 Record に子がない場合は機能しません。adModeShareDenyNone のみと組み合わせて使用すると、実行時エラーが発生します。 ただし、その他の値と組み合わせた場合は adModeShareDenyNone と組み合わせて使用できます。
    adModeShareDenyNone = 16  '権限の種類に関係なく、他のユーザーが接続を開けるようにします。他のユーザーに対して、読み取りと書き込みの両方のアクセスを許可します。
    adModeShareDenyRead = 4   '他のユーザーが読み取り権限で接続を開くのを禁止します。
    adModeShareDenyWrite = 8  '他のユーザーが書き込み権限で接続を開くのを禁止します。
    adModeShareExclusive = 12 '他のユーザーが接続を開くのを禁止します。
    adModeUnknown = 0         '既定値。権限が設定されていないか、権限を判定できないことを示します。
    adModeWrite = 2           '書き込み専用の権限を示します。
End Enum

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mFso As Object
Private mFile As Object
Private mAdoStream As Object
Private mCharSet As String
Private mLineSeparator As LineSeparatorsEnum
Private mHeaderExists As Boolean
Private mFilePath As String
Private mValidFormat As Boolean
Private mBOM As Boolean

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
End Property

'*-----------------------------------------------------------------------------
'* Charset プロパティ
'*-----------------------------------------------------------------------------
Property Get charset() As String
    charset = mCharSet
End Property
Property Let charset(arg As String)
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
'* HeaderExists プロパティ
'*-----------------------------------------------------------------------------
Property Get HeaderExists() As Boolean
    HeaderExists = mHeaderExists
End Property
Property Let HeaderExists(arg As Boolean)
    mHeaderExists = arg
End Property

'*-----------------------------------------------------------------------------
'* BOM プロパティ
'*-----------------------------------------------------------------------------
Property Get BOM() As Boolean
    BOM = mBOM
End Property
Property Let BOM(arg As Boolean)
    mBOM = arg
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
    mLineSeparator = LineSeparatorsEnum.adLF
    mHeaderExists = False
    mValidFormat = True
    mBOM = False
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
'* [詳  細] ファイルが存在するかどうか判定する｡
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
'* [概  要] ShowCsvSaveFileDialog メソッド
'* [詳  細] CSVファイルを保存するダイアログを表示する｡
'*
'* @param initialFile 初期表示用ファイルパス。省略可能。
'* @return 保存ファイルのパス
'*
'******************************************************************************
Public Function ShowCsvSaveFileDialog(Optional ByVal initialFile As String) As String
    Dim fileName As Variant
    fileName = Application.GetSaveAsFilename(initialFile, "CSVファイル,*.csv")
    If fileName = False Then
        Exit Function
    End If
    mFilePath = CStr(fileName)
    ShowCsvSaveFileDialog = CStr(fileName)
End Function

'******************************************************************************
'* [概  要] テキストファイル書込処理。
'* [詳  細] テキストファイルに配列データを行単位で書込を行う。
'*
'* @param vArray 書込配列データ（Variant）
'******************************************************************************
Public Sub WriteTextFileFromVArray(vArray As Variant)
    Dim tmp
    With mAdoStream
        .Mode = adModeReadWrite
        .Type = adTypeText
        .charset = mCharSet
        .LineSeparator = mLineSeparator
        .Open
        For Each tmp In vArray
            .WriteText CStr(tmp), adWriteLine
        Next
        If UCase(mCharSet) = "UTF-8" And Not mBOM Then
            Call ExcludeBom
        End If
        .SaveToFile mFilePath, adSaveCreateOverWrite 'ファイル上書指定
        .Close
    End With
End Sub

'******************************************************************************
'* [概  要] テキストファイル書込処理。
'* [詳  細] テキストファイルにRecordSetデータを行単位で書込を行う。
'*
'* @param records 書込データ（RecordFormat）
'******************************************************************************
Public Sub WriteTextFileFromRecordSet(records As RecordFormat)
    With mAdoStream
        .Mode = adModeReadWrite
        .Type = adTypeText
        .charset = mCharSet
        .LineSeparator = mLineSeparator
        .Open
        Dim recText As String
        'ヘッダ行出力
        If mHeaderExists Then
            recText = EditHeader(records.RecordDefinition)
            .WriteText recText, adWriteLine
        End If
        'データ行出力
        Dim record As Collection
        For Each record In records.RecordSet
            recText = EditRecord(record)
            .WriteText recText, adWriteLine
        Next
        If UCase(mCharSet) = "UTF-8" And Not mBOM Then
            Call ExcludeBom
        End If
        .SaveToFile mFilePath, adSaveCreateOverWrite 'ファイル上書指定
        .Close
    End With
End Sub

'******************************************************************************
'* [概  要] ヘッダ編集処理。
'* [詳  細] レコード情報をCSV文字列（ヘッダ）に編集する。
'*
'* @param record レコード情報（Collection）
'* @return CSV文字列（ヘッダ）
'******************************************************************************
Private Function EditHeader(record As Collection) As String
    Dim itm As Item, ret As String, first As Boolean
    first = True
    For Each itm In record
        If itm.OutputTarget Then
            If first Then
                ret = EditItem(itm.Name)
                first = False
            Else
                ret = ret & "," & EditItem(itm.Name)
            End If
        End If
    Next
    EditHeader = ret
End Function

'******************************************************************************
'* [概  要] レコード編集処理。
'* [詳  細] レコード情報をCSV文字列に編集する。
'*
'* @param record レコード情報（Collection）
'* @return CSV文字列
'******************************************************************************
Private Function EditRecord(record As Collection) As String
    Dim sInfo As SettingInfo: Set sInfo = GetSettingInfo()
           
    Dim itm As Item, ret As String, val As String, first As Boolean
    first = True
    For Each itm In record
        If itm.OutputTarget Then
            If itm.OutputEditKind = "マスタ変換（Code→Value）" Then
                val = sInfo.GetMasterValueByCode(itm.OutputEditValue, itm.value)
            ElseIf itm.OutputEditKind = "マスタ変換（Value→Code）" Then
                val = sInfo.GetMasterCodeByValue(itm.OutputEditValue, itm.value)
            Else
                val = itm.value
            End If
        
            If first Then
                ret = EditItem(val)
                first = False
            Else
                ret = ret & "," & EditItem(val)
            End If
        End If
    Next
    EditRecord = ret
End Function

'******************************************************************************
'* [概  要] 項目値編集処理。
'* [詳  細] 項目値から改行コードを除外し、ダブルクォートをエスケープする。
'*
'* @param val 項目値
'* @return 編集後項目値
'******************************************************************************
Private Function EditItem(val As String) As String
    Dim ret As String
    ret = Replace(val, vbLf, "")
    ret = Replace(ret, vbCr, "")
    ret = Replace(ret, """", """""")
    ret = """" & ret & """"
    EditItem = ret
End Function

'******************************************************************************
'* [概  要] BOM除外処理。
'* [詳  細] UTF-8テキストのBOMを除外する。
'* [備　考] 呼び出す前にStreamがOpenされていること。
'*
'* @return BOM除外データ
'******************************************************************************
Private Function ExcludeBom() As Byte()
    With mAdoStream
        .Position = 0
        .Type = adTypeBinary
        .Position = 3 'BOM飛ばし
        ExcludeBom = .Read
        .Position = 0
        .Write ExcludeBom
        .SetEOS
    End With
End Function

