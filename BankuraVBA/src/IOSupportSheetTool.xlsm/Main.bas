Attribute VB_Name = "Main"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ツールメイン処理モジュール
'* [詳  細] ツールのメイン処理を行うモジュール。
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 定数定義
'******************************************************************************
'ツール名
Public Const TOOL_NAME As String = "入力支援ツール"
'ツールパスワード
Public Const TOOL_PASSWORD As String = "1234"
'ツールシート名
Public Const TOOL_SHEET_NAME As String = "data"
'入力CSVファイル設定シート名
Public Const INPUTCSV_SHEET_NAME As String = "inputcsv_setting"
'ツールフォーム（データシート）設定シート名
Public Const FORM_SHEET_NAME As String = "form_setting"
'HELP（使い方）シート名
Public Const HELP_SHEET_NAME As String = "help"
'設定シート情報
Private Const SETTING_SHEET_NAME As String = "setting"
Private Const SETTING_SH_START_ROW As Long = 4
Private Const SETTING_SH_START_COL As Long = 4

'******************************************************************************
'* Enum定義
'******************************************************************************
'なし

'******************************************************************************
'* 変数定義
'******************************************************************************
Private mSettingInfo As SettingInfo
Private csStartRow As Long
Private csStartCol As Long
Private csItemCount As Long
Private fmStartRow As Long
Private fmStartCol As Long
Private fmItemCount As Long
Public dsStartRow As Long
Public dsStartCol As Long
Public dsKoubanCol As Long
Public dsItemCount As Long

'******************************************************************************
'* 関数定義
'******************************************************************************

'******************************************************************************
'* [概  要] 初期化処理。
'* [詳  細] 初期化の処理を行う。
'*
'******************************************************************************
Public Sub MyInit()
    Base.SettingSheetName = SETTING_SHEET_NAME
    Base.SettingSheetStartRow = SETTING_SH_START_ROW
    Base.SettingSheetStartCol = SETTING_SH_START_COL
    
    Set mSettingInfo = Base.GetSettingInfo
    Call Base.SaveApplicationProperties
    csStartRow = mSettingInfo.GetSettingValue("InputCsvSettingStartRowNo")
    csStartCol = mSettingInfo.GetSettingValue("InputCsvSettingStartColNo")
    csItemCount = mSettingInfo.GetSettingValue("InputCsvSettingItemCount")
    fmStartRow = mSettingInfo.GetSettingValue("FormSettingStartRowNo")
    fmStartCol = mSettingInfo.GetSettingValue("FormSettingStartColNo")
    fmItemCount = mSettingInfo.GetSettingValue("FormSettingItemCount")
    dsStartRow = mSettingInfo.GetSettingValue("DataSheetStartRowNo")
    dsStartCol = mSettingInfo.GetSettingValue("DataSheetStartColNo")
    dsKoubanCol = mSettingInfo.GetSettingValue("DataSheetKoubanColNo")
    dsItemCount = mSettingInfo.GetSettingValue("DataSheetItemCount")
End Sub

'******************************************************************************
'* [概  要] ファイル読込ボタン押下時処理。
'* [詳  細] CSVファイルを読み込みデータシートに出力する。
'*
'******************************************************************************
Public Sub ReadFileButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, TOOL_SHEET_NAME)
        
    Dim fReader As CsvFileReader
    Set fReader = New CsvFileReader
    fReader.QuotExists = True
    
    'Dim myReporter As SBProgressReporter: Set myReporter = New SBProgressReporter
    Dim myReporter As FormProgressReporter: Set myReporter = New FormProgressReporter
    myReporter.BaseMessage = "CSV読込処理中"
    Set fReader.ProgressReporter = myReporter
    
    ' ダイアログ表示
    fReader.ShowCsvFileDialog
    
    ' ファイルが選択されていれば読込を実行
    If fReader.FileExists Then
        Call MyStartProcess
        
        ' 入力CSV定義情報読込
        Dim rf As RecordFormat
        Set rf = GetDefinedRecordFormatFromSheet(INPUTCSV_SHEET_NAME, csStartRow, csStartCol, csItemCount)

        ' CSVファイル読込
        fReader.HeaderExists = True
        Dim csvData As Array2DEx: Set csvData = fReader.Read
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        DoEvents
        
        If csvData.IsEmptyArray() Then
            If fReader.ValidFormat Then
                Call Err.Raise(9999, "CSVファイル読込処理", "読込ファイルにデータがありません。")
            Else
                Call Err.Raise(9999, "CSVファイル読込処理", "読込ファイルのフォーマットが不正です。")
            End If
        End If

        ' データ検証
        myReporter.BaseMessage = "CSVデータ検証中"
        Set rf.ProgressReporter = myReporter
        If Not rf.Validate(csvData) Then
            DebugUtils.Show rf.ErrMessage
            Call Err.Raise(9999, "CSVファイル読込処理", "読込ファイルのフォーマットが不正です。")
        End If
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        DoEvents
        
        ' フォーム項目定義情報読込
        Dim formRecDef As RecordFormat
        Set formRecDef = GetDefinedRecordFormatFromSheet(FORM_SHEET_NAME, fmStartRow, fmStartCol, fmItemCount)
        
        ' 項目定義に基づきCSVデータをシート出力用フォームデータに変換
        myReporter.BaseMessage = "シートデータ変換処理中"
        Set formRecDef.ProgressReporter = myReporter
        Dim formData As Array2DEx: Set formData = formRecDef.Convert(csvData)
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        DoEvents

        'シートをクリア
        Call mysheet.ClearActualUsedRange(dsStartRow, dsKoubanCol, dsItemCount)
        Call mysheet.DeleteNoUsedRange(dsStartRow)
        Call UXUtils.CheckEvents
        
        'シートに出力
        Call mysheet.ImportArray(formData, dsStartRow, dsStartCol)
        Call UXUtils.CheckEvents
        Call mysheet.NumbersToIndexCells(dsStartRow, dsKoubanCol, formData.RowLength)
        Call UXUtils.CheckEvents

        Call MyEndProcess
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        mysheet.Cells(dsStartRow, dsStartCol).Select
        MsgBox "CSVファイルの読込が完了しました｡", vbOKOnly + vbInformation, TOOL_NAME
    End If

    Exit Sub
ErrorHandler:
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [概  要] ファイル出力ボタン押下時処理。
'* [詳  細] データシートを読み込みCSVファイルに出力する。
'*
'******************************************************************************
Public Sub OutputFileButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit
    
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, TOOL_SHEET_NAME)
    
    ' 項目定義情報読込
    Dim rf As RecordFormat
    Set rf = GetDefinedRecordFormatFromSheet(FORM_SHEET_NAME, fmStartRow, fmStartCol, fmItemCount)
    
    ' 項目データ読込・データ検証
    Dim formData As Array2DEx: Set formData = CheckDataSheet(mysheet, rf)
    If formData Is Nothing Then
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        Exit Sub
    End If

    ' ファイル出力準備
    Dim fWriter As CsvFileWriter
    Set fWriter = New CsvFileWriter
    fWriter.EnclosureChar = """"
    fWriter.WillRemoveNewlineCode = True

    ' 出力先選択ダイアログ表示
    If fWriter.ShowCsvSaveFileDialog <> "" Then
        Dim ret As Long: ret = vbYes
        If fWriter.FileExists Then
            ret = MsgBox("既にファイルが存在します。" & vbCrLf & "上書きしてよろしいですか。", vbYesNo + vbQuestion, TOOL_NAME)
        End If

        If ret = vbYes Then
            Call MyStartProcess
            
            ' 項目データ編集
            Dim myReporter As FormProgressReporter: Set myReporter = New FormProgressReporter
            myReporter.BaseMessage = "項目データ編集中"
            Set rf.ProgressReporter = myReporter
            Dim editedFormData As Array2DEx
            Set editedFormData = rf.Convert(formData, True)
            
            ' ファイル出力
            'Dim myReporter As SBProgressReporter: Set myReporter = New SBProgressReporter
            myReporter.BaseMessage = "CSV出力処理中"
            Set fWriter.ProgressReporter = myReporter
            fWriter.HeaderExists = True
            Set fWriter.ItemNames = rf.ItemNames
            Call fWriter.WriteFile(editedFormData)

            Call MyEndProcess
            AppActivate ThisWorkbook.Name
            mysheet.Activate
            MsgBox "CSVファイルの出力が完了しました｡", vbOKOnly + vbInformation, TOOL_NAME
        End If
    End If

    Exit Sub
ErrorHandler:
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [概  要] 行削除ボタン押下時処理。
'* [詳  細] データシートを読み込みCSVファイルに出力する。
'*
'******************************************************************************
Public Sub DeleteRowButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, TOOL_SHEET_NAME)

    If ActiveWindow.RangeSelection.row >= dsStartRow Then
        Call MyStartProcess
        
        '選択範囲の行を削除
        ActiveWindow.RangeSelection.EntireRow.Delete
        
        '項番振り直し
        Dim lMaxRow As Long: lMaxRow = mysheet.GetFinalKeyRow(dsKoubanCol)
        Call mysheet.NumbersToIndexCells(dsStartRow, dsKoubanCol, lMaxRow - dsStartRow + 1)
        
        Call MyEndProcess
    End If
    Exit Sub

ErrorHandler:
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyErrorProcess
End Sub


'******************************************************************************
'* [概  要] チェックボタン押下時処理。
'* [詳  細] データのチェックを行う。
'*
'******************************************************************************
Sub CheckButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit
    
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, TOOL_SHEET_NAME)

    ' 項目定義情報読込
    Dim rf As RecordFormat
    Set rf = GetDefinedRecordFormatFromSheet(FORM_SHEET_NAME, fmStartRow, fmStartCol, fmItemCount)
    
    '項目データ読込・データ検証
    Dim formData As Array2DEx: Set formData = CheckDataSheet(mysheet, rf)
    If formData Is Nothing Then
        Exit Sub
    End If

    'メッセージ表示
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    mysheet.Cells(dsStartRow, dsStartCol).Select
    MsgBox "チェックが完了しました｡" + vbNewLine + "問題ありません。", vbOKOnly + vbInformation, TOOL_NAME
    
    Exit Sub

ErrorHandler:
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [概  要] クリアボタン押下時処理。
'* [詳  細] データシートをクリアする。
'*
'******************************************************************************
Public Sub ClearButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit

    Call MyStartProcess
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, TOOL_SHEET_NAME)
    
    'フィルタ選択解除
    If Not mysheet.AutoFilter Is Nothing Then
        If mysheet.FilterMode Then
            mysheet.ShowAllData
        End If
    End If
    
    'シートをクリア
    Call mysheet.ClearActualUsedRange(dsStartRow, dsKoubanCol, dsItemCount)
    Call mysheet.DeleteNoUsedRange(dsStartRow)
    
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    mysheet.Cells(dsStartRow, dsStartCol).Select

    Exit Sub
ErrorHandler:
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [概  要] 使い方ボタン押下時処理。
'* [詳  細] helpシートへ移動する。
'*
'******************************************************************************
Sub GotoHelpButton_Click()
    On Error GoTo ErrorHandler

    Call XlWorkSheetUtils.GotoSheet(HELP_SHEET_NAME)
    Exit Sub

ErrorHandler:
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [概  要] 戻るボタン押下時処理。
'* [詳  細] データシートに戻る。
'*
'******************************************************************************
Sub ReturnFromHelpButton_Click()
    On Error GoTo ErrorHandler
    Call XlWorkSheetUtils.GotoSheet(TOOL_SHEET_NAME)
    
    Exit Sub

ErrorHandler:
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [概  要] 設定シート「更新」ボタン押下時処理。
'* [詳  細] 設定情報を更新する。
'*
'******************************************************************************
Sub UpdateSettingButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit
    Set mSettingInfo = New SettingInfo
    
    Exit Sub

ErrorHandler:
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [概  要] フォームデータ取得・検証処理。
'* [詳  細] フォーム（シート）から項目データを取得し項目定義情報をもとに検証を行う。
'*
'* @param mysheet    ワークシート
'* @param rf         項目定義情報
'* @return Array2DEx レコードデータ情報
'*
'******************************************************************************
Function CheckDataSheet(mysheet As WorkSheetEx, rf As RecordFormat) As Array2DEx
    Call MyInit
    Call MyStartProcess
    
    ' 項目データ読込
    Dim formData As Array2DEx
    Set formData = mysheet.GetActualUsedRangeToArray2DEx(dsStartRow, dsStartCol, rf.ColumnCount, dsKoubanCol)
    If formData.IsEmptyArray Then
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        Call MyEndProcess
        MsgBox "データが入力されていません。", vbOKOnly + vbExclamation, TOOL_NAME
        Set CheckDataSheet = Nothing
        Exit Function
    End If

    ' 項番振り直し
    Call mysheet.ClearActualUsedRange(dsStartRow, dsKoubanCol, 1)
    Call mysheet.NumbersToIndexCells(dsStartRow, dsKoubanCol, formData.RowLength)
    
    ' データ検証
    Dim myReporter As FormProgressReporter: Set myReporter = New FormProgressReporter
    myReporter.BaseMessage = "データ検証中"
    Set rf.ProgressReporter = myReporter
    If Not rf.Validate(formData, True) Then
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        mysheet.Cells(dsStartRow + rf.ErrRowNo - 1, dsStartCol + rf.ErrColNo - 1).Select
        Call MyEndProcess
        MsgBox rf.ErrMessage, vbOKOnly + vbExclamation, TOOL_NAME
        Set CheckDataSheet = Nothing
        Exit Function
    End If
    Set CheckDataSheet = formData
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyEndProcess
    Exit Function
End Function

'******************************************************************************
'* [概  要] エラー処理。
'* [詳  細] エラー発生時の処理を行う。
'*
'******************************************************************************
Public Sub MyErrorProcess()
    Call ErrorProcess
    
    If Err.Number = 9999 Then
        MsgBox Err.Description, vbOKOnly + vbExclamation, TOOL_NAME
    ElseIf Err.Number = 3004 Then
        MsgBox "ファイルへ書き込めませんでした。" & vbNewLine & _
               "別プログラムでファイルを開いているなどの原因が考えられます。" & vbNewLine & _
                "ご確認ください。", vbOKOnly + vbExclamation, TOOL_NAME
    Else
        MsgBox "システムエラーが発生しました｡", vbOKOnly + vbCritical, TOOL_NAME
    End If
End Sub

'******************************************************************************
'* [概  要] 開始処理。
'* [詳  細] 処理のスピード向上のため、Excelの設定を変更する。
'*
'******************************************************************************
Public Sub MyStartProcess()
    Call StartProcess

    'シート保護解除
    Call XlWorkSheetUtils.UnprotectSheet(TOOL_SHEET_NAME, TOOL_PASSWORD)
End Sub

'******************************************************************************
'* [概  要] 終了処理。
'* [詳  細] 処理のスピード向上のため変更したExcelの設定を元に戻す。
'*
'******************************************************************************
Public Sub MyEndProcess()
    Call EndProcess
    
    'シート保護
    Call XlWorkSheetUtils.ProtectSheet(TOOL_SHEET_NAME, TOOL_PASSWORD)
End Sub

'******************************************************************************
'* [概  要] 項目表レコード定義情報取得・設定処理。
'* [詳  細] worksheetの項目表からレコード定義情報を取得する｡
'*
'* @param defSheetName 項目表ワークシート名
'* @param lStartRow 項目表データ開始行番号
'* @param lStartCol 項目表データ開始列番号
'* @param itemCount 項目列数
'* @return RecordFormat レコード定義情報
'*
'******************************************************************************
Public Function GetDefinedRecordFormatFromSheet(defSheetName As String, lStartRow As Long, lStartCol As Long, Optional colCount As Long) As RecordFormat
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, defSheetName)
    
    Dim vArr: vArr = mysheet.ExportArray(lStartRow, lStartCol, colCount)
    If IsEmpty(vArr) Then
        Set GetDefinedRecordFormatFromSheet = Nothing
        Exit Function
    End If
    Set GetDefinedRecordFormatFromSheet = GetDefinedRecordFormat(vArr)
End Function

'******************************************************************************
'* [概  要] レコード定義情報取得処理。
'* [詳  細] 項目定義情報（2次元配列）からレコード定義情報を作成・取得する｡
'*
'* @param vArr Variant型2次元配列（項目表データ）
'* @return RecordFormat レコード定義情報
'*
'******************************************************************************
Private Function GetDefinedRecordFormat(vArr) As RecordFormat
    Dim itmdefs As Collection: Set itmdefs = New Collection
    Dim i As Long
    For i = LBound(vArr, 1) To UBound(vArr, 1)
        itmdefs.Add DefineItem(vArr, i)
    Next
    Set GetDefinedRecordFormat = Core.Init(New RecordFormat, itmdefs)
End Function

'******************************************************************************
'* [概  要] 項目設定処理。
'* [詳  細] Itemに定義情報を設定する｡
'*
'* @param vArr Variant型2次元配列（項目表データ）
'* @param rownum 配列行（1次元添え字）
'* @return Item 定義済み項目
'*
'******************************************************************************
Private Function DefineItem(vArr, rowNum As Long) As Item
    Dim itm As Item
    Set itm = New Item
    itm.Name = vArr(rowNum, 1)
    If vArr(rowNum, 2) = "○" Then
        itm.required = True
    End If
    Select Case vArr(rowNum, 3)
        Case "半角"
            itm.Attr = AttributeEnum.attrHalf
        Case "半角英数"
            itm.Attr = AttributeEnum.attrHalfAlphaNumeric
        Case "半角英数記号"
            itm.Attr = AttributeEnum.attrHalfAlphaNumericSymbol
        Case "数値"
            itm.Attr = AttributeEnum.attrNumeric
        Case "全角カタカナ"
            itm.Attr = AttributeEnum.attrZenKatakana
        Case "全角ひらがな"
            itm.Attr = AttributeEnum.attrZenHiragana
        Case "日付"
            itm.Attr = AttributeEnum.attrDate
        Case "郵便番号"
            itm.Attr = AttributeEnum.attrZipCode
        Case "電話番号"
            itm.Attr = AttributeEnum.attrTelNo
        Case "メールアドレス"
            itm.Attr = AttributeEnum.attrMailAddress
        Case Else
            itm.Attr = AttributeEnum.attrString
    End Select
    Select Case vArr(rowNum, 4)
        Case "固定"
            itm.KindOfDigits = KindOfDigitsEnum.digitFixed
        Case "以内"
            itm.KindOfDigits = KindOfDigitsEnum.digitWithin
        Case "範囲"
            itm.KindOfDigits = KindOfDigitsEnum.digitRange
        Case Else
            itm.KindOfDigits = KindOfDigitsEnum.digitNone
    End Select
    If vArr(rowNum, 5) <> "" And VBA.IsNumeric(vArr(rowNum, 5)) Then
        itm.MinNumOfDigits = CLng(vArr(rowNum, 5))
    End If
    If vArr(rowNum, 6) <> "" And VBA.IsNumeric(vArr(rowNum, 6)) Then
        itm.MaxNumOfDigits = CLng(vArr(rowNum, 6))
    End If
    itm.Pattern = vArr(rowNum, 7)
    If UBound(vArr, 2) = 13 Then
        If vArr(rowNum, 8) <> "" And VBA.IsNumeric(vArr(rowNum, 8)) Then
            itm.InputColNo = vArr(rowNum, 8)
        End If
        Select Case vArr(rowNum, 9)
            Case "マスタ変換（Code→Value）"
                itm.InitValueKind = EditKindEnum.mstCodeToValue
            Case "マスタ変換（Value→Code）"
                itm.InitValueKind = EditKindEnum.mstValueToCode
            Case "デフォルト"
                itm.InitValueKind = EditKindEnum.useDefaultValue
            Case Else
                itm.InitValueKind = EditKindEnum.edtNone
        End Select
        itm.InitValue = vArr(rowNum, 10)
        If vArr(rowNum, 11) = "○" Then
            itm.OutputTarget = True
        End If
        Select Case vArr(rowNum, 12)
            Case "マスタ変換（Code→Value）"
                itm.OutputEditKind = EditKindEnum.mstCodeToValue
            Case "マスタ変換（Value→Code）"
                itm.OutputEditKind = EditKindEnum.mstValueToCode
            Case "デフォルト"
                itm.OutputEditKind = EditKindEnum.useDefaultValue
            Case Else
                itm.OutputEditKind = EditKindEnum.edtNone
        End Select
        itm.OutputEditValue = vArr(rowNum, 13)
    End If
    Set DefineItem = itm
End Function

