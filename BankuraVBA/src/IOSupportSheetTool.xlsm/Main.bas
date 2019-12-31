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
'* WinowsAPI関数定義
'******************************************************************************
#If VBA7 Then
    'プログラムを任意の時間だけ待機させるAPI関数
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
    'イベントキュー待機中のイベントチェックAPI
    Public Declare PtrSafe Function GetInputState Lib "user32" () As Long
#Else
    'プログラムを任意の時間だけ待機させるAPI関数
    Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    'イベントキュー待機中のイベントチェックAPI
    Public Declare Function GetInputState Lib "user32" () As Long
#End If

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

'******************************************************************************
'* Enum定義
'******************************************************************************
'なし

'******************************************************************************
'* 変数定義
'******************************************************************************
Private mDisplayAlerts As Boolean
Private mScreenUpdating As Boolean
Private mCalculation As Long
Private mEnableEvents As Boolean
Private mRegExp As Object
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
Private dsItemCount As Long

Private mTime As Variant

'******************************************************************************
'* 関数定義
'******************************************************************************

'******************************************************************************
'* [概  要] 初期化処理。
'* [詳  細] 初期化の処理を行う。
'*
'******************************************************************************
Public Sub Init()
    Call GetSettingInfo
    Call SaveApplicationProperties
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
    Call Init

    Dim fReader As FileReader
    Set fReader = New FileReader

    'ダイアログ表示
    fReader.ShowCsvFileDialog
    
    'ファイルが選択されていれば読込を実行
    If fReader.FileExists Then
        Call StartProcess
        
        '入力CSV定義情報読込
        Dim rf As RecordFormat: Set rf = New RecordFormat
        Call rf.GetItemDataFromSheet(ThisWorkbook.Sheets(INPUTCSV_SHEET_NAME), csStartRow, csStartCol, csItemCount)
        
        'CSVファイル読込
        fReader.HeaderExists = True
        Dim vArr: vArr = fReader.ReadTextFileToVArray
        
        If IsEmpty(vArr) Then
            If fReader.ValidFormat Then
                Call Err.Raise(9999, "CSVファイル読込処理", "読込ファイルにデータがありません。")
            Else
                Call Err.Raise(9999, "CSVファイル読込処理", "読込ファイルのフォーマットが不正です。")
            End If
        End If
        
        'デバッグ出力
        'Call PrintVariantArray(vArr)
        
        'データ検証
        If Not rf.Validate(vArr) Then
            Call Err.Raise(9999, "CSVファイル読込処理", "読込ファイルのフォーマットが不正です。")
        End If
        
        'フォーム項目定義情報読込
        Dim formRecDef As RecordFormat: Set formRecDef = New RecordFormat
        Call formRecDef.GetItemDataFromSheet(ThisWorkbook.Sheets(FORM_SHEET_NAME), fmStartRow, fmStartCol, fmItemCount)
        
        'フォーム項目定義に基づきCSVデータをシート出力用フォームデータに変換
        Dim vFormArr: vFormArr = formRecDef.GetFormVariantData(vArr)
        Call CheckEvents

        'シートをクリア
        Dim mysheet As Worksheet
        Set mysheet = ThisWorkbook.Sheets(TOOL_SHEET_NAME)
        Call ClearActualUsedRangeFromSheet(mysheet, dsStartRow, dsKoubanCol, dsItemCount)
        Call DeleteNoUsedRange(mysheet, dsStartRow)
        Call CheckEvents
        
        'シートに出力
        Call InjectVariantArrayToCells(mysheet, vFormArr, dsStartRow, dsStartCol)
        Call CheckEvents
        Call InjectNumbersToIndexCells(mysheet, dsStartRow, dsKoubanCol, UBound(vFormArr, 1) - LBound(vFormArr, 1) + 1)
        Call CheckEvents

        Call EndProcess
        
        mysheet.Cells(dsStartRow, dsStartCol).Select
        MsgBox "CSVファイルの読込が完了しました｡", vbOKOnly + vbInformation, TOOL_NAME
    End If

    Exit Sub
ErrorHandler:
    Call EndProcess
    Call ErrorProcess
End Sub

'******************************************************************************
'* [概  要] ファイル出力ボタン押下時処理。
'* [詳  細] データシートを読み込みCSVファイルに出力する。
'*
'******************************************************************************
Public Sub OutputFileButton_Click()
    On Error GoTo ErrorHandler
    Call Init
    
    Dim mysheet As Worksheet
    Set mysheet = ThisWorkbook.Sheets(TOOL_SHEET_NAME)
    
    Dim fWriter As FileWriter
    Set fWriter = New FileWriter
    
    '項目データ読込・データ検証
    Dim rf As RecordFormat: Set rf = CheckDataSheet(mysheet)
    If rf Is Nothing Then
        Exit Sub
    End If

    'デバッグ出力
    'Call PrintRecordSet(rf)

    '出力先選択ダイアログ表示
    If fWriter.ShowCsvSaveFileDialog <> "" Then
        Dim ret As Long: ret = vbYes
        If fWriter.FileExists Then
            ret = MsgBox("既にファイルが存在します。" & vbCrLf & "上書きしてよろしいですか。", vbYesNo + vbQuestion, TOOL_NAME)
        End If
    
        If ret = vbYes Then
            Call StartProcess
            
            'ファイル出力
            fWriter.HeaderExists = True
            Call fWriter.WriteTextFileFromRecordSet(rf)

            Call EndProcess
            
            MsgBox "CSVファイルの出力が完了しました｡", vbOKOnly + vbInformation, TOOL_NAME
        End If
    End If

    Exit Sub
ErrorHandler:
    Call EndProcess
    Call ErrorProcess
End Sub

'******************************************************************************
'* [概  要] チェックボタン押下時処理。
'* [詳  細] データのチェックを行う。
'*
'******************************************************************************
Sub CheckButton_CLick()
    On Error GoTo ErrorHandler
    Call Init
    
    Dim mysheet As Worksheet
    Set mysheet = ThisWorkbook.Sheets(TOOL_SHEET_NAME)
    
    '項目データ読込・データ検証
    Dim rf As RecordFormat: Set rf = CheckDataSheet(mysheet)
    If rf Is Nothing Then
        Exit Sub
    End If

    'メッセージ表示
    mysheet.Cells(dsStartRow, dsStartCol).Select
    MsgBox "チェックが完了しました｡" + vbNewLine + "問題ありません。", vbOKOnly + vbInformation, TOOL_NAME
    
    Exit Sub

ErrorHandler:
    Call EndProcess
    Call ErrorProcess
End Sub

'******************************************************************************
'* [概  要] クリアボタン押下時処理。
'* [詳  細] データシートをクリアする。
'*
'******************************************************************************
Public Sub ClearButton_Click()
    On Error GoTo ErrorHandler
    Call Init

    'シートをクリア
    Call StartProcess
    Dim mysheet As Worksheet
    Set mysheet = ThisWorkbook.Sheets(TOOL_SHEET_NAME)
    Call ClearActualUsedRangeFromSheet(mysheet, dsStartRow, dsKoubanCol, dsItemCount)
    Call DeleteNoUsedRange(mysheet, dsStartRow)
    Call EndProcess
    mysheet.Cells(dsStartRow, dsStartCol).Select

    Exit Sub
ErrorHandler:
    Call EndProcess
    Call ErrorProcess
End Sub

'******************************************************************************
'* [概  要] 使い方ボタン押下時処理。
'* [詳  細] helpシートへ移動する。
'*
'******************************************************************************
Sub GotoHelpButton_Click()
    On Error GoTo ErrorHandler

    Call GotoSheet(HELP_SHEET_NAME)
    Exit Sub

ErrorHandler:
    Call ErrorProcess
End Sub

'******************************************************************************
'* [概  要] 戻るボタン押下時処理。
'* [詳  細] データシートに戻る。
'*
'******************************************************************************
Sub ReturnFromHelpButton_Click()
    On Error GoTo ErrorHandler
    Call GotoSheet(TOOL_SHEET_NAME)
    
    Exit Sub

ErrorHandler:
    Call ErrorProcess
End Sub

'******************************************************************************
'* [概  要] 設定シート「更新」ボタン押下時処理。
'* [詳  細] 設定情報を更新する。
'*
'******************************************************************************
Sub UpdateSettingButton_Click()
    On Error GoTo ErrorHandler
    Call Init
    Set mSettingInfo = New SettingInfo
    
    Exit Sub

ErrorHandler:
    Call ErrorProcess
End Sub

'******************************************************************************
'* [概  要] フォームデータ取得・検証処理。
'* [詳  細] フォーム（シート）から項目データを取得し項目定義情報をもとに検証を行う。
'*
'* @param mysheet ワークシート
'* @return レコードデータ情報
'*
'******************************************************************************
Function CheckDataSheet(mysheet As Worksheet) As RecordFormat
    Call Init
    Call StartProcess
    
    '項目定義情報読込
    Dim rf As RecordFormat: Set rf = New RecordFormat
    Call rf.GetItemDataFromSheet(ThisWorkbook.Sheets(FORM_SHEET_NAME), fmStartRow, fmStartCol, fmItemCount)

    '項目データ読込
    Dim readOk As Boolean: readOk = rf.GetRecordDataFromSheet(mysheet, dsStartRow, dsStartCol, rf.ColumnCount, dsKoubanCol)
    
    '項番振り直し
    Call ClearActualUsedRangeFromSheet(mysheet, dsStartRow, dsKoubanCol, 1)
    Call InjectNumbersToIndexCells(mysheet, dsStartRow, dsKoubanCol, rf.DataRowCount)
    
    'データ検証
    If Not readOk Then
        Call EndProcess
        mysheet.Cells(6 + rf.ErrRowNo - 1, 3 + rf.ErrColNo - 1).Select
        MsgBox rf.ErrMessage, vbOKOnly + vbExclamation, TOOL_NAME
        Set CheckDataSheet = Nothing
        Exit Function
    End If
    Set CheckDataSheet = rf
    
    Call EndProcess
    Exit Function

End Function

'******************************************************************************
'* [概  要] シート貼り付け処理。
'* [詳  細] Variant配列データをシートに出力する。
'*
'* @param dataSheet ワークシート
'* @param vArray Variant配列データ
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'*
'******************************************************************************
Private Sub InjectVariantArrayToCells(ByVal dataSheet As Worksheet, ByVal vArray, lStartRow As Long, lStartCol As Long)
    dataSheet.Cells(lStartRow, lStartCol).Resize(UBound(vArray, 1) + 1, UBound(vArray, 2) + 1).Value = vArray
End Sub

'******************************************************************************
'* [概  要] 項番設定処理。
'* [詳  細] 項番に連番を出力する。
'*
'* @param dataSheet ワークシート
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param rowNum 番号数
'*
'******************************************************************************
Private Sub InjectNumbersToIndexCells(ByVal dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, rownum As Long)
    With dataSheet
        .Cells(lStartRow, lStartCol) = 1
        If rownum > 1 Then
            .Cells(lStartRow, lStartCol).AutoFill _
              Destination:=Range(.Cells(lStartRow, lStartCol), .Cells(lStartRow + rownum - 1, lStartCol)), Type:=xlLinearTrend
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] エラー処理。
'* [詳  細] エラー発生時の処理を行う。
'*
'******************************************************************************
Public Sub ErrorProcess()
    Debug.Print "エラー発生 Number: " & Err.Number & " Source: " & Err.Source & " Description: " & Err.Description
    
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
Public Sub StartProcess()
    Call SaveApplicationProperties
    
    'シート保護解除
    Call UnprotectSheet
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

'******************************************************************************
'* [概  要] 終了処理。
'* [詳  細] 処理のスピード向上のため変更したExcelの設定を元に戻す。
'*
'******************************************************************************
Public Sub EndProcess()
    With Application
        .DisplayAlerts = mDisplayAlerts
        .ScreenUpdating = mScreenUpdating
        .Calculation = mCalculation
        .EnableEvents = mEnableEvents
        .StatusBar = False
    End With
    
    'シート保護
    Call ProtectSheet
End Sub

'******************************************************************************
'* [概  要] シート保護解除処理。
'* [詳  細] シートの保護を解除する。
'*
'******************************************************************************
Public Sub UnprotectSheet()
    'シート保護解除
    If TOOL_PASSWORD = "" Then
        ThisWorkbook.Sheets(TOOL_SHEET_NAME).Unprotect
    Else
        ThisWorkbook.Sheets(TOOL_SHEET_NAME).Unprotect Password:=TOOL_PASSWORD
    End If
End Sub

'******************************************************************************
'* [概  要] シート保護処理。
'* [詳  細] シートの保護をする。
'*
'******************************************************************************
Public Sub ProtectSheet()
    With ThisWorkbook.Sheets(TOOL_SHEET_NAME)
        .EnableOutlining = True  'アウトライン有効
        .EnableAutoFilter = True 'オートフィルタ有効
        
        'シート保護
        If TOOL_PASSWORD = "" Then
            .Protect Contents:=True, UserInterfaceOnly:=True
        Else
            .Protect Contents:=True, UserInterfaceOnly:=True, Password:=TOOL_PASSWORD
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] Application設定退避処理。
'* [詳  細] Applicationの設定をメンバ変数に退避する。
'*
'******************************************************************************
Public Sub SaveApplicationProperties()
    With Application
        mDisplayAlerts = .DisplayAlerts
        mScreenUpdating = .ScreenUpdating
        mCalculation = .Calculation
        mEnableEvents = .EnableEvents
    End With
End Sub

'******************************************************************************
'* [概  要] 正規表現オブジェクト取得処理。
'* [詳  細] 正規表現オブジェクトを取得する。未生成の場合生成する。
'*
'******************************************************************************
Public Function GetRegExp() As Object
    If mRegExp Is Nothing Then
        Set mRegExp = CreateObject("VBScript.RegExp")
    End If
    Set GetRegExp = mRegExp
End Function

'******************************************************************************
'* [概  要] 設定情報オブジェクト取得処理。
'* [詳  細] 設定情報オブジェクトを取得する。未生成の場合生成する。
'*
'******************************************************************************
Public Function GetSettingInfo() As SettingInfo
    If mSettingInfo Is Nothing Then
        Set mSettingInfo = New SettingInfo
    End If
    Set GetSettingInfo = mSettingInfo
End Function


'******************************************************************************
'* [概  要] GotoSheet
'* [詳  細] アクティブなブックの指定したシート・アドレスへ移動する。
'*
'* @param sheetName 移動先シート名
'* @param strAddr 移動先セルのアドレス
'******************************************************************************
Public Sub GotoSheet(SheetName As String, Optional strAddr As String = "A1")
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(SheetName).Select
    ThisWorkbook.Worksheets(SheetName).Range(strAddr).Activate
End Sub


'******************************************************************************
'* [概  要] 表情報取得処理。
'* [詳  細] worksheetの表から情報を取得し、Variant配列を返却します｡
'*
'* @param dataSheet ワークシート
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param itemCount 項目列数
'*
'******************************************************************************
Public Function GetVariantDataFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long)
    Dim lMaxRow As Long: lMaxRow = dataSheet.Cells(Rows.Count, lStartCol).End(xlUp).row
    Dim lMaxCol As Long
    If colCount = 0 Then
        lMaxCol = Cells(lStartRow, Columns.Count).End(xlToLeft).Column
    Else
        lMaxCol = lStartCol + colCount - 1
    End If
    
    'レコードが存在しない場合
    If lMaxRow < lStartRow Or lMaxCol < lStartCol Then
        GetVariantDataFromSheet = Empty
        Exit Function
    End If
    
    Dim vArr: vArr = dataSheet.Range(dataSheet.Cells(lStartRow, lStartCol), dataSheet.Cells(lMaxRow, lMaxCol))
    
    GetVariantDataFromSheet = vArr
End Function


'******************************************************************************
'* [概  要] 使用セル範囲クリア処理。
'* [詳  細] worksheetのデータ表の使用セル範囲をクリアします｡
'*
'* @param dataSheet data表ワークシート
'* @param lStartRow data表データ開始行番号
'* @param lStartCol data表データ開始列番号
'* @param itemCount 項目列数
'* @param ignoreColnum 走査対象外の列番号
'*
'******************************************************************************
Public Sub ClearActualUsedRangeFromSheet(dataSheet As Worksheet, _
                                         lStartRow As Long, _
                                         lStartCol As Long, _
                                         Optional colCount As Long, _
                                         Optional ignoreColnum As Long)
    Dim rng As Range
    Set rng = GetActualUsedRangeFromSheet(dataSheet, lStartRow, lStartCol, colCount, ignoreColnum)
    If rng Is Nothing Then
        Exit Sub
    End If
    rng.ClearContents
End Sub

'******************************************************************************
'* [概  要] 未使用範囲行削除処理。
'* [詳  細] worksheetのデータ表の未使用範囲行を削除します（UsedRangeを縮小）｡
'*
'* @param dataSheet data表ワークシート
'* @param lStartRow data表データ開始行番号
'*
'******************************************************************************
Public Sub DeleteNoUsedRange(dataSheet As Worksheet, lStartRow As Long)
    Dim delStartRow As Long
    Dim delEndRow As Long
    
    Dim rng As Range
    Set rng = GetActualUsedRangeFromSheet(dataSheet, lStartRow, 1)
    If rng Is Nothing Then
        delStartRow = lStartRow
    Else
        delStartRow = rng.Item(rng.Count).row + 1
    End If
    delEndRow = dataSheet.UsedRange.Item(dataSheet.UsedRange.Count).row
    
    If delStartRow > delEndRow Then
        Exit Sub
    End If
    With dataSheet
        .Range(.Rows(delStartRow), .Rows(delEndRow)).Delete
    End With
End Sub

'******************************************************************************
'* [概  要] 使用セル範囲取得処理。
'* [詳  細] worksheetのデータ表の使用セル範囲を取得する｡
'*
'* @param dataSheet data表ワークシート
'* @param lStartRow data表データ開始行番号
'* @param lStartCol data表データ開始列番号
'* @param itemCount 項目列数
'* @param ignoreColnum 走査対象外の列番号
'* @return 使用セル範囲
'*
'******************************************************************************
Public Function GetActualUsedRangeFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long, Optional ignoreColnum As Long) As Range
    Dim lMaxRow As Long: lMaxRow = GetFinalRow(dataSheet, ignoreColnum)
    Dim lMaxCol As Long
    If colCount = 0 Then
        lMaxCol = GetFinalCol(dataSheet)
    Else
        lMaxCol = lStartCol + colCount - 1
    End If

    'レコードが存在しない場合
    If lMaxRow < lStartRow Or lMaxCol < lStartCol Then
        Set GetActualUsedRangeFromSheet = Nothing
        Exit Function
    End If
    
    Set GetActualUsedRangeFromSheet = dataSheet.Range(dataSheet.Cells(lStartRow, lStartCol), dataSheet.Cells(lMaxRow, lMaxCol))
End Function

'******************************************************************************
'* [概  要] 最終行取得処理。
'* [詳  細] WorksheetのUsedRangeを下から走査し、最終行番号を取得する｡
'*
'* @param dataSheet ワークシート
'* @param ignoreColnum 走査対象外の列番号
'* @return 最終行番号
'*
'******************************************************************************
Public Function GetFinalRow(ByVal dataSheet As Worksheet, Optional ignoreColnum As Long) As Long
    Dim ret As Long
    Dim i As Long, cnta As Long
    With dataSheet.UsedRange
        For i = .Rows.Count To 1 Step -1
            cnta = WorksheetFunction.counta(.Rows(i))
            If cnta > 0 Then
                If cnta <> 1 Then
                    ret = i
                    Exit For
                Else
                    If ignoreColnum > 0 Then
                        If .Cells(i, ignoreColnum) = "" Then
                            ret = i
                            Exit For
                        End If
                    Else
                        ret = i
                        Exit For
                    End If
                End If
            End If
        Next
        If ret > 0 Then
            ret = ret + .row - 1
        End If
    End With
    GetFinalRow = ret
End Function

'******************************************************************************
'* [概  要] 最終列取得処理。
'* [詳  細] WorksheetのUsedRangeを右から走査し、最終列番号を取得する｡
'*
'* @param dataSheet ワークシート
'* @return 最終列番号
'*
'******************************************************************************
Public Function GetFinalCol(ByVal dataSheet As Worksheet) As Long
    Dim ret As Long
    Dim i As Long
    With dataSheet.UsedRange
        For i = .Columns.Count To 1 Step -1
            If WorksheetFunction.counta(.Columns(i)) > 0 Then
                ret = i
                Exit For
            End If
        Next
        If ret > 0 Then
            ret = ret + .Column - 1
        End If
    End With
    GetFinalCol = ret
End Function


'******************************************************************************
'* [概  要] Variant配列デバッグ出力処理
'* [詳  細] Variant配列の内容をイミディエイトウィンドウに出力する。
'*
'* @param vArr Variant配列
'******************************************************************************
Private Sub PrintVariantArray(vArr)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(vArr, 1) To UBound(vArr, 1)
        For j = LBound(vArr, 2) To UBound(vArr, 2)
            tmp = tmp & " | " & vArr(i, j)
        Next
        Debug.Print tmp
        tmp = ""
    Next
End Sub

'******************************************************************************
'* [概  要] Variant配列デバッグ出力処理
'* [詳  細] Variant配列の内容をイミディエイトウィンドウに出力する。
'*
'* @param vArr Variant配列
'******************************************************************************
Private Sub PrintRecordSet(rf As RecordFormat)
    Dim record As Collection, itm As Item, tmp As String
    For Each record In rf.RecordSet
        For Each itm In record
            tmp = tmp & " | " & itm.Value
        Next
        Debug.Print tmp
        tmp = ""
    Next
End Sub

Public Sub CheckEvents()
    If GetInputState() Or (DateDiff("s", mTime, Time) > 3) Then
        DoEvents
        mTime = Time
    End If
End Sub
